import os
import uuid
from datetime import datetime, timezone
from functools import wraps
from urllib.parse import urlparse

import boto3
import pg8000.dbapi as pg8000
from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, jsonify, session
)
from PIL import Image
from io import BytesIO

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration (all from environment variables — see .env.example)
# ---------------------------------------------------------------------------
DATABASE_URL = os.environ.get('DATABASE_URL', '')

R2_ACCOUNT_ID = os.environ.get('R2_ACCOUNT_ID', '')
R2_ACCESS_KEY_ID = os.environ.get('R2_ACCESS_KEY_ID', '')
R2_SECRET_ACCESS_KEY = os.environ.get('R2_SECRET_ACCESS_KEY', '')
R2_BUCKET_NAME = os.environ.get('R2_BUCKET_NAME', '')
R2_PUBLIC_URL = os.environ.get('R2_PUBLIC_URL', '').rstrip('/')

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'changeme123')
VIEWER_PASSWORD = os.environ.get('VIEWER_PASSWORD', '0000')
SECRET_KEY = os.environ.get('SECRET_KEY', 'change-this-secret-key-before-deploying')

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp'}

app = Flask(__name__)
app.secret_key = SECRET_KEY
# Generous cap: a product can now be uploaded with several photos in one go.
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024


# ---------------------------------------------------------------------------
# Database (Postgres / Supabase via pg8000 — pure Python, no compiler needed)
# ---------------------------------------------------------------------------
def get_db():
    if not DATABASE_URL:
        raise RuntimeError(
            'DATABASE_URL is not set. Add it to your .env file (see .env.example).'
        )
    parsed = urlparse(DATABASE_URL)
    return pg8000.connect(
        user=parsed.username,
        password=parsed.password,
        host=parsed.hostname,
        port=parsed.port or 5432,
        database=parsed.path.lstrip('/'),
        ssl_context=True,
    )


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS items (
            id TEXT PRIMARY KEY,
            name TEXT,
            category TEXT,
            image_key TEXT NOT NULL,
            thumb_key TEXT NOT NULL,
            date_added TIMESTAMP NOT NULL
        )
    ''')
    # One product can have many photos. items.image_key / thumb_key still hold
    # the "cover" photo (the first one) so the catalog grid needs no changes.
    cur.execute('''
        CREATE TABLE IF NOT EXISTS item_images (
            id TEXT PRIMARY KEY,
            item_id TEXT NOT NULL REFERENCES items(id) ON DELETE CASCADE,
            image_key TEXT NOT NULL,
            thumb_key TEXT NOT NULL,
            position INTEGER NOT NULL DEFAULT 0,
            date_added TIMESTAMP NOT NULL
        )
    ''')
    cur.execute('CREATE INDEX IF NOT EXISTS idx_item_images_item ON item_images (item_id)')
    # Backfill: give every existing product its current photo as the first image.
    cur.execute('''
        INSERT INTO item_images (id, item_id, image_key, thumb_key, position, date_added)
        SELECT 'legacy-' || i.id, i.id, i.image_key, i.thumb_key, 0, i.date_added
        FROM items i
        WHERE NOT EXISTS (SELECT 1 FROM item_images im WHERE im.item_id = i.id)
        ON CONFLICT (id) DO NOTHING
    ''')
    conn.commit()
    cur.close()
    conn.close()


def rows_to_dicts(cur, rows):
    cols = [c[0] for c in cur.description]
    return [dict(zip(cols, row)) for row in rows]


def query_items(q):
    conn = get_db()
    cur = conn.cursor()
    if q:
        like = f'%{q}%'
        cur.execute(
            'SELECT * FROM items WHERE id LIKE %s OR name LIKE %s OR category LIKE %s ORDER BY date_added DESC',
            (like, like, like)
        )
    else:
        cur.execute('SELECT * FROM items ORDER BY date_added DESC')
    items = rows_to_dicts(cur, cur.fetchall())
    cur.close()
    conn.close()
    return items


def get_item(item_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT * FROM items WHERE id = %s', (item_id,))
    rows = cur.fetchall()
    items = rows_to_dicts(cur, rows)
    cur.close()
    conn.close()
    return items[0] if items else None


def get_item_images(item_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        'SELECT * FROM item_images WHERE item_id = %s ORDER BY position, date_added',
        (item_id,)
    )
    images = rows_to_dicts(cur, cur.fetchall())
    cur.close()
    conn.close()
    return images


# ---------------------------------------------------------------------------
# Storage (Cloudflare R2, via the S3-compatible API)
# ---------------------------------------------------------------------------
def get_r2_client():
    if not (R2_ACCOUNT_ID and R2_ACCESS_KEY_ID and R2_SECRET_ACCESS_KEY):
        raise RuntimeError(
            'R2 credentials are not set. Add them to your .env file (see .env.example).'
        )
    return boto3.client(
        's3',
        endpoint_url=f'https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com',
        aws_access_key_id=R2_ACCESS_KEY_ID,
        aws_secret_access_key=R2_SECRET_ACCESS_KEY,
        region_name='auto',
    )


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def make_thumbnail_bytes(file_bytes, size=(500, 500)):
    with Image.open(BytesIO(file_bytes)) as img:
        img = img.convert('RGB')
        img.thumbnail(size)
        out = BytesIO()
        img.save(out, 'JPEG', quality=85)
        return out.getvalue()


def save_image(file):
    """Upload an image + thumbnail to R2. Returns (image_key, thumb_key)."""
    ext = file.filename.rsplit('.', 1)[1].lower()
    file_bytes = file.read()
    unique_name = uuid.uuid4().hex

    image_key = f'items/{unique_name}.{ext}'
    thumb_key = f'items/thumbs/{unique_name}.jpg'

    content_type = file.mimetype or f'image/{ext}'
    s3 = get_r2_client()
    s3.put_object(Bucket=R2_BUCKET_NAME, Key=image_key, Body=file_bytes, ContentType=content_type)

    try:
        thumb_bytes = make_thumbnail_bytes(file_bytes)
        s3.put_object(Bucket=R2_BUCKET_NAME, Key=thumb_key, Body=thumb_bytes, ContentType='image/jpeg')
    except Exception:
        thumb_key = image_key  # fallback: reuse original if thumbnailing fails

    return image_key, thumb_key


def delete_keys(image_key, thumb_key):
    """Best-effort removal of one image + its thumbnail from R2."""
    try:
        s3 = get_r2_client()
        for key in {image_key, thumb_key}:
            if key:
                s3.delete_object(Bucket=R2_BUCKET_NAME, Key=key)
    except Exception:
        pass  # don't block the DB change if R2 cleanup fails


def delete_all_item_images(images):
    for img in images:
        delete_keys(img['image_key'], img['thumb_key'])


def r2_url(key):
    return f'{R2_PUBLIC_URL}/{key}'


# ---------------------------------------------------------------------------
# Translations (basic English / Arabic)
# ---------------------------------------------------------------------------
TRANSLATIONS = {
    'en': {
        'app_name': 'Rack & ID',
        'nav_catalog': 'Catalog',
        'nav_add_item': 'Add item',
        'nav_login': 'Sign in',
        'nav_logout': 'Log out',
        'search_label': 'Find by ID or name',
        'search_placeholder': 'e.g. 4521 or "denim jacket"',
        'search_btn': 'Search',
        'clear_link': 'Clear',
        'items_count': 'items in catalog',
        'item_count_singular': 'item in catalog',
        'empty_no_match': 'No items match',
        'clear_search': 'Clear search',
        'empty_no_items': 'No items yet. Start by adding the first one.',
        'add_item_cta': 'Add an item →',
        'untitled_item': 'Untitled item',
        'remove_btn': 'Remove',
        'edit_btn': 'Edit',
        'view_full': 'View full size',
        'back_to_catalog': '← Back to catalog',
        'item_id_label': 'Item ID',
        'category_label': 'Category',
        'added_label': 'Added',
        'item_not_found': 'Item not found.',
        'upload_title': 'Add a new item',
        'upload_sub': 'Use the same ID as your existing system so both stay in sync.',
        'field_id': 'Item ID',
        'field_id_hint': '(required, must match your system)',
        'field_name': 'Name',
        'field_optional': '(optional)',
        'field_category': 'Category',
        'field_photo': 'Photo',
        'field_photos': 'Photos',
        'field_photos_hint': '(one or more)',
        'no_photo_selected': 'No photo selected',
        'no_photos_selected': 'No photos selected',
        'edit_photos_current': 'Current photos',
        'edit_photos_add': 'Add more photos',
        'remove_photo': 'Remove',
        'submit_add': 'Add to catalog',
        'edit_title': 'Edit item',
        'edit_sub': 'Tick any photo to remove it, and add more below if you like — the rest stay as they are.',
        'submit_edit': 'Save changes',
        'login_title': 'Sign in',
        'login_sub': 'Enter the viewer password to browse the catalog, or the admin password to add, edit, or remove items.',
        'field_password': 'Password',
        'login_btn': 'Log in',
        'login_required_msg': 'Please log in as admin to do that.',
        'login_required_view': 'Please log in to view the catalog.',
        'login_success': 'Logged in as admin.',
        'login_success_viewer': 'Logged in.',
        'login_failed': 'Incorrect password.',
        'logout_success': 'Logged out.',
        'viewer_note': 'Viewing as guest — log in as admin to add or edit items.',
        'id_exists': 'An item with this ID already exists. Use a different ID.',
        'id_required': 'Item ID is required.',
        'image_required': 'At least one photo is required.',
        'bad_filetype': 'Only PNG, JPG, and WEBP images are supported.',
        'item_added': 'Added to the catalog.',
        'item_updated': 'Changes saved.',
    },
    'ar': {
        'app_name': 'رفّ آي دي',
        'nav_catalog': 'الكتالوج',
        'nav_add_item': 'إضافة عنصر',
        'nav_login': 'تسجيل الدخول',
        'nav_logout': 'تسجيل الخروج',
        'search_label': 'ابحث بالرقم التعريفي أو الاسم',
        'search_placeholder': 'مثال: 4521 أو "جاكيت جينز"',
        'search_btn': 'بحث',
        'clear_link': 'مسح',
        'items_count': 'عنصر في الكتالوج',
        'item_count_singular': 'عنصر واحد في الكتالوج',
        'empty_no_match': 'لا توجد عناصر مطابقة لـ',
        'clear_search': 'مسح البحث',
        'empty_no_items': 'لا توجد عناصر بعد. ابدأ بإضافة أول عنصر.',
        'add_item_cta': 'أضف عنصرًا ←',
        'untitled_item': 'عنصر بدون اسم',
        'remove_btn': 'حذف',
        'edit_btn': 'تعديل',
        'view_full': 'عرض بالحجم الكامل',
        'back_to_catalog': '→ العودة إلى الكتالوج',
        'item_id_label': 'الرقم التعريفي',
        'category_label': 'الفئة',
        'added_label': 'تمت الإضافة',
        'item_not_found': 'العنصر غير موجود.',
        'upload_title': 'إضافة عنصر جديد',
        'upload_sub': 'استخدم نفس الرقم التعريفي في نظامك الحالي ليبقى الاثنان متطابقين.',
        'field_id': 'الرقم التعريفي للعنصر',
        'field_id_hint': '(مطلوب، يجب أن يطابق نظامك)',
        'field_name': 'الاسم',
        'field_optional': '(اختياري)',
        'field_category': 'الفئة',
        'field_photo': 'الصورة',
        'field_photos': 'الصور',
        'field_photos_hint': '(صورة واحدة أو أكثر)',
        'no_photo_selected': 'لم يتم اختيار صورة',
        'no_photos_selected': 'لم يتم اختيار صور',
        'edit_photos_current': 'الصور الحالية',
        'edit_photos_add': 'إضافة المزيد من الصور',
        'remove_photo': 'حذف',
        'submit_add': 'إضافة إلى الكتالوج',
        'edit_title': 'تعديل العنصر',
        'edit_sub': 'حدد أي صورة لحذفها، وأضف المزيد أدناه إن رغبت — وتبقى بقية الصور كما هي.',
        'submit_edit': 'حفظ التغييرات',
        'login_title': 'تسجيل الدخول',
        'login_sub': 'أدخل كلمة مرور العرض لتصفح الكتالوج، أو كلمة مرور المسؤول لإضافة العناصر أو تعديلها أو حذفها.',
        'field_password': 'كلمة المرور',
        'login_btn': 'تسجيل الدخول',
        'login_required_msg': 'الرجاء تسجيل الدخول كمسؤول للقيام بذلك.',
        'login_required_view': 'الرجاء تسجيل الدخول لعرض الكتالوج.',
        'login_success': 'تم تسجيل الدخول كمسؤول.',
        'login_success_viewer': 'تم تسجيل الدخول.',
        'login_failed': 'كلمة المرور غير صحيحة.',
        'logout_success': 'تم تسجيل الخروج.',
        'viewer_note': 'أنت تتصفح كزائر — سجّل الدخول كمسؤول لإضافة العناصر أو تعديلها.',
        'id_exists': 'يوجد عنصر بهذا الرقم التعريفي مسبقًا. استخدم رقمًا مختلفًا.',
        'id_required': 'الرقم التعريفي للعنصر مطلوب.',
        'image_required': 'مطلوب صورة واحدة على الأقل.',
        'bad_filetype': 'يتم دعم صور PNG وJPG وWEBP فقط.',
        'item_added': 'تمت الإضافة إلى الكتالوج.',
        'item_updated': 'تم حفظ التغييرات.',
    },
}


def t(key):
    lang = session.get('lang', 'en')
    return TRANSLATIONS.get(lang, TRANSLATIONS['en']).get(key, key)


@app.context_processor
def inject_globals():
    lang = session.get('lang', 'en')
    return dict(
        t=t,
        current_lang=lang,
        is_rtl=(lang == 'ar'),
        is_admin=bool(session.get('is_admin')),
        is_viewer=bool(session.get('is_viewer') or session.get('is_admin')),
        r2_url=r2_url,
    )


@app.route('/lang/<code>')
def set_lang(code):
    if code in TRANSLATIONS:
        session['lang'] = code
    return redirect(request.referrer or url_for('index'))


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('is_admin'):
            flash(t('login_required_msg'), 'error')
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated


def viewer_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not (session.get('is_viewer') or session.get('is_admin')):
            flash(t('login_required_view'), 'error')
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    next_url = request.values.get('next') or url_for('index')
    if request.method == 'POST':
        password = request.form.get('password', '')
        if ADMIN_PASSWORD and password == ADMIN_PASSWORD:
            session['is_admin'] = True
            session['is_viewer'] = True
            flash(t('login_success'), 'success')
            return redirect(request.form.get('next') or url_for('index'))
        if VIEWER_PASSWORD and password == VIEWER_PASSWORD:
            session['is_viewer'] = True
            session.pop('is_admin', None)
            flash(t('login_success_viewer'), 'success')
            return redirect(request.form.get('next') or url_for('index'))
        flash(t('login_failed'), 'error')
    return render_template('login.html', next=next_url)


@app.route('/logout', methods=['POST'])
def logout():
    session.pop('is_admin', None)
    session.pop('is_viewer', None)
    flash(t('logout_success'), 'success')
    return redirect(url_for('login'))


# ---------------------------------------------------------------------------
# Routes — viewing (any signed-in role: viewer or admin)
# ---------------------------------------------------------------------------
@app.route('/')
@viewer_required
def index():
    q = request.args.get('q', '').strip()
    items = query_items(q)
    conn = get_db()
    cur = conn.cursor()
    cur.execute('SELECT COUNT(*) FROM items')
    total = cur.fetchone()[0]
    cur.close()
    conn.close()
    return render_template('index.html', items=items, query=q, total=total)


@app.route('/api/search')
@viewer_required
def api_search():
    q = request.args.get('q', '').strip()
    items = query_items(q)
    for item in items:
        if isinstance(item.get('date_added'), datetime):
            item['date_added'] = item['date_added'].isoformat()
    return jsonify(items)


@app.route('/item/<item_id>')
@viewer_required
def item_detail(item_id):
    item = get_item(item_id)
    if not item:
        flash(t('item_not_found'), 'error')
        return redirect(url_for('index'))
    return render_template('item_detail.html', item=item, images=get_item_images(item_id))


# ---------------------------------------------------------------------------
# Routes — admin only
# ---------------------------------------------------------------------------
@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload():
    if request.method == 'POST':
        item_id = request.form.get('item_id', '').strip()
        name = request.form.get('name', '').strip()
        category = request.form.get('category', '').strip()
        files = [f for f in request.files.getlist('images') if f and f.filename]

        if not item_id:
            flash(t('id_required'), 'error')
            return redirect(url_for('upload'))
        if not files:
            flash(t('image_required'), 'error')
            return redirect(url_for('upload'))
        if any(not allowed_file(f.filename) for f in files):
            flash(t('bad_filetype'), 'error')
            return redirect(url_for('upload'))

        if get_item(item_id):
            flash(t('id_exists'), 'error')
            return redirect(url_for('upload'))

        saved = [save_image(f) for f in files]
        now = datetime.now(timezone.utc)
        cover_image, cover_thumb = saved[0]

        conn = get_db()
        cur = conn.cursor()
        cur.execute(
            'INSERT INTO items (id, name, category, image_key, thumb_key, date_added) VALUES (%s, %s, %s, %s, %s, %s)',
            (item_id, name, category, cover_image, cover_thumb, now)
        )
        for position, (image_key, thumb_key) in enumerate(saved):
            cur.execute(
                'INSERT INTO item_images (id, item_id, image_key, thumb_key, position, date_added) VALUES (%s, %s, %s, %s, %s, %s)',
                (uuid.uuid4().hex, item_id, image_key, thumb_key, position, now)
            )
        conn.commit()
        cur.close()
        conn.close()
        flash(t('item_added'), 'success')
        return redirect(url_for('upload'))

    return render_template('upload.html')


@app.route('/item/<item_id>/edit', methods=['GET', 'POST'])
@login_required
def edit_item(item_id):
    item = get_item(item_id)
    if not item:
        flash(t('item_not_found'), 'error')
        return redirect(url_for('index'))

    if request.method == 'POST':
        name = request.form.get('name', '').strip()
        category = request.form.get('category', '').strip()
        remove_ids = set(request.form.getlist('remove'))
        new_files = [f for f in request.files.getlist('images') if f and f.filename]

        if any(not allowed_file(f.filename) for f in new_files):
            flash(t('bad_filetype'), 'error')
            return redirect(url_for('edit_item', item_id=item_id))

        current = get_item_images(item_id)
        keeping = [img for img in current if img['id'] not in remove_ids]
        removing = [img for img in current if img['id'] in remove_ids]

        if not keeping and not new_files:
            flash(t('image_required'), 'error')
            return redirect(url_for('edit_item', item_id=item_id))

        saved = [save_image(f) for f in new_files]
        now = datetime.now(timezone.utc)
        next_position = max([img['position'] for img in keeping], default=-1) + 1

        conn = get_db()
        cur = conn.cursor()
        for img in removing:
            cur.execute('DELETE FROM item_images WHERE id = %s', (img['id'],))
        for offset, (image_key, thumb_key) in enumerate(saved):
            cur.execute(
                'INSERT INTO item_images (id, item_id, image_key, thumb_key, position, date_added) VALUES (%s, %s, %s, %s, %s, %s)',
                (uuid.uuid4().hex, item_id, image_key, thumb_key, next_position + offset, now)
            )
        cur.execute(
            'SELECT image_key, thumb_key FROM item_images WHERE item_id = %s ORDER BY position, date_added LIMIT 1',
            (item_id,)
        )
        cover = cur.fetchone()
        cur.execute(
            'UPDATE items SET name = %s, category = %s, image_key = %s, thumb_key = %s WHERE id = %s',
            (name, category, cover[0], cover[1], item_id)
        )
        conn.commit()
        cur.close()
        conn.close()

        delete_all_item_images(removing)
        flash(t('item_updated'), 'success')
        return redirect(url_for('item_detail', item_id=item_id))

    return render_template('edit.html', item=item, images=get_item_images(item_id))


@app.route('/item/<item_id>/delete', methods=['POST'])
@login_required
def delete_item(item_id):
    item = get_item(item_id)
    if item:
        images = get_item_images(item_id)
        conn = get_db()
        cur = conn.cursor()
        cur.execute('DELETE FROM item_images WHERE item_id = %s', (item_id,))
        cur.execute('DELETE FROM items WHERE id = %s', (item_id,))
        conn.commit()
        cur.close()
        conn.close()
        delete_all_item_images(images)
    return redirect(url_for('index'))


# Make sure the schema exists on startup — safe to run repeatedly, including
# under gunicorn where __main__ below never executes.
try:
    init_db()
except Exception as exc:  # best effort; real errors surface on the first request
    print(f'[init_db] schema setup skipped: {exc}')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
