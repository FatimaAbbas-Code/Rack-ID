"""Helpers for the photographer's bulk photo upload.

Pure functions (no database, no storage) so they can be tested on their own:

  * parse_photo_name   - "4533_2.jpg" / "4533(2).jpg" / "4533 (2).jpg" -> the
                         product-ID candidates and the photo's number
  * prepare_photo      - any phone/camera photo -> a web-sized JPEG
                         (rotation fixed, long edge capped, colours kept)
"""
import re
from io import BytesIO

from PIL import Image, ImageCms, ImageOps

PHOTO_EXTENSIONS = {'jpg', 'jpeg', 'png', 'webp'}
MAX_EDGE = 2000          # px, long side of the stored photo
JPEG_QUALITY = 85

# "<id>(<n>)", "<id> (<n>)", "<id>_(<n>)"  and  "<id>_<n>"
_PAREN_SUFFIX = re.compile(r'^(?P<id>.+?)[\s_]*\((?P<n>\d{1,3})\)$')
_UNDERSCORE_SUFFIX = re.compile(r'^(?P<id>.+?)_(?P<n>\d{1,3})$')


class PhotoError(Exception):
    """The file is not a usable image."""


def split_extension(filename):
    """('4533_1', 'jpg') for a bare or full path; extension is lower-cased and
    '' when there is none."""
    base = re.split(r'[\\/]', filename or '')[-1].strip()
    if '.' in base:
        stem, ext = base.rsplit('.', 1)
        return stem.strip(), ext.strip().lower()
    return base, ''


def parse_photo_name(filename):
    """Return (candidates, ext).

    candidates is a list of (product_id, photo_number) pairs in the order they
    should be tried against the catalog. The whole file name is tried first, so
    a product whose real ID happens to end in "_1" still wins; then the
    "_n" / "(n)" suffix is stripped. photo_number is 0 when there is no suffix.
    """
    stem, ext = split_extension(filename)
    candidates = []
    if stem:
        candidates.append((stem, 0))
        for pattern in (_PAREN_SUFFIX, _UNDERSCORE_SUFFIX):
            m = pattern.match(stem)
            if m and m.group('id').strip():
                pair = (m.group('id').strip(), int(m.group('n')))
                if pair not in candidates:
                    candidates.append(pair)
    return candidates, ext


def resolve_photo_names(filenames, find_ids):
    """Match file names to catalog IDs.

    find_ids(lowercased_ids) -> iterable of real catalog IDs whose lower-case
    form is in that collection. Returns one dict per file name, in order:
      {'name', 'status', 'id', 'number'}
    status is one of: ok, bad_type, bad_name, unknown_id, ambiguous.
    """
    parsed = [(name,) + parse_photo_name(name) for name in filenames]
    wanted = {cid.lower() for _, cands, ext in parsed if ext in PHOTO_EXTENSIONS
              for cid, _ in cands}
    by_lower = {}
    for real in find_ids(sorted(wanted)) if wanted else []:
        by_lower.setdefault(real.lower(), []).append(real)

    out = []
    for name, cands, ext in parsed:
        row = {'name': name, 'status': 'bad_name', 'id': None, 'number': 0}
        if ext not in PHOTO_EXTENSIONS:
            row['status'] = 'bad_type'
        elif cands:
            row['status'] = 'unknown_id'
            for cid, number in cands:
                reals = by_lower.get(cid.lower(), [])
                exact = [r for r in reals if r == cid]
                if exact or len(reals) == 1:
                    row.update(status='ok', id=(exact or reals)[0], number=number)
                    break
                if len(reals) > 1:
                    row['status'] = 'ambiguous'
                    break
        out.append(row)
    return out


_SRGB = None


def _to_srgb(img):
    """Convert a wide-gamut photo (e.g. Display P3 from a phone) to sRGB so the
    colours look the same once the profile is dropped. Returns the image
    unchanged if it has no profile or the conversion fails."""
    global _SRGB
    icc = img.info.get('icc_profile')
    if not icc:
        return img
    try:
        if _SRGB is None:
            _SRGB = ImageCms.createProfile('sRGB')
        src = ImageCms.ImageCmsProfile(BytesIO(icc))
        return ImageCms.profileToProfile(img, src, _SRGB, outputMode='RGB')
    except Exception:
        return img


def prepare_photo(raw, max_edge=MAX_EDGE, quality=JPEG_QUALITY):
    """Turn an uploaded photo into a web-sized JPEG. Returns (bytes, (w, h)).

    - any of JPG / PNG / WEBP in, JPEG out
    - EXIF rotation applied (phones store portrait shots sideways)
    - long edge capped at max_edge; smaller photos are never enlarged
    - transparency flattened onto white, metadata (GPS etc.) dropped
    Raises PhotoError if the bytes are not a readable image.
    """
    try:
        with Image.open(BytesIO(raw)) as img:
            if img.format not in ('JPEG', 'PNG', 'WEBP'):
                raise PhotoError('unsupported format')
            if img.format == 'JPEG':
                img.draft('RGB', (max_edge, max_edge))   # decode at reduced size: faster, less RAM
            img = ImageOps.exif_transpose(img)
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                rgba = img.convert('RGBA')
                flat = Image.new('RGB', rgba.size, (255, 255, 255))
                flat.paste(rgba, mask=rgba.getchannel('A'))
                img = flat
            else:
                img = _to_srgb(img).convert('RGB')
            img.thumbnail((max_edge, max_edge), Image.LANCZOS)
            out = BytesIO()
            img.save(out, 'JPEG', quality=quality, optimize=True, progressive=True)
            return out.getvalue(), img.size
    except PhotoError:
        raise
    except Exception as exc:     # corrupt file, decompression bomb, truncated upload...
        raise PhotoError(str(exc)) from exc
