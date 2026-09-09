"""Write a JSON snapshot of the catalog to R2, under backups/.

Run it on a schedule so the store's data is captured even if nobody clicks
the Export button. On Render: add a Cron Job service pointing at this repo
with the command `python backup.py` and a schedule like `0 2 * * *`
(daily, 02:00 UTC). It needs the same DATABASE_URL and R2_* environment
variables as the web service.

Each run writes:
  backups/rack-and-id-YYYYMMDD-HHMMSS.json   (timestamped snapshot)
  backups/latest.json                         (always the newest)
and prunes timestamped snapshots beyond the newest KEEP.

Restore from a snapshot via the app's Import page (upload the file, or use
"Restore the latest automatic backup", which reads backups/latest.json).
"""
import json
import sys
from datetime import datetime, timezone

from app import build_backup_payload, get_r2_client, R2_BUCKET_NAME

KEEP = 30


def main():
    payload = build_backup_payload()
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode('utf-8')
    stamp = datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')
    key = f'backups/rack-and-id-{stamp}.json'

    s3 = get_r2_client()
    s3.put_object(Bucket=R2_BUCKET_NAME, Key=key, Body=body,
                  ContentType='application/json')
    s3.put_object(Bucket=R2_BUCKET_NAME, Key='backups/latest.json', Body=body,
                  ContentType='application/json')
    print(f'backup written: {key}  ({payload["count"]} items, {len(body)} bytes)')

    resp = s3.list_objects_v2(Bucket=R2_BUCKET_NAME, Prefix='backups/rack-and-id-')
    old_keys = sorted(obj['Key'] for obj in resp.get('Contents', []))
    for stale in old_keys[:-KEEP]:
        s3.delete_object(Bucket=R2_BUCKET_NAME, Key=stale)
        print(f'pruned: {stale}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
