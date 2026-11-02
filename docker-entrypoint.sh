#!/bin/sh
# Keep everything that must survive a redeploy on the /data volume.
set -e
for d in uploads machine_files files; do
  mkdir -p "/data/$d"; [ -e "/app/$d" ] || ln -s "/data/$d" "/app/$d"
done
[ -e /app/_hls_secret ] || ln -s /data/_hls_secret /app/_hls_secret   # login-token key
exec "$@"
