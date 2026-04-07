#!/bin/bash
# owner_check デプロイスクリプト
# GitHub から最新を pull → 依存更新 → サービス再起動
set -e

cd "$(dirname "$0")"

echo "==> git pull"
git pull --ff-only

echo "==> pip install"
venv/bin/pip install -q -r requirements.txt

echo "==> restart service"
if systemctl --user is-active --quiet owner_check 2>/dev/null; then
    systemctl --user restart owner_check
elif sudo -n systemctl is-active --quiet owner_check 2>/dev/null; then
    sudo systemctl restart owner_check
else
    pkill -f "venv/bin/python3 web.py" || true
    nohup venv/bin/python3 web.py > web.log 2>&1 &
    echo "  PID: $!"
fi

echo "==> done"
