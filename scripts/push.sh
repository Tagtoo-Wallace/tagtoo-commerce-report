#!/bin/bash
# tagtoo-commerce-report/scripts/push.sh：以 Tagtoo-Wallace 身份推送本 repo，結束時一律切回原本的作用中帳號。
# 原因：git 對 github.com 的憑證走 `gh auth git-credential`，它只認 gh 目前的作用中帳號；本機平時作用中的是個人帳號
#       wallace255201，推不進 Tagtoo-Wallace 的 repo。用法：scripts/push.sh [git push 的其他參數]
set -euo pipefail
cd "$(dirname "$0")/.."
cur="$(gh auth status 2>&1 | grep -B1 'Active account: true' | head -1 | sed -E 's/.*account ([^ ]+).*/\1/')"
[ -n "$cur" ] || cur="wallace255201"
restore() { gh auth switch --user "$cur" >/dev/null 2>&1 || echo "⚠ 切回 $cur 失敗，請手動 gh auth switch" >&2; }
trap restore EXIT
gh auth switch --user Tagtoo-Wallace >/dev/null
git push "$@"
echo "✓ 已以 Tagtoo-Wallace 推送，作用中帳號切回 $cur"
