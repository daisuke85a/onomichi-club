#!/bin/sh
# apps-script/index.html を元に、他の配布形を生成する
#   preview.html                → 本文だけ（Claude Artifact / 単体プレビュー用、サンプルデータで動く）
#   cloudflare/public/index.html → Cloudflare Workers 用（/api/clubs から読む）
set -eu
cd "$(dirname "$0")/.."
SRC=apps-script/index.html

# <title> から </script> までを本文として切り出す
sed -n '/^<title>/,/^<\/script>$/p' "$SRC" | grep -v -x -e '</head>' -e '<body>' > preview.html

sed 's|const CONFIG = { apiUrl: "", csvUrl: "" };|const CONFIG = { apiUrl: "/api/clubs", csvUrl: "" };|' "$SRC" > cloudflare/public/index.html
grep -q 'apiUrl: "/api/clubs"' cloudflare/public/index.html || { echo "build: apiUrl の置換に失敗" >&2; exit 1; }
echo "build: preview.html, cloudflare/public/index.html を更新しました"
