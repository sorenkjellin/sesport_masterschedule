#!/bin/bash
# Fetches an app's files from its Apps Script project (for changes made in the browser editor)
# and copies them into apps/<app>/src. The shared lib files are not copied back: if the editor's
# copy of one differs from lib/, this says so and you merge by hand.
#
#   scripts/pull.sh master-schedule
#
# Review with "git diff" afterwards.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
app="${1:-}"
if [ -z "$app" ] || [ ! -d "$root/apps/$app/src" ]; then
  echo "Usage: scripts/pull.sh <app>" >&2
  exit 1
fi

dir="$root/apps/$app"
rm -rf "$dir/build"
mkdir -p "$dir/build"
(cd "$dir" && clasp pull)

for f in "$dir"/build/*; do
  name="$(basename "$f")"
  if [ -e "$root/lib/$name" ]; then
    cmp -s "$f" "$root/lib/$name" || echo "Differs from lib/$name: apps/$app/build/$name (not copied)"
  else
    cp "$f" "$dir/src/$name"
  fi
done

echo "Pulled into apps/$app/src. Check with: git diff"
