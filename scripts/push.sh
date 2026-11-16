#!/bin/bash
# Deploys an app to its Apps Script project: assembles apps/<app>/build from the app's src/ plus
# the shared lib/, then runs clasp push from the app folder.
#
#   scripts/push.sh master-schedule          build and push
#   scripts/push.sh master-schedule --build  build only (to inspect apps/<app>/build)
#
# clasp push replaces every file in the Apps Script project with what is in build/.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
app="${1:-}"
mode="${2:-}"
if [ -z "$app" ] || [ ! -d "$root/apps/$app/src" ]; then
  echo "Usage: scripts/push.sh <app> [--build]" >&2
  echo "Apps: $(ls "$root/apps" | tr '\n' ' ')" >&2
  exit 1
fi

dir="$root/apps/$app"
build="$dir/build"

# A file in both the app and lib would silently replace one with the other.
for f in "$root"/lib/*; do
  name="$(basename "$f")"
  if [ -e "$dir/src/$name" ]; then
    echo "apps/$app/src/$name has the same name as lib/$name; rename one of them." >&2
    exit 1
  fi
done

rm -rf "$build"
mkdir -p "$build"
cp -R "$dir/src/." "$build/"
cp -R "$root/lib/." "$build/"

echo "Built apps/$app/build:"
(cd "$build" && find . -type f | sed 's|^\./|  |' | sort)

if [ "$mode" = "--build" ]; then exit 0; fi

cd "$dir"
clasp push --force
