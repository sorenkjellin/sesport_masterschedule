#!/bin/bash
# Deploys an app to its Apps Script project: runs the unit tests for lib/ and the app, assembles
# apps/<app>/build from the app's src/ plus the shared lib/, then runs clasp push from the app
# folder. Nothing is pushed when a test fails.
#
#   scripts/push.sh master-schedule          test, build and push
#   scripts/push.sh master-schedule --build  test and build only (to inspect apps/<app>/build)
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

tests=("$root"/tests/*.test.js)
if [ -d "$dir/tests" ]; then
  for f in "$dir"/tests/*.test.js; do [ -e "$f" ] && tests+=("$f"); done
else
  echo "Warning: apps/$app has no tests/ folder; only the lib tests run." >&2
fi
echo "Running unit tests..."
if ! node --test --test-reporter=dot "${tests[@]}"; then
  echo "Unit tests failed: nothing was pushed." >&2
  exit 1
fi

rm -rf "$build"
mkdir -p "$build"
cp -R "$dir/src/." "$build/"
cp -R "$root/lib/." "$build/"

echo "Built apps/$app/build:"
(cd "$build" && find . -type f | sed 's|^\./|  |' | sort)

if [ "$mode" = "--build" ]; then exit 0; fi

cd "$dir"
clasp push --force
if [ -e "$dir/src/postDeployTests.js" ]; then
  echo
  echo "Now open postDeployTests in the Apps Script editor (clasp open-script) and press Run."
fi
