#!/bin/bash
# Tests the media probe job, then builds and deploys it to Cloud Run from services/media-probe with
# the settings in services/media-probe/job.env (copy job.env.example). Needs gcloud, logged in as
# someone who may deploy to the project. See docs/media-probe.md.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
dir="$root/services/media-probe"
if [ ! -f "$dir/job.env" ]; then
  echo "Copy services/media-probe/job.env.example to job.env and fill it in first." >&2
  exit 1
fi
# shellcheck disable=SC1091
source "$dir/job.env"
for v in GCP_PROJECT REGION JOB SERVICE_ACCOUNT SPREADSHEET_ID; do
  if [ -z "${!v:-}" ]; then echo "$v is empty in job.env" >&2; exit 1; fi
done

echo "Running the job's tests..."
(cd "$dir" && npm test --silent -- --test-reporter=dot) || {
  echo "Tests failed: nothing was deployed." >&2
  exit 1
}

# The job's own settings; FILE_IDS and REQUESTED_BY are added by the Asset Manager for each run.
# No retries: a video that cannot be probed gets its reason in "Probe status", and a retry would
# only probe the whole task's videos again.
gcloud run jobs deploy "$JOB" \
  --project "$GCP_PROJECT" \
  --region "$REGION" \
  --source "$dir" \
  --service-account "$SERVICE_ACCOUNT@$GCP_PROJECT.iam.gserviceaccount.com" \
  --set-env-vars "SPREADSHEET_ID=$SPREADSHEET_ID" \
  --tasks 1 \
  --parallelism 1 \
  --max-retries 0 \
  --task-timeout 1h \
  --cpu 1 \
  --memory 1Gi

echo
echo "Deployed. New videos in Inbox/Video are probed by the asset watcher; Inkorg → Uppdatera"
echo "inkorgarna probes the ones that have not been."
