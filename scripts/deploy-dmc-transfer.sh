#!/bin/bash
# Tests the DMC transfer job, then builds and deploys it to Cloud Run from services/dmc-transfer
# with the settings in services/dmc-transfer/job.env (copy job.env.example). Needs gcloud, logged
# in as someone who may deploy to the project. See docs/dmc-transfer.md.
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
dir="$root/services/dmc-transfer"
if [ ! -f "$dir/job.env" ]; then
  echo "Copy services/dmc-transfer/job.env.example to job.env and fill it in first." >&2
  exit 1
fi
# shellcheck disable=SC1091
source "$dir/job.env"
for v in GCP_PROJECT REGION JOB SERVICE_ACCOUNT FTP_PASSWORD_SECRET FTP_HOST FTP_USER SPREADSHEET_ID; do
  if [ -z "${!v:-}" ]; then echo "$v is empty in job.env" >&2; exit 1; fi
done

echo "Running the job's tests..."
(cd "$dir" && npm ci --silent && npm test --silent -- --test-reporter=dot) || {
  echo "Tests failed: nothing was deployed." >&2
  exit 1
}

# The job's own settings; FILE_IDS and REQUESTED_BY are added by the sheet for each run.
env_vars="FTP_HOST=$FTP_HOST,FTP_PORT=${FTP_PORT:-21},FTP_USER=$FTP_USER,FTP_DIR=${FTP_DIR:-}"
env_vars="$env_vars,FTP_TEMP_SUFFIX=${FTP_TEMP_SUFFIX:-.part},SPREADSHEET_ID=$SPREADSHEET_ID"

gcloud run jobs deploy "$JOB" \
  --project "$GCP_PROJECT" \
  --region "$REGION" \
  --source "$dir" \
  --service-account "$SERVICE_ACCOUNT@$GCP_PROJECT.iam.gserviceaccount.com" \
  --set-secrets "FTP_PASSWORD=$FTP_PASSWORD_SECRET:latest" \
  --set-env-vars "$env_vars" \
  --tasks 1 \
  --parallelism "${PARALLELISM:-2}" \
  --max-retries 1 \
  --task-timeout 24h \
  --cpu 1 \
  --memory 512Mi

echo
echo "Deployed. In the sheet: tick a clip in Video Assets and choose Video → Skicka markerade till DMC."
