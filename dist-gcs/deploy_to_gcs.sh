#!/usr/bin/env bash
# Bash Script (Cloud Shell) — Xuất bản Gói Tĩnh MASTER THÚY — English Studio v2.0 lên GCS Bucket
set -euo pipefail

BUCKET="gs://ai-studio-bucket-336999988003-asia-southeast1"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "[1/3] Đồng bộ nhánh ngang hàng master-thuy-studio/ lên ${BUCKET}/master-thuy-studio/ ..."
gcloud storage rsync "${SCRIPT_DIR}/master-thuy-studio" "${BUCKET}/master-thuy-studio" --recursive --cache-control="no-cache, max-age=0"

echo "[2/3] Cập nhật trang điều hành control/index.html lên ${BUCKET}/control/index.html ..."
gcloud storage cp "${SCRIPT_DIR}/control/index.html" "${BUCKET}/control/index.html" --cache-control="no-cache, max-age=0"

if [ -f "${SCRIPT_DIR}/student/v1.html" ]; then
  echo "[3/3] Cập nhật Bàn học Toán K6–K9 student/v1.html (Bảo mật 2 pha Web Crypto) lên ${BUCKET}/student/v1.html ..."
  gcloud storage cp "${SCRIPT_DIR}/student/v1.html" "${BUCKET}/student/v1.html" --cache-control="no-cache, max-age=0"
fi

echo "✅ Đã xuất bản thành công lên Google Cloud Storage!"
echo "👉 Phòng Cô Thúy:   https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/teacher/index.html"
echo "👉 Góc Học Sinh TA: https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/student/index.html"
echo "👉 Bàn Học Toán:    https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/student/v1.html"
echo "👉 Trang Điều Hành: https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/control/index.html"
