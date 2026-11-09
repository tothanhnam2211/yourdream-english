# PowerShell Script — Xuất bản Gói Tĩnh MASTER THÚY — English Studio v2.0 lên GCS Bucket
# Bucket: gs://ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/
# Trang điều hành: gs://ai-studio-bucket-336999988003-asia-southeast1/control/index.html

$BUCKET = "gs://ai-studio-bucket-336999988003-asia-southeast1"
$SCRIPT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "[1/3] Đồng bộ nhánh ngang hàng master-thuy-studio/ lên $BUCKET/master-thuy-studio/ ..." -ForegroundColor Cyan
gcloud storage rsync "$SCRIPT_DIR\master-thuy-studio" "$BUCKET/master-thuy-studio" --recursive --cache-control="no-cache, max-age=0"

Write-Host "[2/3] Cập nhật trang điều hành control/index.html lên $BUCKET/control/index.html ..." -ForegroundColor Cyan
gcloud storage cp "$SCRIPT_DIR\control\index.html" "$BUCKET/control/index.html" --cache-control="no-cache, max-age=0"

if (Test-Path "$SCRIPT_DIR\student\v1.html") {
    Write-Host "[3/3] Cập nhật Bàn học Toán K6–K9 student/v1.html (Bảo mật 2 pha Web Crypto) lên $BUCKET/student/v1.html ..." -ForegroundColor Cyan
    gcloud storage cp "$SCRIPT_DIR\student\v1.html" "$BUCKET/student/v1.html" --cache-control="no-cache, max-age=0"
}

Write-Host "✅ Đã xuất bản thành công lên Google Cloud Storage!" -ForegroundColor Green
Write-Host "👉 Phòng Cô Thúy:   https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/teacher/index.html"
Write-Host "👉 Góc Học Sinh TA: https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/master-thuy-studio/student/index.html"
Write-Host "👉 Bàn Học Toán:    https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/student/v1.html"
Write-Host "👉 Trang Điều Hành: https://storage.googleapis.com/ai-studio-bucket-336999988003-asia-southeast1/control/index.html"
