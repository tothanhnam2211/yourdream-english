# -*- coding: utf-8 -*-
"""
Đồng bộ MASTER THÚY — English Studio v2.0 lên Google Drive (YOURDREAM_AI_WORKSPACE),
Git Repository & Kiểm chứng đọc ngược (Verified Readback SHA-256).
"""
import hashlib
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

PROJECT_ROOT = Path(r"D:\YOURDREAM.VN\master-thuy-english-studio")
DRIVE_STUDIO_ROOT = Path(r"G:\Drive của tôi\YOURDREAM_AI_WORKSPACE\ENGLISH_TEACHING_STUDIO")
CHECKP_DIR = DRIVE_STUDIO_ROOT / "CHECKPOINTS" / "MASTER_THUY_CHECKP_v2_0_20261105"
GENERATED_DOCX_DRIVE = DRIVE_STUDIO_ROOT / "GENERATED_DOCX_v2_0"
DIST_GCS_DRIVE = DRIVE_STUDIO_ROOT / "DIST_GCS_v2_0"


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def sync_to_google_drive():
    print("=" * 72)
    print(" 1. ĐỒNG BỘ LÊN GOOGLE DRIVE (YOURDREAM_AI_WORKSPACE)")
    print("=" * 72)
    CHECKP_DIR.mkdir(parents=True, exist_ok=True)
    GENERATED_DOCX_DRIVE.mkdir(parents=True, exist_ok=True)
    DIST_GCS_DRIVE.mkdir(parents=True, exist_ok=True)
    (DRIVE_STUDIO_ROOT / "00_PRINT_QUEUE" / "PROCESSED").mkdir(parents=True, exist_ok=True)

    # 1. Copy CCP markdown to Drive root & Checkpoint folder
    ccp_src = PROJECT_ROOT / "CHECKP_MASTER_THUY_ENGLISH_STUDIO_v2_0_CCP.md"
    ccp_dst_root = DRIVE_STUDIO_ROOT / "CHECKP_MASTER_THUY_ENGLISH_STUDIO_v2_0_CCP.md"
    ccp_dst_chk = CHECKP_DIR / "CHECKP_MASTER_THUY_ENGLISH_STUDIO_v2_0_CCP.md"
    shutil.copy2(ccp_src, ccp_dst_root)
    shutil.copy2(ccp_src, ccp_dst_chk)

    # 2. Copy all generated DOCX files (Student + Teacher 101-104) to Drive
    gen_dir = PROJECT_ROOT / "public" / "generated-docx"
    docx_count = 0
    if gen_dir.exists():
        for docx_file in gen_dir.glob("*.docx"):
            shutil.copy2(docx_file, GENERATED_DOCX_DRIVE / docx_file.name)
            docx_count += 1
    print(f"  [OK] Đã đồng bộ {docx_count} tệp Word (.docx) xuất bản sang: {GENERATED_DOCX_DRIVE}")

    # 3. Copy QA verification report & catalog to Checkpoint folder
    for rel in [
        "public/shards/qa-verification-report.json",
        "public/shards/worksheets-catalog.json",
        "dist-gcs/control/index.html",
        "dist-gcs/deploy_to_gcs.ps1",
        "dist-gcs/deploy_to_gcs.sh",
    ]:
        src = PROJECT_ROOT / rel
        if src.exists():
            shutil.copy2(src, CHECKP_DIR / src.name)

    # 4. Create full production ZIP archive (source + DB + shards + dist-gcs control) on Drive
    zip_path = CHECKP_DIR / "MASTER_THUY_ENGLISH_STUDIO_v2_0_PRODUCTION.zip"
    exclude_dirs = {"raw_63_archive", ".git", "__pycache__", "dist-gcs"}
    file_count = 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in PROJECT_ROOT.rglob("*"):
            if item.is_file():
                rel_parts = item.relative_to(PROJECT_ROOT).parts
                if any(part in exclude_dirs for part in rel_parts):
                    continue
                zf.write(item, item.relative_to(PROJECT_ROOT))
                file_count += 1
    zip_hash = sha256_file(zip_path)
    print(f"  [OK] Đã đóng gói {file_count} tệp vào {zip_path.name} ({zip_path.stat().st_size / (1024*1024):.2f} MB)")
    print(f"  [OK] SHA-256 ZIP: {zip_hash}")

    # 5. Verified readback from Google Drive
    src_hash = sha256_file(ccp_src)
    dst_hash = sha256_file(ccp_dst_chk)
    assert src_hash == dst_hash, "SHA-256 mismatch on CCP readback!"
    print(f"  [OK] Verified Readback CCP SHA-256 khớp 100%: {dst_hash}")


def sync_git_repo():
    print("\n" + "=" * 72)
    print(" 2. ĐỒNG BỘ GIT REPOSITORY (D:\\YOURDREAM.VN\\master-thuy-english-studio)")
    print("=" * 72)
    gitignore = PROJECT_ROOT / ".gitignore"
    gitignore.write_text(
        "__pycache__/\n*.pyc\nraw_63_archive/\ndist-gcs/master-thuy-studio/archive-files/\n",
        encoding="utf-8",
    )
    if not (PROJECT_ROOT / ".git").exists():
        subprocess.run(["git", "init"], cwd=str(PROJECT_ROOT), check=False)
    subprocess.run(["git", "add", "-A"], cwd=str(PROJECT_ROOT), check=False)
    res = subprocess.run(
        [
            "git",
            "commit",
            "-m",
            "Release MASTER THUY English Studio v2.0 (Production & Cloud-Ready, 953 items / 63 files, Hybrid DOCX & 00_PRINT_QUEUE)",
        ],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    print("  [OK] Git commit status:", (res.stdout or res.stderr).strip().splitlines()[0] if (res.stdout or res.stderr) else "Committed")

    # Check if remote exists
    rem = subprocess.run(
        ["git", "remote", "-v"],
        cwd=str(PROJECT_ROOT),
        capture_output=True,
        text=True,
    )
    if rem.stdout.strip():
        print("  [INFO] Git remote detected, pushing...")
        subprocess.run(["git", "push"], cwd=str(PROJECT_ROOT), check=False)
    else:
        print("  [INFO] Git repository committed locally at D:\\YOURDREAM.VN\\master-thuy-english-studio (.git)")


if __name__ == "__main__":
    sync_to_google_drive()
    sync_git_repo()
