import argparse
import csv
import logging
import sqlite3
import subprocess
import sys
import time
from pathlib import Path

import imap_outlook_oauth2
from public.sqlite_ops import Base64ASqliteStore


def setup_logger(project_root: Path) -> logging.Logger:
    log_dir = project_root / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"{time.strftime('%Y%m%d')}_validate.log"

    logger = logging.getLogger("validate_restore_export")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(name)s - %(message)s")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(formatter)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(sh)
    return logger


def resolve_provider_and_account() -> tuple[str, str]:
    profile = "outlook"
    try:
        config_path = Path(imap_outlook_oauth2.resolve_default_config_path())
        row = imap_outlook_oauth2.OutlookMailService.load_outlook_config(config_path, profile)
        provider = str(row.get("mail", "")).strip() or profile
        account = str(row.get("user", "")).strip()
        return provider, account
    except Exception:
        return profile, "mail"


def ensure_main_run(project_root: Path, logger: logging.Logger, run_main: bool) -> None:
    if not run_main:
        return
    logger.info("run main process by project rule: cmd /c main.bat")
    subprocess.run(["cmd", "/c", "main.bat"], cwd=project_root, check=True)


def check_sqlite_ready(db_path: Path, table_name: str, logger: logging.Logger) -> int:
    if not db_path.exists():
        raise RuntimeError(f"sqlite db not found: {db_path}")
    logger.info("sqlite found: %s size=%s", db_path, db_path.stat().st_size)

    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?",
            (table_name,),
        ).fetchone()
        if not row or int(row[0]) <= 0:
            raise RuntimeError(f"table not found: {table_name}")
        count = int(conn.execute(f'SELECT count(*) FROM "{table_name}"').fetchone()[0])
    if count <= 0:
        raise RuntimeError(f"table is empty: {table_name}")
    logger.info("sqlite table %s rows=%s", table_name, count)
    return count


def pick_sample_md5(project_root: Path, logger: logging.Logger) -> str:
    db_dir = project_root / "db"
    csv_files = sorted(p for p in db_dir.glob("*.csv") if not p.name.endswith("_folders.csv"))
    if not csv_files:
        raise RuntimeError("no title csv found in db/")

    for csv_path in csv_files:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                md5_value = str(row.get("Base64A_MD5", "")).strip()
                if md5_value:
                    logger.info("sample md5 picked from %s", csv_path.name)
                    return md5_value
    raise RuntimeError("no Base64A_MD5 found in title csv files")


def run_restore(
    project_root: Path,
    python_exe: Path,
    db_path: Path,
    table_name: str,
    md5_value: str,
    logger: logging.Logger,
) -> tuple[Path, Path]:
    tmp_dir = project_root / "tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    html_path = tmp_dir / "validate_restore_output.html"
    txt_path = tmp_dir / "validate_restore_output.txt"

    cmd = [
        str(python_exe),
        str(project_root / "restore_base64a.py"),
        "--db",
        str(db_path),
        "--table",
        table_name,
        "--md5",
        md5_value,
        "--save-html",
        str(html_path),
        "--save-summary",
        str(txt_path),
    ]
    logger.info("run restore command: %s", " ".join(cmd))
    subprocess.run(cmd, cwd=project_root, check=True)
    return html_path, txt_path


def assert_outputs_complete(html_path: Path, txt_path: Path, logger: logging.Logger) -> None:
    if not html_path.exists():
        raise RuntimeError(f"html output missing: {html_path}")
    if not txt_path.exists():
        raise RuntimeError(f"txt output missing: {txt_path}")

    html = html_path.read_text(encoding="utf-8", errors="replace")
    txt = txt_path.read_text(encoding="utf-8", errors="replace")

    if "<html" not in html.lower():
        raise RuntimeError("html output invalid: missing <html tag")
    if "</html>" not in html.lower():
        raise RuntimeError("html output may be incomplete: missing </html>")
    if "Email Summary (decoded)" not in txt:
        raise RuntimeError("txt output invalid: missing summary header")
    if txt.rstrip().endswith("..."):
        raise RuntimeError("txt output appears truncated: ends with ...")

    logger.info("html/txt validation passed: html_chars=%s txt_chars=%s", len(html), len(txt))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate restore_base64a export flow")
    parser.add_argument(
        "--python-exe",
        default=r"D:\0Code2\py312\python.exe",
        help="python executable path",
    )
    parser.add_argument(
        "--run-main",
        action="store_true",
        help="run cmd /c main.bat before validation",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parents[1]
    logger = setup_logger(project_root)
    logger.info("validation started: project_root=%s", project_root)

    provider, account = resolve_provider_and_account()
    safe_provider = Base64ASqliteStore.normalize_identifier(provider, fallback="outlook").lower()
    table_name = Base64ASqliteStore.build_account_table_name(account)
    db_path = project_root / "db" / f"{safe_provider}.db"
    logger.info("validation target: provider=%s account=%s db=%s table=%s", provider, account, db_path, table_name)

    ensure_main_run(project_root, logger, args.run_main)
    check_sqlite_ready(db_path, table_name, logger)
    md5_value = pick_sample_md5(project_root, logger)
    html_path, txt_path = run_restore(project_root, Path(args.python_exe), db_path, table_name, md5_value, logger)
    assert_outputs_complete(html_path, txt_path, logger)
    logger.info("validation finished: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
