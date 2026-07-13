import shutil
import sys
from datetime import datetime
from pathlib import Path


def get_base_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = get_base_dir()
DB_PATH = BASE_DIR / "database.db"
BACKUP_DIR = BASE_DIR / "backups"


def backup_database():
    if not DB_PATH.exists():
        return False

    BACKUP_DIR.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    latest_backup = BACKUP_DIR / "database_latest.db"
    timestamp_backup = BACKUP_DIR / f"database_{timestamp}.db"

    shutil.copy2(DB_PATH, latest_backup)
    shutil.copy2(DB_PATH, timestamp_backup)

    return True