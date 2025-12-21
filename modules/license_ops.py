import os
import json
import shutil
print("✅ LOADED license_ops from:", __file__)

PRODUCTS_FILE = "products.json"
HIDDEN_SUFFIX = ".hidden"


def _load_programs():
    if not os.path.exists(PRODUCTS_FILE):
        return []
    try:
        with open(PRODUCTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        programs = data.get("programs", [])
        return programs if isinstance(programs, list) else []
    except Exception:
        return []


def _get_backup_root():
    # %LOCALAPPDATA%\TOOLKIT\license_backups
    base = os.getenv("LOCALAPPDATA")
    if not base:
        # якщо LOCALAPPDATA нема — падаємо в домашню папку
        base = os.path.expanduser("~")
    return os.path.join(base, "TOOLKIT", "license_backups")


def _slug(text: str) -> str:
    text = (text or "").strip().lower().replace(" ", "_")
    allowed = []
    for ch in text:
        if ch.isalnum() or ch in ("_", "-", "."):
            allowed.append(ch)
    return "".join(allowed) or "unknown"


def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def _backup_files(program_name: str, license_folder: str, license_names: list[str]):
    backup_root = _get_backup_root()
    program_key = _slug(program_name)
    latest_dir = os.path.join(backup_root, program_key)
    _ensure_dir(latest_dir)

    copied = 0
    for fname in license_names:
        src = os.path.join(license_folder, fname)
        if os.path.isfile(src):
            dst = os.path.join(latest_dir, fname)
            shutil.copy2(src, dst)
            copied += 1

    return copied


def hide_licenses():
    print("✅ hide_licenses() called")
    programs = _load_programs()
    backup_root = _get_backup_root()
    _ensure_dir(backup_root)

    processed = 0
    renamed = 0
    backed_up = 0

    for p in programs:
        name = (p.get("name") or "Unknown").strip()

        folder = (p.get("license_folder") or "").strip().strip('"')
        raw_names = p.get("license_names", [])
        if isinstance(raw_names, str):
            license_names = [n.strip() for n in raw_names.split(",") if n.strip()]
        elif isinstance(raw_names, list):
            license_names = raw_names
        else:
            continue

        if not folder or not os.path.isdir(folder) or not license_names:
            continue

        processed += 1
        backed_up += _backup_files(name, folder, license_names)

        for fname in license_names:
            src = os.path.join(folder, fname)
            if not os.path.isfile(src):
                continue

            hidden_path = src + HIDDEN_SUFFIX
            if os.path.isfile(hidden_path):
                continue

            try:
                os.rename(src, hidden_path)
                renamed += 1
            except Exception as e:
                print(f"RENAME FAILED: {src} -> {hidden_path} | {e}")

    return {
        "backup_root": backup_root,
        "processed_programs": processed,
        "files_backed_up": backed_up,
        "files_hidden": renamed
    }


def show_licenses():
    programs = _load_programs()
    backup_root = _get_backup_root()
    _ensure_dir(backup_root)

    processed = 0
    restored = 0
    backed_up = 0

    for p in programs:
        name = (p.get("name") or "Unknown").strip()

        folder = (p.get("license_folder") or "").strip().strip('"')
        raw_names = p.get("license_names", [])

        if isinstance(raw_names, str):
            license_names = [n.strip() for n in raw_names.split(",") if n.strip()]
        elif isinstance(raw_names, list):
            license_names = raw_names
        else:
            continue

        if not folder or not os.path.isdir(folder) or not license_names:
            continue

        processed += 1
        backed_up += _backup_files(name, folder, license_names)

        for fname in license_names:
            original = os.path.join(folder, fname)
            hidden_path = original + HIDDEN_SUFFIX

            if os.path.isfile(hidden_path) and not os.path.isfile(original):
                try:
                    os.rename(hidden_path, original)
                    restored += 1
                except Exception:
                    pass

    return {
        "backup_root": backup_root,
        "processed_programs": processed,
        "files_backed_up": backed_up,
        "files_restored": restored
    }


def has_backups():
    backup_root = _get_backup_root()
    if not os.path.isdir(backup_root):
        return False
    for _, _, files in os.walk(backup_root):
        if files:
            return True
    return False

def get_backup_root():
    backup_root = _get_backup_root()
    _ensure_dir(backup_root)
    return backup_root
