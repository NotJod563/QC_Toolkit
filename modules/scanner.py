import os
import json
from modules.registry_ops import get_latest_version_from_subkeys, key_exists, backup_exists

PRODUCTS_FILE = "products.json"


def create_products_file_if_missing():
    if not os.path.exists(PRODUCTS_FILE) or os.path.getsize(PRODUCTS_FILE) == 0:
        with open(PRODUCTS_FILE, "w", encoding="utf-8") as f:
            json.dump({"programs": []}, f, indent=2, ensure_ascii=False)


def _safe_list(value):
    if isinstance(value, list):
        return [str(x) for x in value]
    return []


def _find_existing_licenses(license_folder: str, license_names: list[str]) -> list[str]:
    found = []
    if not license_folder or not os.path.isdir(license_folder):
        return found

    for name in license_names:
        candidate = os.path.join(license_folder, name)
        if os.path.isfile(candidate):
            found.append(name)

    return found


def get_installed_programs():
    create_products_file_if_missing()

    try:
        with open(PRODUCTS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        data = {"programs": []}

    programs = []
    for item in data.get("programs", []):
        license_names = _safe_list(item.get("license_names", []))
        license_folder = item.get("license_folder", "")
        registry_path = item.get("registry_path", "").strip()

        programs.append({
            "id": item.get("id", "unknown"),
            "name": item.get("name", "Unknown Program"),
            "exe_path": item.get("exe_path", ""),
            "directory": item.get("directory", ""),
            "license_folder": license_folder,
            "license_names": license_names,
            "found_licenses": _find_existing_licenses(license_folder, license_names),
            "icon": item.get("icon", "/static/icons/default.png"),
            "registry_path": registry_path,
            "version": get_latest_version_from_subkeys(registry_path) if registry_path else "",
            "reg_is_present": key_exists(registry_path) if registry_path else False,
            "reg_backup_exists": backup_exists(item.get("name", "")) if registry_path else False,
            "log_folders": item.get("log_folders", []),
        })

    return programs
