import os
import json
import base64
import winreg


def _get_backup_root():
    base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "TOOLKIT", "registry_backups")
    os.makedirs(path, exist_ok=True)
    return path


def _slug(text: str) -> str:
    text = (text or "").strip().lower().replace(" ", "_")
    allowed = []
    for ch in text:
        if ch.isalnum() or ch in ("_", "-", "."):
            allowed.append(ch)
    return "".join(allowed) or "unknown"


def _backup_file_for_program(program_name: str) -> str:
    return os.path.join(_get_backup_root(), f"{_slug(program_name)}.json")


def _parse_registry_path(reg_path: str):
    if not reg_path:
        return None, None

    s = reg_path.strip().lstrip("\\")
    if s.lower().startswith("computer\\"):
        s = s[len("computer\\"):]

    parts = s.split("\\", 1)
    if len(parts) != 2:
        return None, None

    hive_str, subkey = parts[0].upper(), parts[1]

    hive_map = {
        "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
        "HKCU": winreg.HKEY_CURRENT_USER,
        "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
        "HKLM": winreg.HKEY_LOCAL_MACHINE,
    }

    hive = hive_map.get(hive_str)
    if not hive:
        return None, None

    return hive, subkey


def key_exists(reg_path: str) -> bool:
    hive, subkey = _parse_registry_path(reg_path)
    if not hive:
        return False
    try:
        k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ)
        winreg.CloseKey(k)
        return True
    except OSError:
        return False


def _read_key_tree(hive, subkey: str):
    """
    Рекурсивно читає ключ: values + subkeys
    """
    node = {"values": [], "subkeys": {}}

    try:
        k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ)
    except OSError:
        return None

    i = 0
    while True:
        try:
            name, value, vtype = winreg.EnumValue(k, i)
            if vtype == winreg.REG_BINARY and isinstance(value, (bytes, bytearray)):
                value = {"__binary__": base64.b64encode(value).decode("ascii")}
            node["values"].append({"name": name, "type": vtype, "data": value})
            i += 1
        except OSError:
            break

    j = 0
    while True:
        try:
            child = winreg.EnumKey(k, j)
            child_path = f"{subkey}\\{child}"
            child_node = _read_key_tree(hive, child_path)
            if child_node is not None:
                node["subkeys"][child] = child_node
            j += 1
        except OSError:
            break

    winreg.CloseKey(k)
    return node


def backup_registry_tree(program_name: str, reg_path: str) -> str | None:

    hive, subkey = _parse_registry_path(reg_path)
    if not hive:
        return None

    tree = _read_key_tree(hive, subkey)
    if tree is None:
        return None

    payload = {"registry_path": reg_path, "tree": tree}
    file_path = _backup_file_for_program(program_name)

    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)

    return file_path


def _ensure_key(hive, subkey: str):
    winreg.CreateKey(hive, subkey)


def _write_key_tree(hive, subkey: str, node: dict):
    _ensure_key(hive, subkey)

    k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_SET_VALUE)

    for v in node.get("values", []):
        name = v["name"]
        vtype = int(v["type"])
        data = v["data"]

        if isinstance(data, dict) and "__binary__" in data:
            data = base64.b64decode(data["__binary__"].encode("ascii"))

        winreg.SetValueEx(k, name, 0, vtype, data)

    winreg.CloseKey(k)

    for child_name, child_node in node.get("subkeys", {}).items():
        _write_key_tree(hive, f"{subkey}\\{child_name}", child_node)


def restore_registry_tree(program_name: str, reg_path: str) -> bool:

    file_path = _backup_file_for_program(program_name)
    if not os.path.isfile(file_path):
        return False

    hive, subkey = _parse_registry_path(reg_path)
    if not hive:
        return False

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            payload = json.load(f)

        tree = payload.get("tree")
        if not isinstance(tree, dict):
            return False

        _write_key_tree(hive, subkey, tree)
        return True
    except Exception:
        return False


def delete_registry_tree(reg_path: str) -> bool:

    hive, subkey = _parse_registry_path(reg_path)
    if not hive:
        return False

    try:
        winreg.DeleteKey(hive, subkey)
        return True
    except OSError:
        pass

    try:
        k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ | winreg.KEY_WRITE)
    except OSError:
        return False

    while True:
        try:
            child = winreg.EnumKey(k, 0)
            delete_registry_tree(f"{_hive_name(hive)}\\{subkey}\\{child}")
        except OSError:
            break

    winreg.CloseKey(k)

    try:
        winreg.DeleteKey(hive, subkey)
        return True
    except OSError:
        return False


def _hive_name(hive):
    if hive == winreg.HKEY_CURRENT_USER:
        return "HKEY_CURRENT_USER"
    if hive == winreg.HKEY_LOCAL_MACHINE:
        return "HKEY_LOCAL_MACHINE"
    return "HKEY_CURRENT_USER"


def backup_exists(program_name: str) -> bool:
    return os.path.isfile(_backup_file_for_program(program_name))


def open_backup_folder_in_explorer():
    os.startfile(_get_backup_root())


def _version_key_tuple(v: str):
    parts = []
    for p in (v or "").split("."):
        try:
            parts.append(int(p))
        except ValueError:
            parts.append(0)
    return tuple(parts) if parts else (0,)


def get_latest_version_from_subkeys(reg_path: str) -> str:

    hive, subkey = _parse_registry_path(reg_path)
    if not hive:
        return ""

    try:
        k = winreg.OpenKey(hive, subkey, 0, winreg.KEY_READ)
    except OSError:
        return ""

    versions = []
    i = 0
    while True:
        try:
            child = winreg.EnumKey(k, i)
            versions.append(child)
            i += 1
        except OSError:
            break

    winreg.CloseKey(k)

    if not versions:
        return ""

    versions.sort(key=_version_key_tuple)
    return versions[-1]
