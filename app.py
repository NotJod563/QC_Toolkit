from flask import Flask, render_template, request, redirect, url_for, abort, send_file, flash
import os
import json
from copy import deepcopy
from modules.scanner import get_installed_programs, create_products_file_if_missing
from modules.license_ops import hide_licenses, show_licenses, has_backups, get_backup_root
from modules.registry_ops import backup_registry_tree, restore_registry_tree, delete_registry_tree
import win32api
import win32con
import win32ui
import win32gui
from werkzeug.utils import secure_filename
import datetime
from PIL import Image
import urllib.parse
import subprocess
import zipfile
import datetime

import tkinter as tk
from tkinter import filedialog

app = Flask(__name__)

def find_program(program_name: str):
    decoded_name = urllib.parse.unquote(program_name)

    try:
        with open("products.json", "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None

    programs = data.get("programs", [])
    for p in programs:
        if (p.get("name") or "").lower() == decoded_name.lower():
            return p

    return None

def ask_save_zip_path(default_name: str) -> str | None:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    path = filedialog.asksaveasfilename(
        title="Save logs archive",
        defaultextension=".zip",
        initialfile=default_name,
        filetypes=[("ZIP archive", "*.zip")],
    )

    root.destroy()
    return path if path else None

@app.route("/")
def index():
    create_products_file_if_missing()
    programs = get_installed_programs()
    return render_template("index.html", programs=programs, has_backups=has_backups())


@app.post("/hide_licenses")
def hide():
    report = hide_licenses()
    print("HIDE REPORT:", report)
    return redirect("/")


@app.post("/show_licenses")
def show():
    report = show_licenses()
    print("SHOW REPORT:", report)
    return redirect("/")

@app.post("/check_backups")
def check_backups():
    backup_root = get_backup_root()
    os.startfile(backup_root)
    return redirect("/")


@app.route("/add", methods=["GET", "POST"])
def add_product():
    if request.method == "POST":
        license_names_raw = request.form.get("license_names", "")
        license_names = [name.strip() for name in license_names_raw.split(",") if name.strip()]

        name = request.form.get("name", "")
        icon_filename = name.lower().replace(" ", "_") + ".png"
        raw_logs = request.form.get("log_folders", "")

        log_folders = []
        for part in raw_logs.replace(",", "\n").splitlines():
            p = part.strip().strip('"').strip("'")
            if p:
                log_folders.append(p)

        product = {
            "id": request.form.get("id"),
            "name": request.form.get("name"),
            "exe_path": request.form.get("exe_path", ""),
            "license_names": license_names,
            "registry_path": request.form.get("registry_path", ""),
            "icon": request.form.get("icon"),
            "license_folder": request.form.get("license_folder", ""),
            "log_folders": log_folders, 
        }

        try:
            with open("products.json", "r", encoding="utf-8") as f:
                data = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            data = {"programs": []}

        data["programs"].append(product)

        with open("products.json", "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        return redirect(url_for("index"))

    return render_template("add_product.html")


@app.route("/product/<program_name>")
def product_details(program_name):
    decoded_name = urllib.parse.unquote(program_name)

    try:
        with open("products.json", "r", encoding="utf-8") as f:
            data = json.load(f)
            programs = data.get("programs", [])

            for p in programs:
                if p["name"].lower() == decoded_name.lower():
                    return render_template("product_details.html", program=p)
    except Exception as e:
        print(f"Error reading product: {e}")

    return abort(404)

@app.route("/edit/<program_name>", methods=["GET", "POST"])
def edit_product(program_name):
    decoded_name = urllib.parse.unquote(program_name)

    try:
        with open("products.json", "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return "Products file not found or broken", 500

    programs = data.get("programs", [])
    program = next(
        (p for p in programs if p["name"].lower() == decoded_name.lower()),
        None
    )

    if not program:
        return "Program not found", 404

    if request.method == "POST":
        exe_path = request.form.get("exe_path", "").strip().strip('"')
        raw_logs = request.form.get("log_folders", "")

        log_folders = []
        for part in raw_logs.replace(",", "\n").splitlines():
            p = part.strip().strip('"').strip("'")
            if p:
                log_folders.append(p)


        icon = request.form.get("icon", program.get("icon", ""))

        license_names_raw = request.form.get("license_names", "")
        license_names = [
            name.strip()
            for name in license_names_raw.split(",")
            if name.strip()
        ]

        updated_program = {
            "id": request.form.get("id", program.get("id")),
            "name": request.form.get("name", program["name"]),
            "version": request.form.get("version", ""),
            "exe_path": exe_path,
            "registry_path": request.form.get("registry_path", ""),
            "directory": os.path.dirname(exe_path) if exe_path else program.get("directory", ""),
            "license_folder": request.form.get("license_folder", ""),
            "license_names": license_names,
            "icon": icon,
            "log_folders": log_folders
        }

        for i, p in enumerate(programs):
            if p["name"].lower() == decoded_name.lower():
                programs[i] = updated_program
                break

        with open("products.json", "w", encoding="utf-8") as f:
            json.dump({"programs": programs}, f, indent=2, ensure_ascii=False)

        return redirect(url_for("index"))

    scanned_licenses = []
    license_folder = program.get("license_folder", "")
    if license_folder and os.path.isdir(license_folder):
        for fname in os.listdir(license_folder):
            if fname.lower().endswith((".key", ".lic")):
                scanned_licenses.append(fname)

    return render_template(
        "edit_product.html",
        program=program,
        scanned_licenses=scanned_licenses
    )

@app.route("/duplicate/<program_name>")
def duplicate_program(program_name):
    decoded_name = urllib.parse.unquote(program_name)

    try:
        with open("products.json", "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return "Product list not found or invalid", 500

    programs = data.get("programs", [])

    original = next((p for p in programs if p["name"].lower() == decoded_name.lower()), None)
    if not original:
        return "Original product not found", 404

    new_program = deepcopy(original)
    base_name = f"{original['name']}_copy"
    unique_name = base_name
    counter = 1
    existing_names = {p["name"] for p in programs}

    while unique_name in existing_names:
        unique_name = f"{base_name}_{counter}"
        counter += 1

    new_program["name"] = unique_name
    new_program["id"] = unique_name.lower().replace(" ", "_") 
    programs.append(new_program)

    with open("products.json", "w", encoding="utf-8") as f:
        json.dump({"programs": programs}, f, indent=2, ensure_ascii=False)

    return redirect(url_for("edit_product", program_name=unique_name))

@app.route("/delete/<program_name>", methods=["POST"])
def delete_program(program_name):
    decoded_name = urllib.parse.unquote(program_name)

    try:
        with open("products.json", "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return "Product list not found or invalid", 500

    programs = data.get("programs", [])
    new_programs = [p for p in programs if p["name"].lower() != decoded_name.lower()]

    if len(programs) == len(new_programs):
        return "Product not found", 404

    # Save new list
    with open("products.json", "w", encoding="utf-8") as f:
        json.dump({"programs": new_programs}, f, indent=2, ensure_ascii=False)

    return redirect(url_for("index"))

@app.route("/extract_exe_info", methods=["POST"])
def extract_exe_info():
    exe_path = request.form.get("exe_path", "")
    program_id = request.form.get("program_id", "").strip() 

    if not exe_path or not os.path.isfile(exe_path):
        return {"error": "Invalid path"}, 400

    directory = os.path.dirname(exe_path)
    name = os.path.basename(directory)

    safe_id = (program_id or name).lower().replace(" ", "_")
    icon_name = safe_id + ".png"
    icon_path = os.path.join("static", "icons", icon_name)

    try:
        large, _ = win32gui.ExtractIconEx(exe_path, 0)
        if large:
            hicon = large[0]
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
            hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
            hdc = hdc.CreateCompatibleDC()
            hdc.SelectObject(hbmp)
            win32gui.DrawIconEx(hdc.GetHandleOutput(), 0, 0, hicon, ico_x, ico_y, 0, None, win32con.DI_NORMAL)
            bmpinfo = hbmp.GetInfo()
            bmpstr = hbmp.GetBitmapBits(True)
            img = Image.frombuffer('RGBA', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRA', 0, 1)
            img.save(icon_path)
    except Exception:
        icon_path = "static/icons/default.png"

    return {
        "name": name,
        "directory": directory,
        "icon_url": f"/{icon_path.replace(os.sep, '/')}"
    }

@app.get("/export")
def export_products():
    create_products_file_if_missing()

    return send_file(
        "products.json",
        as_attachment=True,
        download_name="toolkit_products.json",
        mimetype="application/json"
    )

@app.route("/import", methods=["GET", "POST"])
def import_products():
    create_products_file_if_missing()

    if request.method == "GET":
        return render_template("import_products.html")

    uploaded = request.files.get("file")
    mode = request.form.get("mode", "replace") 

    if not uploaded or uploaded.filename.strip() == "":
        return "No file uploaded", 400

    filename = secure_filename(uploaded.filename)
    if not filename.lower().endswith(".json"):
        return "Only .json files are allowed", 400

    try:
        imported_data = json.load(uploaded.stream)
    except Exception:
        return "Invalid JSON file", 400

    imported_programs = imported_data.get("programs")
    if not isinstance(imported_programs, list):
        return "Invalid format: expected { programs: [] }", 400

    cleaned = []
    for item in imported_programs:
        if not isinstance(item, dict):
            continue

        name = (item.get("name") or "").strip()
        if not name:
            continue

        cleaned.append({
            "id": item.get("id") or name.lower().replace(" ", "_"),
            "name": name,
            "version": item.get("version", ""),
            "exe_path": item.get("exe_path", ""),
            "registry_path": item.get("registry_path", ""),
            "directory": item.get("directory", ""),
            "license_folder": item.get("license_folder", ""),
            "license_names": item.get("license_names", []) if isinstance(item.get("license_names", []), list) else [],
            "icon": item.get("icon", "/static/icons/default.png"),
        })

    try:
        ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_name = f"products_backup_{ts}.json"
        with open("products.json", "r", encoding="utf-8") as f:
            current_raw = f.read()
        with open(backup_name, "w", encoding="utf-8") as f:
            f.write(current_raw)
    except Exception:
        pass


    try:
        with open("products.json", "r", encoding="utf-8") as f:
            current_data = json.load(f)
    except Exception:
        current_data = {"programs": []}

    current_programs = current_data.get("programs", [])
    if not isinstance(current_programs, list):
        current_programs = []

    if mode == "replace":
        final_programs = cleaned

    elif mode == "merge":
        by_id = {}
        for p in current_programs:
            if isinstance(p, dict) and p.get("id"):
                by_id[p["id"]] = p

        for p in cleaned:
            by_id[p["id"]] = p

        final_programs = list(by_id.values())

    else:
        return "Invalid mode", 400

    with open("products.json", "w", encoding="utf-8") as f:
        json.dump({"programs": final_programs}, f, indent=2, ensure_ascii=False)

    return redirect(url_for("index"))

def _load_products():
    try:
        with open("products.json", "r", encoding="utf-8") as f:
            return json.load(f).get("programs", [])
    except Exception:
        return []


@app.post("/wipe_settings/<program_name>")
def wipe_settings(program_name):
    decoded = urllib.parse.unquote(program_name)
    programs = _load_products()
    program = next((p for p in programs if p.get("name", "").lower() == decoded.lower()), None)
    if not program:
        abort(404)

    reg_path = (program.get("registry_path") or "").strip()
    if not reg_path:
        return redirect("/")

    backup_registry_tree(program.get("name", "unknown"), reg_path)
    delete_registry_tree(reg_path)

    return redirect("/")


@app.post("/restore_settings/<program_name>")
def restore_settings(program_name):
    decoded = urllib.parse.unquote(program_name)
    programs = _load_products()
    program = next((p for p in programs if p.get("name", "").lower() == decoded.lower()), None)
    if not program:
        abort(404)

    reg_path = (program.get("registry_path") or "").strip()
    if not reg_path:
        return redirect("/")

    restore_registry_tree(program.get("name", "unknown"), reg_path)
    return redirect("/")

@app.post("/logs/<program_name>")
def collect_logs(program_name):
    program = find_program(program_name)
    if not program:
        return abort(404)

    folders = program.get("log_folders", [])
    if not folders:
        return redirect(url_for("index"))

    raw_text = "\n".join(folders) if isinstance(folders, list) else str(folders)

    parts = raw_text.replace(",", "\n").splitlines()
    cleaned = []
    for part in parts:
        p = part.strip().strip('"').strip("'")
        if p:
            cleaned.append(p)

    existing = [p for p in cleaned if os.path.isdir(p)]
    if not existing:
        return redirect(url_for("index"))

    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    safe_name = (program.get("name") or "logs").replace(" ", "_")
    default_zip_name = f"{safe_name}_logs_{ts}.zip"

    zip_path = ask_save_zip_path(default_zip_name)
    if not zip_path:
        return redirect(url_for("index"))

    zip_path = os.path.abspath(zip_path).replace("/", "\\")


    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for folder in existing:
            base = os.path.basename(folder.rstrip("\\/")) or "logs"

            for root_dir, _, files in os.walk(folder):
                for file in files:
                    abs_path = os.path.join(root_dir, file)

                    rel_path = os.path.relpath(abs_path, folder)
                    arcname = os.path.join(base, rel_path)

                    z.write(abs_path, arcname)


    if os.path.isfile(zip_path):
        subprocess.Popen(["explorer.exe", f'/select,{zip_path}'])
    else:
        subprocess.Popen(["explorer.exe", os.path.dirname(zip_path)])

    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)
