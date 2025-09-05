# Created by Alexander Clark of Metatheria, LLC
# Edited and updated by InfiNet Solutions Inc
# Creative Commons CC0 v1.0 Universal Public Domain Dedication. No Rights Reserved
# Version 2.0 (Dropdown Date + Remember Me + UI Enhancements + Theme + Progress + Stop)
# + Partial CSV on Stop

import tkinter as tk
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk
import sys
import os
import webbrowser
import datetime
import urllib
import threading
import queue

import requests
from bs4 import BeautifulSoup
import csv
import winsound
from tkcalendar import DateEntry  # date dropdown

# ---------- App metadata ----------
APP_NAME = "Nebraska Courts E-Services Scraper V2"
APP_VER  = "2.0"
ORG_NAME = "InfiNet Solutions"
ORG_URL  = "https://www.omahait.com"
DOCS_URL = "https://www.omahait.com/ai-automation"  # adjust if needed
ALEX_NAME = "Alexander Clark of Metatheria, LLC"
ALEX_URL  = "https://clarkmanagementconsulting.com"

# ---------- Windows DPAPI encryption for saved creds ----------
import base64
import configparser
import ctypes
from ctypes import Structure, POINTER, byref, c_byte, create_string_buffer, cast, windll, string_at
from ctypes.wintypes import DWORD

CONFIG_DIR  = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "NEJusticeScraper")
CONFIG_PATH = os.path.join(CONFIG_DIR, "settings.ini")

def ensure_config_dir():
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
    except Exception:
        pass

class DATA_BLOB(Structure):
    _fields_ = [("cbData", DWORD),
                ("pbData", POINTER(c_byte))]

def _bytes_to_blob(b: bytes):
    buf = create_string_buffer(b, len(b))
    blob = DATA_BLOB(len(b), cast(buf, POINTER(c_byte)))
    return blob, buf

def _blob_to_bytes(blob: DATA_BLOB) -> bytes:
    cb = int(blob.cbData)
    if cb == 0:
        return b""
    data = string_at(blob.pbData, cb)
    windll.kernel32.LocalFree(blob.pbData)
    return data

def dpapi_encrypt_string(plaintext: str) -> str:
    if not plaintext:
        return ""
    data = plaintext.encode("utf-8")
    blob_in, _buf = _bytes_to_blob(data)
    blob_out = DATA_BLOB()
    ok = windll.crypt32.CryptProtectData(
        byref(blob_in), None, None, None, None, 0, byref(blob_out)
    )
    if not ok:
        return ""
    enc = _blob_to_bytes(blob_out)
    return base64.b64encode(enc).decode("ascii")

def dpapi_decrypt_string(token: str) -> str:
    if not token:
        return ""
    try:
        enc = base64.b64decode(token)
    except Exception:
        return ""
    blob_in, _buf = _bytes_to_blob(enc)
    blob_out = DATA_BLOB()
    ok = windll.crypt32.CryptUnprotectData(
        byref(blob_in), None, None, None, None, 0, byref(blob_out)
    )
    if not ok:
        return ""
    raw = _blob_to_bytes(blob_out)
    try:
        return raw.decode("utf-8")
    except Exception:
        return ""

def load_settings():
    ensure_config_dir()
    cfg = configparser.ConfigParser(interpolation=None)
    if os.path.exists(CONFIG_PATH):
        try:
            cfg.read(CONFIG_PATH, encoding="utf-8")
        except Exception:
            pass
    remember = False
    username = ""
    password = ""
    if "auth" in cfg:
        sec = cfg["auth"]
        remember = sec.get("remember_me", "0") == "1"
        username = sec.get("username", "") if remember else ""
        password = dpapi_decrypt_string(sec.get("password_enc", "")) if remember else ""
    last_output_dir = ""
    if "prefs" in cfg:
        last_output_dir = cfg["prefs"].get("last_output_dir", "")
    return remember, username, password, last_output_dir

def save_settings(remember: bool, username: str, password: str, last_output_dir: str):
    ensure_config_dir()
    cfg = configparser.ConfigParser(interpolation=None)
    cfg["auth"] = {
        "remember_me": "1" if remember else "0",
        "username": username if remember else "",
        "password_enc": dpapi_encrypt_string(password) if remember else "",
    }
    cfg["prefs"] = {
        "last_output_dir": last_output_dir or "",
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        cfg.write(f)

# ---------- Theme loader (Azure .tcl) ----------
def resource_path(rel):
    # Works for dev, one-folder, and one-file builds
    if hasattr(sys, '_MEIPASS'):       # one-file: temp extraction dir
        base = sys._MEIPASS
    elif getattr(sys, 'frozen', False):  # one-folder
        base = os.path.dirname(sys.executable)
    else:                                # dev
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel)

def load_azure_theme(root):
    style = ttk.Style(root)
    try:
        azure_path = resource_path(os.path.join("assets", "azure.tcl"))
        if os.path.exists(azure_path):
            root.tk.call("source", azure_path)
            style.theme_use("azure")
        else:
            azure_dark = resource_path(os.path.join("assets", "azure-dark.tcl"))
            if os.path.exists(azure_dark):
                root.tk.call("source", azure_dark)
                style.theme_use("azure")
            else:
                style.theme_use("vista" if "vista" in style.theme_names() else "clam")
    except Exception:
        try:
            style.theme_use("vista")
        except Exception:
            style.theme_use("clam")
    return style

# ---------- Logging queues ----------
LOG_QUEUE = queue.Queue()
EVENT_QUEUE = queue.Queue()
CANCEL_EVENT = threading.Event()
WORKER = None

def ui_log(msg: str): LOG_QUEUE.put(msg)
def ui_event(kind, payload=None): EVENT_QUEUE.put((kind, payload))

# ---------- Scraping logic ----------
def validate(date_text):
    try:
        datetime.datetime.strptime(date_text, '%m/%d/%Y')
    except ValueError:
        raise ValueError("Incorrect Date format, should be mm/dd/yyyy")

def desktop_folder():
    return os.path.join(os.path.expanduser("~"), "Desktop")

county_numbers_dict = {"Adams" : "14",
 "Antelope" : "26", "Arthur" : "91", "Banner" : "85", "Blaine" : "86", "Boone" : "23",
 "Box Butte" : "65", "Boyd" : "63", "Brown" : "75", "Buffalo" : "09", "Burt" : "31",
 "Butler" : "25", "Cass" : "20", "Cedar" : "13", "Chase" : "72", "Cherry" : "66",
 "Cheyenne" : "39", "Clay" : "30", "Colfax" : "43", "Cuming" : "24", "Custer" : "04",
 "Dakota" : "70", "Dawes" : "69", "Dawson" : "18", "Deuel" : "78", "Dixon" : "35",
 "Dodge" : "05", "Douglas" : "01", "Dundy" : "76", "Fillmore" : "34", "Franklin" : "50",
 "Frontier" : "60", "Furnas" : "38", "Gage" : "03", "Garden" : "77", "Garfield" : "83",
 "Gosper" : "73", "Grant" : "92", "Greeley" : "62", "Hall" : "08", "Hamilton" : "28",
 "Harlan" : "51", "Hayes" : "79", "Hitchcock" : "67", "Holt" : "36", "Hooker" : "93",
 "Howard" : "49", "Jefferson" : "33", "Johnson" : "57", "Kearney" : "52", "Keith" : "68",
 "Keya Paha" : "82", "Kimball" : "71", "Knox" : "12", "Lancaster" : "02", "Lincoln" : "15",
 "Logan" : "87", "Loup" : "88", "Madison" : "07", "McPherson" : "90", "Merrick" : "46",
 "Morrill" : "64", "Nance" : "58", "Nemaha" : "44", "Nuckolls" : "42", "Otoe" : "11",
 "Pawnee" : "54", "Perkins" : "74", "Phelps" : "37", "Pierce" : "40", "Platte" : "10",
 "Polk" : "41", "Red Willow" : "48", "Richardson" : "19", "Rock" : "81", "Saline" : "22",
 "Sarpy" : "59", "Saunders" : "06", "Scotts Bluff" : "21", "Seward" : "16", "Sheridan" : "61",
 "Sherman" : "56", "Sioux" : "80", "Stanton" : "53", "Thayer" : "32", "Thomas" : "89",
 "Thurston" : "55", "Valley" : "47", "Washington" : "29", "Wayne" : "27", "Webster" : "45",
 "Wheeler" : "84", "York" : "17"}

# ---- NEW: helpers to write a partial CSV on cancel ----
def _parse_case_and_county_from_url(url: str):
    """Robustly build (case_number, county_name) from case URL query params."""
    try:
        parsed = urllib.parse.urlparse(url)
        qs = urllib.parse.parse_qs(parsed.query)
        cy = (qs.get("case_year") or [""])[0]
        cid = (qs.get("case_id") or [""])[0]
        cnum = (qs.get("county_num") or [""])[0]
        case_no = f"{cy}CI{cid}" if cy and cid else ""
        county = ""
        if cnum:
            for k, v in county_numbers_dict.items():
                if v == cnum:
                    county = k
                    break
        return case_no, county
    except Exception:
        return "", ""

def write_partial_csv(addresses, restitution_cases, targetDate, out_dir):
    """Write whatever we have to a _partial CSV."""
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        pass

    headers = ['name', 'address', 'city state zip', 'case number', 'county']
    rows = []

    # Prefer addresses we've started enriching
    if addresses:
        for addr in addresses:
            url = addr[0] if addr else ""
            case_no, county = _parse_case_and_county_from_url(url) if url else ("", "")
            name = addr[1] if len(addr) > 1 else ""
            addr1 = addr[2] if len(addr) > 2 else ""
            city_state_zip = addr[3] if len(addr) > 3 else ""
            rows.append([name, addr1, city_state_zip, case_no, county])

    # If we had not reached docket fetch yet, fall back to restitution_cases
    if not rows and restitution_cases:
        for rc in restitution_cases:
            if len(rc) >= 9:
                url = rc[8]
                case_no, county_from_url = _parse_case_and_county_from_url(url)
                county = rc[7] if len(rc) >= 8 else county_from_url
                rows.append(["", "", "", case_no, county])

    # If truly nothing, just write headers
    data = [headers] + rows
    filename = (
        "eviction_cases_for_"
        + datetime.datetime.strptime(targetDate, '%m/%d/%Y').strftime('%Y-%m-%d')
        + "_partial_generated_on_"
        + datetime.datetime.now().strftime('%Y-%m-%d-%H-%M')
        + ".csv"
    )
    out_path = os.path.join(out_dir or desktop_folder(), filename)
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f, quoting=csv.QUOTE_ALL).writerows(data)
    return out_path
# -------------------------------------------------------

def scrapeCalendar():
    """
    Runs in a background thread. Uses CANCEL_EVENT for graceful stop.
    """
    try:
        # Persist settings each run
        save_settings(remember_var.get(), user_entry.get(), pass_entry.get(), save_dir_var.get())

        # County list
        if (c_option.get() == "2"):
            counties_list = ["Adams", "Antelope", "Arthur", "Banner", "Blaine", "Boone", "Box Butte", "Boyd", "Brown", "Buffalo", "Burt", "Butler", "Cass", "Cedar", "Chase", "Cherry", "Cheyenne", "Clay", "Colfax", "Cuming", "Custer", "Dakota", "Dawes", "Dawson", "Deuel", "Dixon", "Dodge", "Douglas", "Dundy", "Fillmore", "Franklin", "Frontier", "Furnas", "Gage", "Garden", "Garfield", "Gosper", "Grant", "Greeley", "Hall", "Hamilton", "Harlan", "Hayes", "Hitchcock", "Holt", "Hooker", "Howard", "Jefferson", "Johnson", "Kearney", "Keith", "Keya Paha", "Kimball", "Knox", "Lancaster", "Lincoln", "Logan", "Loup", "Madison", "McPherson", "Merrick", "Morrill", "Nance", "Nemaha", "Nuckolls", "Otoe", "Pawnee", "Perkins", "Phelps", "Pierce", "Platte", "Polk", "Red Willow", "Richardson", "Rock", "Saline", "Sarpy", "Saunders", "Scotts Bluff", "Seward", "Sheridan", "Sherman", "Sioux", "Stanton", "Thayer", "Thomas", "Thurston", "Valley", "Washington", "Wayne", "Webster", "Wheeler", "York"]
        if (c_option.get() == "1"):
            counties_list = ["Douglas", "Lancaster", "Sarpy"]
        if (c_option.get() == "3"):
            counties_list = ["Douglas", "Lancaster", "Sarpy", "Hall", "Buffalo", "Dodge", "Scotts Bluff", "Madison", "Platte", "Lincoln"]

        ui_event("phase", "Processing…")
        ui_log("Processing...")
        targetDate = entry1.get()
        username = user_entry.get()
        password = pass_entry.get()
        validate(targetDate)

        restitution_cases = []
        addresses = []

        # calendar scrape
        for county in counties_list:
            if CANCEL_EVENT.is_set():
                ui_event("phase", "Stopping… writing partial CSV…")
                out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
                ui_event("done", out_path)
                return
            ui_log(f"Getting case numbers for eviction cases (Restitution, Real Fed, FED or LLT is in description) for {targetDate} from the calendar for {county} county...")
            params = {
              ('court', 'C'),
              ('countyC', county),
              ('countyD', ''),
              ('selectRadio', 'date'),
              ('searchField', targetDate),
              ('submitButton', 'Submit'),
            }
            try:
                response = requests.get('https://www.nebraska.gov/courts/calendar/index.cgi', params=params, timeout=45)
            except Exception as e:
                ui_log(f"[WARN] Calendar failed for {county}: {e}")
                continue
            if CANCEL_EVENT.is_set():
                ui_event("phase", "Stopping… writing partial CSV…")
                out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
                ui_event("done", out_path)
                return
            soup = BeautifulSoup(response.content, 'lxml')
            rows = soup.find_all('tr')
            for row in rows:
                if CANCEL_EVENT.is_set():
                    ui_event("phase", "Stopping… writing partial CSV…")
                    out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
                    ui_event("done", out_path)
                    return
                if "Restitution" in row.get_text() or "Real Fed" in row.get_text() or "LLT" in row.get_text() or "FED" in row.get_text():
                    listrow = row.get_text().splitlines()
                    if ("CR" not in listrow[6]):
                        listrow.append(county)
                        ui_log(f"Adding {listrow[7]} county case number {listrow[6]} to the list to scrape.")
                        case_url = 'https://www.nebraska.gov/justice/case.cgi?search=1&from_case_search=1&court_type=C&county_num='
                        case_url += county_numbers_dict.get(listrow[7])
                        case_url += '&case_type=CI&case_year='
                        case_url += listrow[6][2:4]
                        case_url += '&case_id='
                        case_url += listrow[6][4:]
                        case_url += '&client_data=&search=Search+Now'
                        listrow.append(case_url)
                        restitution_cases.append(listrow)

        if CANCEL_EVENT.is_set():
            ui_event("phase", "Stopping… writing partial CSV…")
            out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
            ui_event("done", out_path)
            return

        ui_event("phase", "Deduplicating list…")
        for restitution_case in restitution_cases:
            address = [restitution_case[8]]
            if address not in addresses:
                addresses.append(address)

        # dockets
        ui_event("phase", "Retrieving dockets…")
        for address in addresses:
            if CANCEL_EVENT.is_set():
                ui_event("phase", "Stopping… writing partial CSV…")
                out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
                ui_event("done", out_path)
                return
            ui_log("Retrieving " + address[0])
            try:
                docket_response = requests.get(address[0], auth=(username, password), timeout=60)
            except Exception as e:
                ui_log(f"[WARN] Docket failed {address[0]}: {e}")
                address.extend(['could not retrieve'] * 4)
                continue
            docket_soup = BeautifulSoup(docket_response.content, 'lxml')
            docket_blocks = docket_soup.find_all('pre')
            num_pre_blocks = len(docket_blocks)
            if num_pre_blocks < 2:
                ui_log("Could not find docket party and address info in case at " + address[0])
                address.extend(['could not retrieve'] * 4)
            else:
                ui_log("Docket Party and Address Info" + docket_blocks[1].get_text())
                attorney_column_offset = docket_blocks[1].get_text().find("Attorney")
                ui_log("Client Info Ends at Text Column #" + str(attorney_column_offset))
                addresslines = docket_blocks[1].get_text().splitlines()
                addresslines_no_attys = list()
                if attorney_column_offset > 0:
                    for addressline in addresslines:
                        addresslines_no_attys.append(addressline[0:attorney_column_offset])
                addresslines = [line.strip() for line in addresslines_no_attys]
                ui_log("Extracted Client Info:")
                ui_log(str(addresslines))
                start_yet = 0
                defendant_count = 0
                current_line = -1
                for addressline in addresslines:
                    if CANCEL_EVENT.is_set():
                        ui_event("phase", "Stopping… writing partial CSV…")
                        out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
                        ui_event("done", out_path)
                        return
                    current_line += 1
                    if "Limited Representation Attorney" in addressline or " owes " in addressline or "Alias is " in addressline:
                        start_yet = 0
                    if "Defendant" in addressline:
                        start_yet = 1
                        defendant_count += 1
                    if start_yet == 1 and defendant_count == 1:
                        address.append(addressline)
                    if start_yet == 1 and defendant_count > 1:
                        if "Defendant" in addressline:
                            nxt = addresslines[current_line + 1] if current_line + 1 < len(addresslines) else ""
                            if ("ccupants" not in nxt and "CCUPANTS" not in nxt and
                                "ll other" not in nxt and "LL OTHER" not in nxt and
                                "ll Other" not in nxt and "John Doe" not in nxt and
                                "Jane Doe" not in nxt and "Real Name Unknown" not in nxt):
                                address[2] = address[2] + ", " + nxt

        if CANCEL_EVENT.is_set():
            ui_event("phase", "Stopping… writing partial CSV…")
            out_path = write_partial_csv(addresses, restitution_cases, targetDate, save_dir_var.get().strip() or desktop_folder())
            ui_event("done", out_path)
            return

        # tidy and write CSV (final full write)
        ui_event("phase", "Tidying rows…")
        for address in addresses:
            if (len(address) == 2):
                address.insert(1, " "); address.insert(2, " "); address.insert(3, " "); address.insert(4, " ")
            if (len(address) == 3):
                address.insert(2, " "); address.insert(3, " "); address.insert(4, " ")
            if (len(address) == 4):
                address.insert(3, " "); address.insert(4, " ")
            if (len(address) == 5):
                address.insert(4, " ")
            address[5] = " ".join(address[5].split())
            address.append(address[0][120:122] + "CI" + address[0][131:138])
            address.append(list(county_numbers_dict.keys())[list(county_numbers_dict.values()).index(address[0][94:96])])
            address.pop(1)
            address[2] = address[2] + " " + address[3]
            address.pop(3)
            address[2].rstrip(" ,")
        for address in addresses:
            address.pop(0)
            if len(address) == 6 and address[3] == "":
                address.pop(3)

        headers = ['name', 'address', 'city state zip', 'case number', 'county']
        addresses.insert(0, headers)
        filename = "eviction_cases_for_" + datetime.datetime.strptime(targetDate, '%m/%d/%Y').strftime('%Y-%m-%d') + "_generated_on_" + datetime.datetime.now().strftime('%Y-%m-%d-%H-%M') + ".csv"

        out_dir = save_dir_var.get().strip() or desktop_folder()
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception:
            pass
        out_path = os.path.join(out_dir, filename)

        ui_event("phase", "Writing CSV…")
        with open(out_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f, quoting=csv.QUOTE_ALL)
            writer.writerows(addresses)

        try:
            winsound.Beep(2500, 250)
        except Exception:
            pass

        ui_log("Done.")
        ui_event("done", out_path)

    except Exception as e:
        ui_log(f"[ERROR] {e}")
        ui_event("error", str(e))

# ---------------- UI ----------------
root = tk.Tk()
root.title(APP_NAME)

# --- window/taskbar icon fixes ---
try:
    # Preferred: PNG via iconphoto (affects title bar & taskbar on Win11)
    png_icon = resource_path(os.path.join("assets", "app_256.png"))
    if os.path.exists(png_icon):
        _img_icon = tk.PhotoImage(file=png_icon)
        root.iconphoto(True, _img_icon)
    else:
        # Fallback for older Tk: .ico via iconbitmap
        ico_icon = resource_path(os.path.join("assets", "app.ico"))
        if os.path.exists(ico_icon):
            root.iconbitmap(ico_icon)
except Exception:
    pass

# Theme (Azure if available)
style = load_azure_theme(root)

# Menu (Help -> About)
menubar = tk.Menu(root)
help_menu = tk.Menu(menubar, tearoff=0)

def open_url(url): webbrowser.open(url)

def show_about():
    top = tk.Toplevel(root)
    top.title("About")
    top.resizable(False, False)
    try:
        png_icon = resource_path(os.path.join("assets", "app_256.png"))
        if os.path.exists(png_icon):
            _about_img = tk.PhotoImage(file=png_icon)
            top.iconphoto(True, _about_img)
        else:
            ico_icon = resource_path(os.path.join("assets", "app.ico"))
            if os.path.exists(ico_icon):
                top.iconbitmap(ico_icon)
    except Exception:
        pass
    ttk.Label(top, text=APP_NAME, font=("Segoe UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(12,2))
    ttk.Label(top, text="License: CC0 — No Rights Reserved").grid(row=1, column=0, sticky="w", padx=12)
    ttk.Label(top, text=f"Credits: Alex Clark, {ORG_NAME}").grid(row=2, column=0, sticky="w", padx=12, pady=(0,2))  # UPDATED
    link_docs = ttk.Label(top, text="Support / Docs", foreground="#2563eb", cursor="hand2")
    link_docs.grid(row=3, column=0, sticky="w", padx=12, pady=(4,2))
    link_docs.bind("<Button-1>", lambda e: open_url(DOCS_URL))
    frame_links = ttk.Frame(top)
    frame_links.grid(row=4, column=0, sticky="w", padx=12)
    link_infi = ttk.Label(frame_links, text="InfiNet Solutions", foreground="#2563eb", cursor="hand2")
    link_infi.grid(row=0, column=0, sticky="w")
    ttk.Label(frame_links, text=", ").grid(row=0, column=1)
    link_alex = ttk.Label(frame_links, text=ALEX_NAME, foreground="#2563eb", cursor="hand2")
    link_alex.grid(row=0, column=2, sticky="w")
    link_infi.bind("<Button-1>", lambda e: open_url(ORG_URL))
    link_alex.bind("<Button-1>", lambda e: open_url(ALEX_URL))
    ttk.Button(top, text="Close", command=top.destroy).grid(row=5, column=0, sticky="e", padx=12, pady=(8,12))
    top.transient(root)
    top.grab_set()
    top.focus_set()

help_menu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=help_menu)
root.config(menu=menubar)

# Frames
date_frame   = ttk.LabelFrame(root, text="Target Date", padding=6)
options_frame= ttk.LabelFrame(root, text="Choose a County Option", padding=6)
cred_frame   = ttk.LabelFrame(root, text="Justice Login Credentials", padding=6)
save_frame   = ttk.LabelFrame(root, text="CSV Output Location", padding=6)
status_frame = ttk.LabelFrame(root, text="Status", padding=6)

# Date selector
label1 = ttk.Label(date_frame, text="Please select a date:")
entry1 = DateEntry(date_frame, width=12, background='darkblue',
                   foreground='white', borderwidth=2, date_pattern='mm/dd/yyyy')

# County radios
c_option = tk.StringVar(None, "1")
option1 = ttk.Radiobutton(options_frame, text="Douglas, Lancaster and Sarpy Only", variable=c_option, value="1")
option3 = ttk.Radiobutton(options_frame, text="Top 10 Counties", variable=c_option, value="3")
option2 = ttk.Radiobutton(options_frame, text="All Nebraska Counties", variable=c_option, value="2")

# Credentials + remember me
user_entry_label = ttk.Label(cred_frame, text="Username")
user_entry = ttk.Entry(cred_frame)
pass_entry_label = ttk.Label(cred_frame, text="Password")
pass_entry = ttk.Entry(cred_frame, show="*")
remember_var = tk.BooleanVar(value=False)
remember_chk = ttk.Checkbutton(cred_frame, text="Remember me (Windows-encrypted)", variable=remember_var)

# Load saved settings
_saved_remember, _saved_user, _saved_pass, _saved_dir = load_settings()
remember_var.set(_saved_remember)
if _saved_remember:
    if _saved_user:
        user_entry.insert(0, _saved_user)
    if _saved_pass:
        pass_entry.insert(0, _saved_pass)

# Save directory chooser
def desktop_folder():
    return os.path.join(os.path.expanduser("~"), "Desktop")

save_dir_var = tk.StringVar(value=_saved_dir if _saved_dir else desktop_folder())
save_dir_entry = ttk.Entry(save_frame, textvariable=save_dir_var)
def browse_dir():
    initial = save_dir_var.get().strip() or desktop_folder()
    folder = filedialog.askdirectory(initialdir=initial, title="Choose output folder")
    if folder:
        save_dir_var.set(folder)
        save_settings(remember_var.get(), user_entry.get(), pass_entry.get(), save_dir_var.get())
browse_btn = ttk.Button(save_frame, text="Browse…", command=browse_dir)

# Status log + progress + buttons
status_label = ttk.Label(status_frame, text="Idle.")
progress = ttk.Progressbar(status_frame, orient="horizontal", mode="indeterminate")
log_widget = ScrolledText(status_frame, height=12, width=100, state="disabled", wrap="word")

btn_start = ttk.Button(root, text="Start", width=12)
btn_stop  = ttk.Button(root, text="Stop",  width=12, state="disabled")
def open_folder():
    path = save_dir_var.get().strip() or desktop_folder()
    try:
        os.startfile(path)  # Windows
    except Exception:
        webbrowser.open(f"file://{path}")
btn_open  = ttk.Button(root, text="Open Folder", width=12, command=open_folder)

# Bottom-right link
link_label = ttk.Label(root, text="InfiNet Solutions", foreground="#2563eb", cursor="hand2")
link_label.bind("<Button-1>", lambda e: webbrowser.open(ORG_URL))

# Layout
date_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nw")
label1.grid(row=0, column=0, sticky="w")
entry1.grid(row=1, column=0, sticky="w")

options_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nw")
option1.grid(row=0, column=0, sticky="w")
option3.grid(row=1, column=0, sticky="w")
option2.grid(row=2, column=0, sticky="w")

cred_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nw")
user_entry_label.grid(row=0, column=0, sticky="w")
user_entry.grid(row=0, column=1, sticky="we", padx=(6,0))
pass_entry_label.grid(row=1, column=0, sticky="w", pady=(6,0))
pass_entry.grid(row=1, column=1, sticky="we", padx=(6,0), pady=(6,0))
remember_chk.grid(row=2, column=0, columnspan=2, sticky="w", pady=(6,0))
cred_frame.grid_columnconfigure(1, weight=1)

save_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=(0,10), sticky="we")
save_dir_entry.grid(row=0, column=0, sticky="we")
browse_btn.grid(row=0, column=1, padx=(6,0))
save_frame.grid_columnconfigure(0, weight=1)

btn_start.grid(row=2, column=0, sticky="w", padx=10, pady=(0,10))
btn_stop.grid(row=2, column=1, sticky="w", padx=10, pady=(0,10))
btn_open.grid(row=2, column=2, sticky="e", padx=10, pady=(0,10))

status_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=(0,10), sticky="nsew")
status_label.grid(row=0, column=0, sticky="w")
progress.grid(row=0, column=1, sticky="e", padx=(10,0))
log_widget.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(6,0))
status_frame.grid_rowconfigure(1, weight=1)
status_frame.grid_columnconfigure(1, weight=1)

link_label.grid(row=4, column=2, sticky="e", padx=12, pady=(0, 10))

# Make main window responsive
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_rowconfigure(3, weight=1)

# ---------- Start/Stop/plumbing ----------
progress_running = tk.BooleanVar(value=False)
WORKER = None

def set_inputs_enabled(enabled: bool):
    statez = "normal" if enabled else "disabled"
    entry1.config(state=statez)
    for rb in (option1, option2, option3):
        rb.config(state=statez)
    user_entry.config(state=statez)
    pass_entry.config(state=statez)
    save_dir_entry.config(state=statez)
    browse_btn.config(state=statez)
    remember_chk.config(state=statez)

def set_run_state(running: bool):
    set_inputs_enabled(not running)
    btn_start.config(state="disabled" if running else "normal")
    btn_stop.config(state="normal" if running else "disabled")
    if running:
        status_label.config(text="Starting…")
        progress.config(mode="indeterminate")
        if not progress_running.get():
            progress.start(50)
            progress_running.set(True)
    else:
        if progress_running.get():
            progress.stop()
            progress_running.set(False)
        status_label.config(text="Idle.")

def poll_queues():
    # logs
    try:
        while True:
            msg = LOG_QUEUE.get_nowait()
            log_widget.config(state="normal")
            log_widget.insert("end", msg + "\n")
            log_widget.see("end")
            log_widget.config(state="disabled")
    except queue.Empty:
        pass
    # events
    try:
        while True:
            kind, payload = EVENT_QUEUE.get_nowait()
            if kind == "phase":
                status_label.config(text=str(payload))
            elif kind == "done":
                status_label.config(text=f"Done. Saved file: {payload}")
                set_run_state(False)
            elif kind == "error":
                status_label.config(text=f"Error: {payload}")
                set_run_state(False)
                messagebox.showerror("Error", str(payload))
            elif kind == "canceled":
                status_label.config(text="Canceled.")
                ui_log("Operation canceled by user.")
                set_run_state(False)
    except queue.Empty:
        pass
    root.after(100, poll_queues)

def start_scrape():
    # clear log
    log_widget.config(state="normal"); log_widget.delete("1.0", "end"); log_widget.config(state="disabled")
    CANCEL_EVENT.clear()
    set_run_state(True)
    # launch worker
    global WORKER
    WORKER = threading.Thread(target=scrapeCalendar, daemon=True)
    WORKER.start()

def stop_scrape():
    if WORKER and WORKER.is_alive():
        ui_log("Stopping… will write partial CSV shortly.")
        status_label.config(text="Stopping…")
        CANCEL_EVENT.set()
        btn_stop.config(state="disabled")

btn_start.config(command=start_scrape)
btn_stop.config(command=stop_scrape)

root.after(100, poll_queues)

def on_close():
    save_settings(remember_var.get(), user_entry.get(), pass_entry.get(), save_dir_var.get())
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)

root.mainloop()
