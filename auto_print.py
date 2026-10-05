"""
Zeeva Auto Print - runs on PC2 (the PC with the EPSON L3250).
Watches a network folder and prints every new PNG / JPG / PDF on A4 glossy photo paper.
Landscape / portrait is detected for each image (and each PDF page).
After a successful print the file is MOVED to the Print-Done folder.
If a file with the same name is already there, it is renamed 1(1).png, 1(2).png ...

Shows a small icon in the Windows system tray (bottom-right, maybe under the ^ arrow):
green = running, orange = problem. Right-click it for status, folders, log and Exit.
Run it on PC2 (the PC that has the printer connected).

Install once:   pip install pywin32 pillow pymupdf pystray pyinstaller
Build exe:      pyinstaller --onefile --noconsole --collect-all pymupdf --hidden-import pystray._win32 auto_print.py
"""

import os
import re
import sys
import time
import ctypes
import shutil
import logging
import threading

import win32con
import win32gui
import win32print
import win32ui
import win32api
import win32event
import winerror
import pystray
from PIL import Image, ImageDraw, ImageOps, ImageWin

try:
    import pymupdf
except ImportError:
    import fitz as pymupdf

# ---------------- SETTINGS ----------------
WATCH_FOLDER = r"\\DESKTOP-68BR98S\Color_Print"
DONE_FOLDER = r"\\DESKTOP-68BR98S\Color_Print\Print-Done"
PRINTER_KEYWORD = "L3250"      # part of the printer name; falls back to default printer
POLL_SECONDS = 2               # how often the folder is checked
STABLE_SECONDS = 2             # file size must stay the same this long before printing
COPIES = 1
MEDIA_NAME = "Photo Paper Glossy"   # exact name as shown in the Paper type list
USE_PAPER_TYPE = False              # False = script does NOT choose paper type (printer's own default is used)
REQUIRE_GLOSSY = True               # only used when USE_PAPER_TYPE = True
PDF_DPI = 300
SUPPORTED = (".png", ".jpg", ".jpeg", ".pdf")

DMPAPER_A4 = 9
DM_MEDIATYPE = 0x00800000
# ------------------------------------------

LOG_FILE = os.path.join(
    os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)),
    "auto_print.log",
)
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
)


def _dev_caps():
    """The Windows printer library is winspool.drv (not winspool.dll)."""
    return ctypes.WinDLL("winspool.drv").DeviceCapabilitiesW


def is_color_printer(name):
    """False only when the driver says the printer is NOT a color printer."""
    DC_COLORDEVICE = 32
    try:
        return _dev_caps()(name, None, DC_COLORDEVICE, None, None) != 0
    except Exception:
        return True


def find_printer():
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    for p in win32print.EnumPrinters(flags):
        if PRINTER_KEYWORD.lower() in p[2].lower() and is_color_printer(p[2]):
            return p[2]
    raise RuntimeError("printer '%s' not found on this PC - run this on the PC with the printer" % PRINTER_KEYWORD)


def get_media_type_id(printer, port, wanted):
    """Ask the Epson driver for its paper type names and return the ID of `wanted`."""
    DC_MEDIATYPENAMES, DC_MEDIATYPES = 34, 35
    dc = _dev_caps()
    count = dc(printer, port, DC_MEDIATYPENAMES, None, None)
    if count <= 0:
        return None
    names_buf = ctypes.create_unicode_buffer(count * 64)
    ids_buf = (ctypes.c_uint32 * count)()
    dc(printer, port, DC_MEDIATYPENAMES, names_buf, None)
    dc(printer, port, DC_MEDIATYPES, ids_buf, None)
    raw = names_buf[:]
    names = [raw[i * 64:(i + 1) * 64].split("\x00")[0] for i in range(count)]
    logging.info("Paper types from driver: %s", list(zip(names, list(ids_buf))))
    for n, i in zip(names, ids_buf):
        if n.strip().lower() == wanted.lower():
            return i
    return None


def to_rgb(img):
    if img.mode in ("RGBA", "LA") or "transparency" in img.info:
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[3])
        return bg
    return img.convert("RGB")


def load_pages(path):
    """Return a list of PIL images: one for a picture, one per page for a PDF."""
    if path.lower().endswith(".pdf"):
        pages = []
        doc = pymupdf.open(path)
        try:
            for page in doc:
                pix = page.get_pixmap(dpi=PDF_DPI, alpha=False)
                pages.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        finally:
            doc.close()
        return pages
    img = Image.open(path)
    img.load()
    img = ImageOps.exif_transpose(img)
    img = to_rgb(img)
    img.info.pop("transparency", None)
    return [img]


def print_image(img, label):
    printer = find_printer()
    w, h = img.size
    landscape = w > h

    hprinter = win32print.OpenPrinter(printer)
    try:
        info = win32print.GetPrinter(hprinter, 2)
        devmode = info["pDevMode"]
        port = info["pPortName"]
    finally:
        win32print.ClosePrinter(hprinter)

    devmode.PaperSize = DMPAPER_A4
    devmode.FormName = "A4"
    devmode.Orientation = 2 if landscape else 1   # 2 = landscape, 1 = portrait
    devmode.Copies = COPIES
    devmode.Fields |= (
        win32con.DM_FORMNAME
        | win32con.DM_PAPERSIZE
        | win32con.DM_ORIENTATION
        | win32con.DM_COPIES
    )

    if USE_PAPER_TYPE:
        media_id = get_media_type_id(printer, port, MEDIA_NAME)
        if media_id is not None:
            devmode.MediaType = media_id
            devmode.Fields |= DM_MEDIATYPE
            logging.info("Using paper type '%s' (id %s)", MEDIA_NAME, media_id)
        else:
            logging.info("Paper type '%s' NOT found in driver list", MEDIA_NAME)
            if REQUIRE_GLOSSY:
                raise RuntimeError("paper type '%s' not found, print cancelled" % MEDIA_NAME)

    hdc = win32gui.CreateDC("WINSPOOL", printer, devmode)
    dc = win32ui.CreateDCFromHandle(hdc)
    try:
        dc.StartDoc(label)
        dc.StartPage()

        page_w = dc.GetDeviceCaps(win32con.HORZRES)
        page_h = dc.GetDeviceCaps(win32con.VERTRES)

        # fit inside page, keep aspect ratio, no cropping, centered
        scale = min(page_w / w, page_h / h)
        new_w, new_h = int(w * scale), int(h * scale)
        x = (page_w - new_w) // 2
        y = (page_h - new_h) // 2

        dib = ImageWin.Dib(img)
        dib.draw(dc.GetHandleOutput(), (x, y, x + new_w, y + new_h))

        dc.EndPage()
        dc.EndDoc()
    finally:
        dc.DeleteDC()

    logging.info("PRINTED %s (%s) on %s", label,
                 "landscape" if landscape else "portrait", printer)


def unique_destination(name):
    """Done-folder path for `name`; if taken, use name(1).ext, name(2).ext ..."""
    base, ext = os.path.splitext(name)
    dest = os.path.join(DONE_FOLDER, name)
    n = 1
    while os.path.exists(dest):
        dest = os.path.join(DONE_FOLDER, "%s(%d)%s" % (base, n, ext))
        n += 1
    return dest


def move_to_done(path):
    os.makedirs(DONE_FOLDER, exist_ok=True)
    dest = unique_destination(os.path.basename(path))
    shutil.move(path, dest)
    logging.info("MOVED %s -> %s", os.path.basename(path), dest)


def sort_key(name):
    m = re.match(r"(\d+)", name)
    return (0, int(m.group(1)), name.lower()) if m else (1, 0, name.lower())


def file_signature(path):
    st = os.stat(path)
    return (st.st_size, st.st_mtime)


def is_stable(path):
    try:
        s1 = file_signature(path)
        time.sleep(STABLE_SECONDS)
        s2 = file_signature(path)
        if s1 != s2 or s2[0] == 0:
            return False
        with open(path, "rb"):
            pass
        return True
    except OSError:
        return False


STATE = {"text": "Starting...", "ok": True, "last": "", "icon": None}


def make_icon_image(ok=True):
    color = (34, 160, 80) if ok else (230, 140, 20)
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((2, 2, 62, 62), radius=12, fill=color)
    d.rectangle((20, 10, 44, 28), fill="white")
    d.rectangle((10, 26, 54, 46), fill="white")
    d.rectangle((20, 40, 44, 56), fill="white", outline=color, width=2)
    return img


def set_status(text, ok=True):
    if STATE["text"] == text and STATE["ok"] == ok:
        return
    STATE["text"], STATE["ok"] = text, ok
    icon = STATE["icon"]
    if icon is not None:
        try:
            icon.title = ("Zeeva Auto Print - " + text)[:120]
            icon.icon = make_icon_image(ok)
            icon.update_menu()
        except Exception:
            pass


def idle_text():
    return "Running - waiting for files" + (" | " + STATE["last"] if STATE["last"] else "")


def watch_loop():
    logging.info("Started. Watching %s", WATCH_FOLDER)
    handled = {}   # filename -> signature (printed or failed; never retried unless file changes)

    while True:
        try:
            if not os.path.isdir(WATCH_FOLDER):
                set_status("Folder not reachable (is the other PC on?)", False)
                time.sleep(10)   # network folder not reachable yet
                continue
            set_status(idle_text(), STATE["ok"])

            names = [
                n for n in os.listdir(WATCH_FOLDER)
                if n.lower().endswith(SUPPORTED)
                and not n.startswith(("~", "."))
                and os.path.isfile(os.path.join(WATCH_FOLDER, n))
            ]

            # forget files that left the folder, so a new 1.png later prints again
            for n in list(handled):
                if n not in names:
                    del handled[n]

            for name in sorted(names, key=sort_key):
                path = os.path.join(WATCH_FOLDER, name)
                try:
                    sig = file_signature(path)
                except OSError:
                    continue
                if handled.get(name) == sig:
                    continue
                if not is_stable(path):
                    break   # keep order: wait for this file before the next ones

                try:
                    set_status("Printing " + name, True)
                    pages = load_pages(path)
                    for i, img in enumerate(pages, 1):
                        label = name if len(pages) == 1 else "%s (page %d)" % (name, i)
                        print_image(img, label)
                except Exception as e:
                    logging.info("ERROR printing %s: %s (file left in place)", name, e)
                    STATE["last"] = "ERROR with " + name
                    set_status("ERROR printing %s: %s" % (name, e), False)
                    handled[name] = sig   # do not retry forever
                    continue

                handled[name] = sig       # printed: never print this file twice
                STATE["last"] = "last printed: " + name
                set_status("Printed " + name, True)
                try:
                    move_to_done(path)
                except Exception as e:
                    logging.info("ERROR moving %s: %s (file left in place, will not reprint)", name, e)
        except Exception as e:
            logging.info("LOOP ERROR: %s", e)

        time.sleep(POLL_SECONDS)


_mutex = None


def already_running():
    """Only one copy at a time, so files are never printed twice."""
    global _mutex
    _mutex = win32event.CreateMutex(None, False, "ZeevaAutoPrintMutex")
    return win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS


def open_path(p):
    try:
        os.makedirs(p, exist_ok=True) if os.path.splitext(p)[1] == "" else None
        os.startfile(p)
    except Exception as e:
        logging.info("Could not open %s: %s", p, e)


def main():
    if already_running():
        return
    threading.Thread(target=watch_loop, daemon=True).start()
    menu = pystray.Menu(
        pystray.MenuItem(lambda item: STATE["text"], None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open Color_Print folder", lambda icon, item: open_path(WATCH_FOLDER)),
        pystray.MenuItem("Open Print-Done folder", lambda icon, item: open_path(DONE_FOLDER)),
        pystray.MenuItem("Open log file", lambda icon, item: open_path(LOG_FILE)),
        pystray.MenuItem("Exit", lambda icon, item: icon.stop()),
    )
    icon = pystray.Icon("ZeevaAutoPrint", make_icon_image(True),
                        "Zeeva Auto Print - starting", menu)
    STATE["icon"] = icon
    icon.run()


if __name__ == "__main__":
    main()
