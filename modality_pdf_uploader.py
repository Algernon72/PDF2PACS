
# ---- Crash logging helpers (for PyInstaller --noconsole) ----
import sys, traceback, datetime as _dt
from pathlib import Path as _Path

def _app_dir() -> _Path:
    try:
        if getattr(sys, "frozen", False):
            return _Path(sys.executable).parent
        return _Path(__file__).parent
    except Exception:
        return _Path.cwd()

def _log_path() -> _Path:
    return _app_dir() / "modality_pdf_uploader.log"

def _crash_dialog(title: str, msg: str):
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        _r = _tk.Tk()
        _r.withdraw()
        _mb.showerror(title, msg)
        _r.destroy()
    except Exception:
        # last resort: print to stderr
        print(f"{title}: {msg}", file=sys.stderr)

def _safe_main(entrypoint):
    try:
        return entrypoint()
    except Exception as e:
        tb = traceback.format_exc()
        stamp = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        payload = f"[{stamp}] Unhandled exception:\\n{tb}\\n"
        try:
            _log_path().write_text(payload, encoding="utf-8")
        except Exception:
            pass
        _crash_dialog("Errore applicazione", "Si è verificato un errore non gestito.\\n"
                     "Dettagli salvati in 'modality_pdf_uploader.log' nella stessa cartella.\\n"
                     "Mostra il log a Stella per la diagnosi.")
        sys.exit(1)
# ---- end helpers ----

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Modality PDF Uploader → Orthanc (GUI, STOW-RS) — COMPATTO (no sezione PDF singolo)
----------------------------------------------------------------------------------
- Interfaccia riprogettata per stare tutta nello schermo senza scroll.
- Rimossa la sezione "Selezione PDF singolo": si utilizza solo l'elenco multiplo.
- Conservate tutte le altre funzionalità (multi-PDF, anteprima, WIA, Nuovo invio, impostazioni).
"""

import io
import os
import json
import datetime as dt
import traceback
from pathlib import Path
from typing import Tuple, Optional, List

import pydicom
from pydicom.dataset import FileDataset, Dataset
from pydicom.uid import ExplicitVRLittleEndian, generate_uid
from pydicom.filewriter import dcmwrite

try:
    import requests
except ImportError:
    import sys
    sys.stderr.write("Modulo 'requests' mancante. Installa con: pip install requests\n")
    sys.exit(1)

# --- Optional rendering: PDF → image (first page) ---
HAS_RENDER = False
try:
    import fitz  # PyMuPDF
    from PIL import Image  # only for save/encode if available
    HAS_RENDER = True
except Exception:
    HAS_RENDER = False

# --- Optional: WIA Scanning (Windows) ---
HAS_WIA = False
try:
    import platform
    if platform.system().lower() == "windows":
        import win32com.client  # pywin32
        from PIL import Image as PILImage
        HAS_WIA = True
except Exception:
    HAS_WIA = False

# --- GUI ---
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, font
    HAS_TK = True
except Exception:
    HAS_TK = False

APP_NAME = "ModalityPDFUploader"
CONFIG_FILE = "modality_uploader_config.json"

DEFAULT_CONFIG = {
    "stow": {
        "url": "http://127.0.0.1:8042/dicom-web/studies",
        "username": "",
        "password": "",
        "verify_tls": True,
        "timeout": 30
    },
    "defaults": {
        "StudyDescription": "Documenti",
        "SeriesDescription": "PDF Upload",
        "ReferringPhysicianName": "",
        "AccessionNumber": ""
    }
}


def parse_birth_date(human: str) -> str:
    s = (human or "").strip()
    if not s:
        return ""
    s = s.replace(".", "/").replace("-", "/").replace(" ", "")
    if len(s) == 8 and s.isdigit():
        return s
    parts = s.split("/")
    try:
        if len(parts) == 3:
            if len(parts[0]) == 4 and parts[0].isdigit():
                y, m, d = parts
            else:
                d, m, y = parts
            y, m, d = int(y), int(m), int(d)
            if 1 <= m <= 12 and 1 <= d <= 31 and 1800 <= y <= 2200:
                return f"{y:04d}{m:02d}{d:02d}"
    except Exception:
        pass
    return ""

def generate_pid(prefix: str = "ICCPV") -> str:
    import datetime as _dt
    return f"{prefix}{_dt.datetime.now().strftime('%Y%m%d%H%M%S')}"


def res_dir() -> Path:
    import sys
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def cfg_path() -> Path:
    return res_dir() / CONFIG_FILE

def load_cfg() -> dict:
    p = cfg_path()
    if not p.exists():
        save_cfg(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        save_cfg(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))

def save_cfg(cfg: dict) -> None:
    cfg_path().write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")

# ------------------------ Helpers ------------------------

def human_to_dicom_pn(cognome_nome: str) -> str:
    s = (cognome_nome or "").strip()
    if not s:
        return "Anon^Anon"
    if "^" in s:
        return s
    parts = s.split()
    if len(parts) == 1:
        return f"{parts[0]}^"
    cognome = parts[0]
    nome = " ".join(parts[1:])
    return f"{cognome}^{nome}"

def base_dataset(pn: str, pid: str, cfg: dict) -> FileDataset:
    now = dt.datetime.now()
    file_meta = Dataset()
    file_meta.FileMetaInformationGroupLength = 0
    file_meta.FileMetaInformationVersion = b"\x00\x01"
    file_meta.TransferSyntaxUID = ExplicitVRLittleEndian
    file_meta.ImplementationClassUID = "1.2.826.0.1.3680043.9.9999.2"
    file_meta.ImplementationVersionName = APP_NAME[:16]

    ds = FileDataset("", {}, file_meta=file_meta, preamble=b"\x00"*128)
    ds.is_little_endian = True
    ds.is_implicit_VR = False

    ds.PatientName = pn
    ds.PatientID = pid or ""
    ds.StudyDate = now.strftime("%Y%m%d")
    ds.StudyTime = now.strftime("%H%M%S")

    dflt = cfg.get("defaults", {})
    ds.StudyDescription = dflt.get("StudyDescription", "Documenti")
    ds.SeriesDescription = dflt.get("SeriesDescription", "PDF Upload")
    ds.ReferringPhysicianName = dflt.get("ReferringPhysicianName", "")
    ds.AccessionNumber = dflt.get("AccessionNumber", "")

    ds.StudyInstanceUID = generate_uid()
    ds.SeriesInstanceUID = generate_uid()
    ds.SOPInstanceUID = generate_uid()
    ds.SeriesNumber = 1
    ds.InstanceNumber = 1
    return ds

def build_encapsulated_pdf(pdf_path: Path, pn: str, pid: str, cfg: dict) -> FileDataset:
    ds = base_dataset(pn, pid, cfg)
    ds.SOPClassUID = "1.2.840.10008.5.1.4.1.1.104.1"  # Encapsulated PDF Storage
    ds.Modality = "DOC"
    blob = pdf_path.read_bytes()
    ds.MIMETypeOfEncapsulatedDocument = "application/pdf"
    ds.EncapsulatedDocument = blob
    # Helpful tags for viewer compatibility
    try:
        now = dt.datetime.now()
        ds.ContentDate = now.strftime("%Y%m%d")
        ds.ContentTime = now.strftime("%H%M%S")
        ds.InstanceCreationDate = now.strftime("%Y%m%d")
        ds.InstanceCreationTime = now.strftime("%H%M%S")
        ds.SpecificCharacterSet = "ISO_IR 100"
        ds.BurnedInAnnotation = "NO"
        # Title = file name without extension
        ds.DocumentTitle = pdf_path.stem
    except Exception:
        pass
    ds.file_meta.MediaStorageSOPClassUID = ds.SOPClassUID
    ds.file_meta.MediaStorageSOPInstanceUID = ds.SOPInstanceUID
    return ds


def build_sc_from_pdf_first_page(pdf_path: Path, pn: str, pid: str, cfg: dict) -> Optional[FileDataset]:
    """
    Render the first page of a PDF as a Secondary Capture DICOM.
    Returns a FileDataset or None if rendering isn't available.
    """
    if not HAS_RENDER:
        return None
    try:
        doc = fitz.open(pdf_path.as_posix())
        if doc.page_count < 1:
            doc.close()
            return None
        page = doc.load_page(0)
        mat = fitz.Matrix(2, 2)  # 2x scaling
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_bytes = pix.tobytes("png")
        doc.close()

        from PIL import Image as PILImage
        import io as _io
        im = PILImage.open(_io.BytesIO(img_bytes)).convert("RGB")

        ds = base_dataset(pn, pid, cfg)
        ds.SOPClassUID = "1.2.840.10008.5.1.4.1.1.7"  # Secondary Capture
        ds.Modality = "SC"
        ds.ImageType = ["DERIVED", "SECONDARY"]
        ds.ConversionType = "WSD"

        ds.Rows = im.height
        ds.Columns = im.width
        ds.SamplesPerPixel = 3
        ds.PhotometricInterpretation = "RGB"
        ds.PlanarConfiguration = 0
        ds.BitsAllocated = 8
        ds.BitsStored = 8
        ds.HighBit = 7
        ds.PixelRepresentation = 0
        ds.PixelData = im.tobytes()

        # Helpful tags for viewer compatibility
        try:
            now = dt.datetime.now()
            ds.InstanceCreationDate = now.strftime("%Y%m%d")
            ds.InstanceCreationTime = now.strftime("%H%M%S")
            ds.BurnedInAnnotation = "NO"
            ds.SpecificCharacterSet = "ISO_IR 100"
        except Exception:
            pass

        ds.file_meta.MediaStorageSOPClassUID = ds.SOPClassUID
        ds.file_meta.MediaStorageSOPInstanceUID = ds.SOPInstanceUID
        return ds
    except Exception:
        return None


def build_sc_from_pdf_all_pages(pdf_path: Path, pn: str, pid: str, cfg: dict) -> List[FileDataset]:
    """Render all pages of a PDF into a list of Secondary Capture DICOM instances."""
    out: List[FileDataset] = []
    if not HAS_RENDER:
        return out
    try:
        doc = fitz.open(pdf_path.as_posix())
        n = doc.page_count
        if n < 1:
            doc.close()
            return out
        from PIL import Image as PILImage
        import io as _io
        for page_index in range(n):
            page = doc.load_page(page_index)
            mat = fitz.Matrix(2, 2)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("png")

            im = PILImage.open(_io.BytesIO(img_bytes)).convert("RGB")

            ds = base_dataset(pn, pid, cfg)
            ds.SOPClassUID = "1.2.840.10008.5.1.4.1.1.7"  # Secondary Capture
            ds.Modality = "SC"
            ds.ImageType = ["DERIVED", "SECONDARY"]
            ds.ConversionType = "WSD"

            ds.Rows = im.height
            ds.Columns = im.width
            ds.SamplesPerPixel = 3
            ds.PhotometricInterpretation = "RGB"
            ds.PlanarConfiguration = 0
            ds.BitsAllocated = 8
            ds.BitsStored = 8
            ds.HighBit = 7
            ds.PixelRepresentation = 0
            ds.PixelData = im.tobytes()

            # Helpful tags
            try:
                now = dt.datetime.now()
                ds.InstanceCreationDate = now.strftime("%Y%m%d")
                ds.InstanceCreationTime = now.strftime("%H%M%S")
                ds.BurnedInAnnotation = "NO"
                ds.SpecificCharacterSet = "ISO_IR 100"
            except Exception:
                pass

            ds.file_meta.MediaStorageSOPClassUID = ds.SOPClassUID
            ds.file_meta.MediaStorageSOPInstanceUID = ds.SOPInstanceUID
            out.append(ds)
        doc.close()
        return out
    except Exception:
        return out

def dcm_to_bytes(ds: FileDataset) -> bytes:
    buff = io.BytesIO()
    dcmwrite(buff, ds, write_like_original=False)
    return buff.getvalue()

def _build_multipart_related(parts: List[bytes], boundary: str) -> bytes:
    pre = []
    for i, pb in enumerate(parts, start=1):
        pre.append(
            (
                f"--{boundary}\r\n"
                "Content-Type: application/dicom\r\n"
                f"Content-Location: instance-{i}\r\n"
                "\r\n"
            ).encode("utf-8")
        )
        pre.append(pb)
        pre.append(b"\r\n")
    pre.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(pre)

def stow_send(dicom_bytes: bytes, cfg: dict) -> Tuple[bool, str]:
    return stow_send_multi([dicom_bytes], cfg)

def stow_send_multi(dicom_bytes_list: List[bytes], cfg: dict) -> Tuple[bool, str]:
    import uuid
    url = (cfg["stow"]["url"] or "").rstrip("/")
    auth = None
    if cfg["stow"].get("username"):
        auth = (cfg["stow"]["username"], cfg["stow"].get("password", ""))
    verify = cfg["stow"].get("verify_tls", True)
    timeout = cfg["stow"].get("timeout", 30)

    boundary = "Boundary" + uuid.uuid4().hex
    body = _build_multipart_related(dicom_bytes_list, boundary)
    headers = {
        "Content-Type": f"multipart/related; type=\"application/dicom\"; boundary={boundary}",
        "Accept": "application/dicom+json, application/json;q=0.9, */*;q=0.1"
    }
    try:
        r = requests.post(url, data=body, headers=headers, auth=auth, verify=verify, timeout=timeout)
        if 200 <= r.status_code < 300:
            detail = ""
            try:
                resp = r.json()
                if isinstance(resp, dict):
                    succ = resp.get('Success', [])
                    fail = resp.get('Failed', [])
                    detail = f"Successi: {len(succ)}  Falliti: {len(fail)}"
                    if fail:
                        detail += f"  Motivi: {fail}"
                elif isinstance(resp, list):
                    tot_s = sum(len(x.get('Success', [])) for x in resp if isinstance(x, dict))
                    tot_f = sum(len(x.get('Failed', [])) for x in resp if isinstance(x, dict))
                    detail = f"Successi: {tot_s}  Falliti: {tot_f}"
            except Exception:
                detail = r.text[:300]
            return True, f"STOW OK {r.status_code} (inviati {len(dicom_bytes_list)} oggetti) — {detail}"
        else:
            frag = r.text
            if len(frag) > 900:
                frag = frag[:900] + "..."
            return False, f"STOW ERR {r.status_code}: {frag}"
    except Exception as e:
        return False, f"STOW EXC: {e}"

# ------------------------ Modern UI Styling ------------------------

class ModernStyle:
    PRIMARY = "#2563eb"
    PRIMARY_HOVER = "#1d4ed8"
    SUCCESS = "#059669"
    SUCCESS_HOVER = "#047857"
    WARNING = "#d97706"
    DANGER = "#dc2626"
    DANGER_HOVER = "#b91c1c"
    BG_PRIMARY = "#ffffff"
    BG_SECONDARY = "#f8fafc"
    BG_TERTIARY = "#e2e8f0"
    TEXT_PRIMARY = "#0f172a"
    TEXT_SECONDARY = "#64748b"
    BORDER = "#e2e8f0"
    BORDER_FOCUS = "#3b82f6"
    
    @staticmethod
    def configure_ttk_styles(root):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Card.TFrame', background=ModernStyle.BG_PRIMARY, relief='flat', borderwidth=1)
        style.configure('Heading.TLabel', background=ModernStyle.BG_PRIMARY, foreground=ModernStyle.TEXT_PRIMARY, font=('Segoe UI', 11, 'bold'))
        style.configure('TLabel', background=ModernStyle.BG_PRIMARY, foreground=ModernStyle.TEXT_PRIMARY, font=('Segoe UI', 9))
        style.configure('TEntry', fieldbackground=ModernStyle.BG_PRIMARY, borderwidth=1, relief='solid', insertcolor=ModernStyle.TEXT_PRIMARY, font=('Segoe UI', 9))
        style.configure('Primary.TButton', background=ModernStyle.PRIMARY, foreground='white', borderwidth=0, font=('Segoe UI', 9, 'bold'), relief='flat')
        style.map('Primary.TButton', background=[('active', ModernStyle.PRIMARY_HOVER), ('pressed', ModernStyle.PRIMARY_HOVER)])
        style.configure('Success.TButton', background=ModernStyle.SUCCESS, foreground='white', borderwidth=0, font=('Segoe UI', 9, 'bold'), relief='flat')
        style.map('Success.TButton', background=[('active', ModernStyle.SUCCESS_HOVER), ('pressed', ModernStyle.SUCCESS_HOVER)])
        style.configure('Danger.TButton', background=ModernStyle.DANGER, foreground='white', borderwidth=0, font=('Segoe UI', 9, 'bold'), relief='flat')
        style.map('Danger.TButton', background=[('active', ModernStyle.DANGER_HOVER), ('pressed', ModernStyle.DANGER_HOVER)])
        style.configure('Secondary.TButton', background=ModernStyle.BG_SECONDARY, foreground=ModernStyle.TEXT_PRIMARY, borderwidth=1, relief='solid', font=('Segoe UI', 9))
        style.configure('Modern.TLabelframe', background=ModernStyle.BG_PRIMARY, borderwidth=1, relief='solid', labelmargins=(10, 0, 0, 0))
        style.configure('Modern.TLabelframe.Label', background=ModernStyle.BG_PRIMARY, foreground=ModernStyle.TEXT_PRIMARY, font=('Segoe UI', 10, 'bold'))
        style.configure('TCheckbutton', background=ModernStyle.BG_PRIMARY, foreground=ModernStyle.TEXT_PRIMARY, font=('Segoe UI', 9))

# ------------------------ GUI ------------------------

class SettingsDialog(tk.Toplevel):
    def __init__(self, master, cfg: dict):
        super().__init__(master)
        self.title("⚙️ Impostazioni")
        self.resizable(False, False)
        self.cfg = cfg
        self.configure(bg=ModernStyle.BG_PRIMARY)
        main_frame = ttk.Frame(self, style='Card.TFrame', padding=16)
        main_frame.pack(fill="both", expand=True)
        ttk.Label(main_frame, text="Configurazione STOW-RS", style='Heading.TLabel').grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(main_frame, text="STOW URL").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=6)
        self.stow_url = tk.StringVar(value=cfg["stow"]["url"])
        ttk.Entry(main_frame, textvariable=self.stow_url, width=44).grid(row=1, column=1, sticky="w")
        ttk.Label(main_frame, text="Username").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=6)
        self.user = tk.StringVar(value=cfg["stow"].get("username", ""))
        ttk.Entry(main_frame, textvariable=self.user, width=28).grid(row=2, column=1, sticky="w")
        ttk.Label(main_frame, text="Password").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=6)
        self.pwd = tk.StringVar(value=cfg["stow"].get("password", ""))
        ttk.Entry(main_frame, textvariable=self.pwd, width=28, show="•").grid(row=3, column=1, sticky="w")
        self.verify = tk.BooleanVar(value=cfg["stow"].get("verify_tls", True))
        ttk.Checkbutton(main_frame, text="🔒 Verifica certificati TLS (https)", variable=self.verify).grid(row=4, column=1, sticky="w", pady=6)
        ttk.Label(main_frame, text="Timeout (s)").grid(row=5, column=0, sticky="e", padx=(0, 8), pady=6)
        self.timeout = tk.StringVar(value=str(cfg["stow"].get("timeout", 30)))
        ttk.Entry(main_frame, textvariable=self.timeout, width=6).grid(row=5, column=1, sticky="w")
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(btn_frame, text="❌ Annulla", command=self.destroy, style='Secondary.TButton').pack(side="right", padx=(8, 0))
        ttk.Button(btn_frame, text="✅ Salva", command=self._save, style='Success.TButton').pack(side="right")
        self.transient(master)
        self.grab_set()
        self.wait_visibility()
        x = master.winfo_x() + (master.winfo_width() // 2) - (self.winfo_reqwidth() // 2)
        y = master.winfo_y() + (master.winfo_height() // 2) - (self.winfo_reqheight() // 2)
        self.geometry(f"+{x}+{y}")

    def _save(self):
        try:
            self.cfg["stow"]["url"] = self.stow_url.get().strip() or "http://127.0.0.1:8042/dicom-web/studies"
            self.cfg["stow"]["username"] = self.user.get().strip()
            self.cfg["stow"]["password"] = self.pwd.get()
            self.cfg["stow"]["verify_tls"] = bool(self.verify.get())
            self.cfg["stow"]["timeout"] = int(self.timeout.get())
            save_cfg(self.cfg)
            self.destroy()
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Errore salvataggio impostazioni: {e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📄 PDF2PACS Modality Simulator for PDF storing to PACS - Danilo Savioni 2025")
        # dimensioni ottimizzate per schermi 1366×768 e superiori, senza scroll
        self.geometry("1200x720")
        self.minsize(1100, 650)
        self.configure(bg=ModernStyle.BG_SECONDARY)
        ModernStyle.configure_ttk_styles(self)
        self.cfg = load_cfg()
        self.file_list: List[Path] = []
        self._build_ui()

    def _build_ui(self):
        # Layout a griglia 2×2 (senza canvas/scrollbar)
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=0)  # riga top fissa
        root.rowconfigure(1, weight=1)  # riga file-list espandibile
        root.rowconfigure(2, weight=0)  # comandi
        root.rowconfigure(3, weight=1)  # log espandibile

        # === TOP: Dati Paziente (sx) + Opzioni invio (dx) ===
        patient_frame = ttk.LabelFrame(root, text="👤 Dati Paziente (DICOM)", style='Modern.TLabelframe', padding=12)
        patient_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))

        options_frame = ttk.LabelFrame(root, text="⚙️ Opzioni di invio", style='Modern.TLabelframe', padding=12)
        options_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))

        # Campi paziente compatti (3 righe)
        self.var_cognome = tk.StringVar()
        self.var_nome = tk.StringVar()
        self.var_birth = tk.StringVar()
        self.var_pid = tk.StringVar()

        r0 = ttk.Frame(patient_frame); r0.pack(fill="x", pady=2)
        ttk.Label(r0, text="Cognome", font=('Segoe UI', 9, 'bold')).pack(side="left")
        ttk.Entry(r0, textvariable=self.var_cognome, width=22, font=('Segoe UI', 10)).pack(side="left", padx=6)
        ttk.Label(r0, text="Nome", font=('Segoe UI', 9, 'bold')).pack(side="left")
        ttk.Entry(r0, textvariable=self.var_nome, width=22, font=('Segoe UI', 10)).pack(side="left", padx=6)

        r1 = ttk.Frame(patient_frame); r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Data nasc.", font=('Segoe UI', 9, 'bold')).pack(side="left")
        ttk.Entry(r1, textvariable=self.var_birth, width=14, font=('Segoe UI', 10)).pack(side="left", padx=6)
        ttk.Label(r1, text="(DD/MM/YYYY)", foreground=ModernStyle.TEXT_SECONDARY).pack(side="left")

        r2 = ttk.Frame(patient_frame); r2.pack(fill="x", pady=2)
        ttk.Label(r2, text="Patient ID", font=('Segoe UI', 9, 'bold')).pack(side="left")
        ttk.Entry(r2, textvariable=self.var_pid, width=22, font=('Segoe UI', 10)).pack(side="left", padx=6)
        ttk.Button(r2, text="🎲 Genera", command=self._gen_pid, style='Secondary.TButton').pack(side="left", padx=(6, 0))

        ttk.Button(patient_frame, text="🆕 Nuovo invio", command=self.clear_form, style='Primary.TButton').pack(anchor="e", pady=(6, 0))

        # Opzioni
        self.var_make_preview = tk.BooleanVar(value=True)
        preview_text = "🖼️ Invia anteprima come immagine (prima pagina)"
        state = ("normal" if HAS_RENDER else "disabled")
        if not HAS_RENDER:
            preview_text += "\n   ⚠️ Richiede 'pymupdf' + 'pillow'"
        ttk.Checkbutton(options_frame, text=preview_text, variable=self.var_make_preview, state=state).pack(anchor="w", pady=(0, 6))

        self.var_make_all_pages = tk.BooleanVar(value=True)
        all_pages_text = "📚 Crea immagini per TUTTE le pagine (consigliato per Stone Viewer)"
        ttk.Checkbutton(options_frame, text=all_pages_text, variable=self.var_make_all_pages, state=state).pack(anchor="w", pady=(0, 8))

        self.var_series_per_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="🏥 Serie separata per ogni PDF (compatibilità Stone)", variable=self.var_series_per_pdf).pack(anchor="w")

        # === FILE LIST ===
        multi_frame = ttk.LabelFrame(root, text="📚 Elenco PDF multiplo", style='Modern.TLabelframe', padding=12)
        multi_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(0, 8))

        cmd = ttk.Frame(multi_frame); cmd.pack(fill="x", pady=(0, 8))
        ttk.Button(cmd, text="➕ Aggiungi file", command=self.add_files, style='Success.TButton').pack(side="left", padx=(0, 6))
        ttk.Button(cmd, text="➖ Rimuovi selezionati", command=self.remove_selected, style='Danger.TButton').pack(side="left", padx=(0, 6))
        ttk.Button(cmd, text="🗑️ Svuota elenco", command=self.clear_list, style='Secondary.TButton').pack(side="left", padx=(0, 6))
        self.btn_scan = ttk.Button(cmd, text="🖨️ Scansiona... (WIA)", command=self.scan_wia, style='Primary.TButton')
        self.btn_scan.pack(side="right")
        if not HAS_WIA:
            self.btn_scan.state(["disabled"])

        list_container = ttk.Frame(multi_frame)
        list_container.pack(fill="both", expand=True)
        self.lst = tk.Listbox(list_container, height=8, selectmode="extended", font=('Segoe UI', 9),
                              bg=ModernStyle.BG_PRIMARY, fg=ModernStyle.TEXT_PRIMARY,
                              selectbackground=ModernStyle.PRIMARY, selectforeground='white',
                              borderwidth=1, relief='solid', highlightthickness=0)
        self.lst.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(list_container, orient="vertical", command=self.lst.yview)
        sb.pack(side="right", fill="y")
        self.lst.config(yscrollcommand=sb.set)

        # === COMANDI PRINCIPALI ===
        main_cmd = ttk.Frame(root)
        main_cmd.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(main_cmd, text="⚙️ Configurazione", command=self.open_settings, style='Secondary.TButton').pack(side="left")
        ttk.Button(main_cmd, text="🚀 Invia a Orthanc", command=self.send_all, style='Success.TButton').pack(side="right")

        # === LOG (compatto) ===
        log_frame = ttk.LabelFrame(root, text="📋 Log", style='Modern.TLabelframe', padding=8)
        log_frame.grid(row=3, column=0, columnspan=2, sticky="nsew")
        log_container = ttk.Frame(log_frame); log_container.pack(fill="both", expand=True)
        self.log = tk.Text(log_container, height=8, wrap="word", font=('Consolas', 9),
                           bg=ModernStyle.BG_SECONDARY, fg=ModernStyle.TEXT_PRIMARY,
                           borderwidth=1, relief='solid', highlightthickness=0)
        self.log.pack(side="left", fill="both", expand=True)
        lsb = ttk.Scrollbar(log_container, orient="vertical", command=self.log.yview)
        lsb.pack(side="right", fill="y")
        self.log.config(yscrollcommand=lsb.set)

        self._log("UI compatta caricata. Usa l'elenco multiplo per aggiungere anche un singolo PDF.")

    # ----------- File picking / list mgmt -----------

    def add_files(self):
        paths = filedialog.askopenfilenames(title="Seleziona uno o più file PDF", filetypes=[("File PDF","*.pdf"), ("Tutti i file","*.*")])
        if not paths:
            return
        added = 0
        for p in paths:
            pth = Path(p)
            if pth.exists() and pth.suffix.lower()==".pdf" and pth not in self.file_list:
                self.file_list.append(pth)
                self.lst.insert("end", f"📄 {pth.name}")
                added += 1
        if added:
            self._log(f"Aggiunti {added} file all'elenco")

    def remove_selected(self):
        sel = list(self.lst.curselection())
        sel.reverse()
        removed = 0
        for idx in sel:
            try:
                self.lst.delete(idx)
                if idx < len(self.file_list):
                    self.file_list.pop(idx)
                    removed += 1
            except Exception:
                pass
        if removed:
            self._log(f"Rimossi {removed} file dall'elenco")

    def clear_list(self):
        self.lst.delete(0, "end")
        n = len(self.file_list)
        self.file_list.clear()
        if n:
            self._log(f"Elenco svuotato ({n} elementi)")

    def open_settings(self):
        SettingsDialog(self, self.cfg)

    def _gen_pid(self):
        new_pid = generate_pid()
        self.var_pid.set(new_pid)
        self._log(f"Generato Patient ID: {new_pid}")
        
    def clear_form(self):
        self.var_cognome.set("")
        self.var_nome.set("")
        self.var_birth.set("")
        self.var_pid.set("")
        self.lst.delete(0, "end")
        self.file_list.clear()
        self._log("Modulo resettato. Elenco svuotato.")

    def _log(self, msg: str):
        import datetime as _dt
        ts = _dt.datetime.now().strftime("%H:%M:%S")
        self.log.insert("end", f"[{ts}] {msg}\n")
        self.log.see("end")
        self.update_idletasks()

    # ----------- Scansione WIA → PDF -----------

    def scan_wia(self):
        if not HAS_WIA:
            messagebox.showwarning(APP_NAME, "❌ Scansione WIA non disponibile su questo sistema.")
            return
        try:
            cd = win32com.client.Dispatch("WIA.CommonDialog")
            self._log("Apertura selezione dispositivo WIA...")
            device = cd.ShowSelectDevice()
            if device is None:
                self._log("Selezione scanner annullata")
                return

            scans_dir = res_dir() / "scans"
            scans_dir.mkdir(parents=True, exist_ok=True)
            tmp_dir = scans_dir / "_tmp"
            tmp_dir.mkdir(parents=True, exist_ok=True)

            images_paths: List[Path] = []
            page = 1
            self._log("Inizio acquisizione pagine...")
            
            while True:
                wia_img = cd.ShowAcquireImage()
                if wia_img is None:
                    if page == 1:
                        self._log("Nessuna immagine acquisita")
                    break
                ext = ".jpg"
                try:
                    fmt = getattr(wia_img, "FileExtension", None)
                    if fmt:
                        ext = f".{str(fmt).lower().strip('. ')}"
                except Exception:
                    pass
                tmp_img = tmp_dir / f"scan_page_{page:03d}{ext}"
                wia_img.SaveFile(tmp_img.as_posix())
                images_paths.append(tmp_img)
                self._log(f"Pagina {page} acquisita: {tmp_img.name}")
                if not messagebox.askyesno("Scansione multipagina", "Acquisire un'altra pagina per lo stesso PDF?", icon='question'):
                    break
                page += 1

            if not images_paths:
                return

            from PIL import Image as PILImage
            pil_imgs = [PILImage.open(p).convert("RGB") for p in images_paths]
            pdf_name = f"scan_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            pdf_path = scans_dir / pdf_name
            if len(pil_imgs) == 1:
                pil_imgs[0].save(pdf_path.as_posix(), "PDF", resolution=300.0)
            else:
                first, rest = pil_imgs[0], pil_imgs[1:]
                first.save(pdf_path.as_posix(), "PDF", save_all=True, append_images=rest, resolution=300.0)

            for im in pil_imgs:
                try: im.close()
                except Exception: pass
            for p in images_paths:
                try: p.unlink(missing_ok=True)
                except Exception: pass
            try: tmp_dir.rmdir()
            except Exception: pass

            self.file_list.append(pdf_path)
            self.lst.insert("end", f"📄 {pdf_path.name}")
            self._log(f"PDF creato da scansione: {pdf_path.name} (aggiunto all'elenco)")
            
        except Exception as e:
            self._log(f"Errore durante scansione WIA: {e}")
            self._log(traceback.format_exc())

    # ----------- Preparazione e invio -----------

    def _prepare_instances_for_pdf(self, pdf_path: Path, pn: str, pid: str, birth: str, study_uid: str, series_uid: str, make_preview: bool, series_number: int = 1) -> List[bytes]:
        out: List[bytes] = []
        # 1) Always include Encapsulated PDF first
        ds_pdf = build_encapsulated_pdf(pdf_path, pn, pid, self.cfg)
        ds_pdf.StudyInstanceUID = study_uid
        ds_pdf.SeriesInstanceUID = series_uid
        if birth:
            ds_pdf.PatientBirthDate = birth
        ds_pdf.SeriesNumber = series_number
        ds_pdf.InstanceNumber = 1
        out.append(dcm_to_bytes(ds_pdf))

        # 2) Optionally add images
        next_instance = 2
        if make_preview and HAS_RENDER:
            if getattr(self, 'var_make_all_pages', None) and bool(self.var_make_all_pages.get()):
                sc_list = build_sc_from_pdf_all_pages(pdf_path, pn, pid, self.cfg)
                for i, ds_sc in enumerate(sc_list, start=0):
                    ds_sc.StudyInstanceUID = study_uid
                    ds_sc.SeriesInstanceUID = series_uid
                    if birth:
                        ds_sc.PatientBirthDate = birth
                    ds_sc.SeriesNumber = series_number
                    ds_sc.InstanceNumber = next_instance + i
                    out.append(dcm_to_bytes(ds_sc))
            else:
                ds_sc = build_sc_from_pdf_first_page(pdf_path, pn, pid, self.cfg)
                if ds_sc is not None:
                    ds_sc.StudyInstanceUID = study_uid
                    ds_sc.SeriesInstanceUID = series_uid
                    if birth:
                        ds_sc.PatientBirthDate = birth
                    ds_sc.SeriesNumber = series_number
                    ds_sc.InstanceNumber = next_instance
                    out.append(dcm_to_bytes(ds_sc))
        return out

    def send_all(self):
        try:
            sources: List[Path] = list(self.file_list)
            if not sources:
                messagebox.showwarning("⚠️ Nessun file", "Aggiungi almeno un PDF all'elenco (anche singolo).")
                return

            cognome = (self.var_cognome.get() or '').strip()
            nome = (self.var_nome.get() or '').strip()
            pn = human_to_dicom_pn(f"{cognome} {nome}".strip())
            pid = self.var_pid.get().strip() or generate_pid()
            birth = parse_birth_date(self.var_birth.get())
            make_preview = bool(self.var_make_preview.get())

            self._log("Avvio invio...")
            self._log(f"Paziente: {pn}  (ID: {pid})")
            self._log(f"File da processare: {len(sources)}")

            study_uid = generate_uid()
            per_series = bool(self.var_series_per_pdf.get())

            all_parts: List[bytes] = []
            
            if per_series:
                self._log("Modalità: serie separata per ogni PDF")
                series_number = 1
                for pdf in sources:
                    series_uid = generate_uid()
                    self._log(f"Processo: {pdf.name} → Serie #{series_number}")
                    parts = self._prepare_instances_for_pdf(pdf, pn, pid, birth, study_uid, series_uid, make_preview, series_number=series_number)
                    all_parts.extend(parts)
                    self._log(f"   Creati {len(parts)} oggetti DICOM per Serie #{series_number}")
                    series_number += 1
            else:
                self._log("Modalità: tutti i PDF nella stessa serie")
                series_uid = generate_uid()
                series_number = 1
                for pdf in sources:
                    self._log(f"Processo: {pdf.name}")
                    parts = self._prepare_instances_for_pdf(pdf, pn, pid, birth, study_uid, series_uid, make_preview, series_number=series_number)
                    all_parts.extend(parts)
                    self._log(f"   Creati {len(parts)} oggetti DICOM")

            self._log(f"Invio {len(all_parts)} oggetti a Orthanc...")
            ok, msg = stow_send_multi(all_parts, self.cfg)
            
            if ok:
                self._log(f"SUCCESSO! {msg}")
                messagebox.showinfo("✅ Invio completato", f"Inviati con successo {len(all_parts)} oggetti.\n\n{msg}")
            else:
                self._log(f"ERRORE: {msg}")
                messagebox.showerror("❌ Errore invio", f"L'invio non è riuscito:\n\n{msg}")

            self._log("— Fine —")
            
        except Exception as e:
            error_msg = f"Errore imprevisto: {e}"
            self._log(error_msg)
            self._log(traceback.format_exc())
            messagebox.showerror("❌ Errore", error_msg)

def main():
    if not HAS_TK:
        print("❌ Tkinter non disponibile. Installa Python con supporto Tcl/Tk.")
        return
        
    app = App()
    app.update_idletasks()
    # centra
    width = app.winfo_width()
    height = app.winfo_height()
    x = (app.winfo_screenwidth() // 2) - (width // 2)
    y = (app.winfo_screenheight() // 2) - (height // 2)
    app.geometry(f"{width}x{height}+{x}+{y}")
    app.mainloop()

if __name__ == "__main__":
    _safe_main(lambda: (
            main()
    ))
