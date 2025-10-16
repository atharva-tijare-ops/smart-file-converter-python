import os
import threading
import traceback
import time
from tkinter import filedialog, messagebox
import customtkinter as ctk
from PIL import Image
from pdf2docx import Converter
from fpdf import FPDF


try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except Exception:
    DOCX2PDF_AVAILABLE = False

# ---------- App appearance ----------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_TITLE = "Smart File Converter"
WINDOW_W = 640
WINDOW_H = 480

CONVERSIONS = [
    "JPG → PNG",
    "PNG → JPG",
    "PDF → DOCX",
    "DOCX → PDF",
    "TXT → PDF"
]

# ---------- Utilities ----------
def get_extension(path):
    return os.path.splitext(path)[1].lower().lstrip(".")

def safe_output_path(input_path, new_ext):
    base = os.path.splitext(input_path)[0]
    candidate = f"{base}.{new_ext}"
    counter = 1
    while os.path.exists(candidate):
        candidate = f"{base}({counter}).{new_ext}"
        counter += 1
    return candidate

# ---------- Converters ----------
def convert_image(input_path, target_ext, progress_callback=None):
    ext = get_extension(input_path)
    if ext not in ("jpg", "jpeg", "png"):
        raise ValueError("Image must be JPG/JPEG or PNG.")
    img = Image.open(input_path)
    # If targeting JPEG ensure RGB (drop alpha)
    if target_ext in ("jpg", "jpeg") and img.mode in ("RGBA", "LA"):
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        img = background
    output = safe_output_path(input_path, target_ext)
    img.save(output)
    if progress_callback:
        progress_callback(100)
    return output

def convert_pdf_to_docx(input_path, progress_callback=None):
    if get_extension(input_path) != "pdf":
        raise ValueError("Input must be a PDF.")
    output = safe_output_path(input_path, "docx")
    cv = Converter(input_path)
    try:
        # simple convert
        cv.convert(output)
    finally:
        cv.close()
    if progress_callback:
        progress_callback(100)
    return output

def convert_docx_to_pdf(input_path, progress_callback=None):
    if get_extension(input_path) not in ("docx",):
        raise ValueError("Input must be a .docx file.")
    if not DOCX2PDF_AVAILABLE:
        raise RuntimeError("DOCX→PDF requires `docx2pdf` and Microsoft Word (Windows/macOS). docx2pdf is not available.")
    output = safe_output_path(input_path, "pdf")
    # docx2pdf_convert can accept (input, output)
    docx2pdf_convert(input_path, output)
    if progress_callback:
        progress_callback(100)
    return output

def convert_txt_to_pdf(input_path, progress_callback=None):
    if get_extension(input_path) != "txt":
        raise ValueError("Input must be a .txt file.")
    output = safe_output_path(input_path, "pdf")
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    with open(input_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    total = max(len(lines), 1)
    for i, line in enumerate(lines, 1):
        pdf.multi_cell(0, 6, txt=line.rstrip())
        if progress_callback:
            progress_callback(int(i / total * 100))
    pdf.output(output)
    if progress_callback:
        progress_callback(100)
    return output

# ---------- UI ----------
class SmartConverter(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{WINDOW_W}x{WINDOW_H}")
        self.resizable(False, False)

        self.selected_conversion = ctk.StringVar(value=CONVERSIONS[0])
        self.file_path = None

        self._build_ui()

    def _build_ui(self):
        # Container with fixed width/height passed in constructor (fixed for CTk v5+)
        card_w = WINDOW_W - 40
        card_h = WINDOW_H - 40
        container = ctk.CTkFrame(self, corner_radius=12, width=card_w, height=card_h)
        container.place(relx=0.5, rely=0.5, anchor="center")

        # Title
        title = ctk.CTkLabel(container, text=APP_TITLE, font=ctk.CTkFont(size=26, weight="bold"))
        title.pack(pady=(18, 10))

        # Option menu (dropdown)
        self.option_menu = ctk.CTkOptionMenu(
            container,
            values=CONVERSIONS,
            variable=self.selected_conversion,
            command=self._on_option_change,
            width=360
        )
        self.option_menu.pack(pady=(6, 14))

        # File row (label + choose button)
        file_row = ctk.CTkFrame(container, fg_color="transparent")
        file_row.pack(fill="x", padx=28)

        self.file_label = ctk.CTkLabel(file_row, text="No file selected", anchor="w")
        self.file_label.pack(side="left", expand=True, fill="x", padx=(4, 8))

        choose_btn = ctk.CTkButton(file_row, text="Choose", width=100, command=self.choose_file)
        choose_btn.pack(side="right")

        # Big convert button
        self.convert_btn = ctk.CTkButton(container, text="Choose File & Convert", width=420, height=48, command=self._choose_or_convert)
        self.convert_btn.pack(pady=(16, 8))

        # Progress bar
        self.progress = ctk.CTkProgressBar(container, width=460)
        self.progress.set(0.0)
        self.progress.pack(pady=(8, 8))

        # Status
        self.status_label = ctk.CTkLabel(container, text="Status: Idle", anchor="w")
        self.status_label.pack(fill="x", padx=28, pady=(6, 6))

        # Footer
        footer = ctk.CTkLabel(container, text="Made By Atharva ❤ in Python", font=ctk.CTkFont(size=11))
        footer.pack(side="bottom", pady=(8, 12))

    def _on_option_change(self, _=None):
        # Reset selected file when conversion type changes
        self.file_path = None
        self.file_label.configure(text="No file selected")
        self.progress.set(0.0)
        self.status_label.configure(text="Status: Idle")

    def choose_file(self):
        conv = self.selected_conversion.get()
        types_map = {
            "JPG → PNG": [("JPEG / JPG", "*.jpg;*.jpeg")],
            "PNG → JPG": [("PNG", "*.png")],
            "PDF → DOCX": [("PDF", "*.pdf")],
            "DOCX → PDF": [("Word Documents", "*.docx")],
            "TXT → PDF": [("Text Files", "*.txt")],
        }
        filetypes = types_map.get(conv, [("All files", "*.*")])
        path = filedialog.askopenfilename(title="Select file", filetypes=filetypes)
        if path:
            self.file_path = path
            display = path if len(path) <= 60 else "..." + path[-57:]
            self.file_label.configure(text=display)
            self.status_label.configure(text=f"Selected: {os.path.basename(path)}")
            self.progress.set(0.0)

    def _choose_or_convert(self):
        # If no file chosen, open dialog
        if not self.file_path:
            self.choose_file()
            if not self.file_path:
                return
        self.start_conversion()

    def start_conversion(self):
        self.convert_btn.configure(state="disabled")
        self.status_label.configure(text="Status: Converting...")
        self.progress.set(0.02)
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        conv = self.selected_conversion.get()
        in_path = self.file_path
        try:
            # brief prep
            self.after(0, lambda: self.status_label.configure(text="Status: Preparing..."))
            time.sleep(0.12)

            def progress_cb(percent):
                self.after(0, lambda: self.progress.set(percent / 100.0))

            if conv == "JPG → PNG":
                out = convert_image(in_path, "png", progress_callback=progress_cb)
            elif conv == "PNG → JPG":
                out = convert_image(in_path, "jpg", progress_callback=progress_cb)
            elif conv == "PDF → DOCX":
                out = convert_pdf_to_docx(in_path, progress_callback=progress_cb)
            elif conv == "DOCX → PDF":
                if not DOCX2PDF_AVAILABLE:
                    raise RuntimeError("DOCX→PDF requires docx2pdf + Microsoft Word (Windows/macOS). It's not available in this environment.")
                out = convert_docx_to_pdf(in_path, progress_callback=progress_cb)
            elif conv == "TXT → PDF":
                out = convert_txt_to_pdf(in_path, progress_callback=progress_cb)
            else:
                raise ValueError("Unsupported conversion.")

            self.after(0, lambda: self.progress.set(1.0))
            self.after(0, lambda: self.status_label.configure(text=f"Status: Done — saved: {os.path.basename(out)}"))
            self.after(0, lambda: messagebox.showinfo("Success", f"Converted and saved:\n{out}"))
        except Exception as e:
            tb = traceback.format_exc()
            print("Conversion error:", tb)
            msg = str(e) or "Unknown error"
            self.after(0, lambda: messagebox.showerror("Conversion Failed", f"Error: {msg}"))
            self.after(0, lambda: self.status_label.configure(text="Status: Error"))
            self.after(0, lambda: self.progress.set(0.0))
        finally:
            self.after(0, lambda: self.convert_btn.configure(state="normal"))

# ---------- Run ----------
if __name__ == "__main__":
    app = SmartConverter()
    app.mainloop()