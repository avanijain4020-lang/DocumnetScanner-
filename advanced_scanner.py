import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import qrcode
import tempfile
import os
import socket
import threading
from http.server import SimpleHTTPRequestHandler, HTTPServer

# PDF & PPT Handling Libraries
import fitz  # PyMuPDF for PDF processing
from pptx import Presentation  # python-pptx
import io

# --- Local Web Server for Multi-Page QR Document Viewing ---

SERVER_DIR = tempfile.mkdtemp(prefix="scanner_web_")
PORT = 8000

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

class ImageServerHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=SERVER_DIR, **kwargs)

def start_local_server():
    try:
        server = HTTPServer(('0.0.0.0', PORT), ImageServerHandler)
        server.serve_forever()
    except Exception as e:
        print(f"Server Error: {e}")

# Run Web Server in Background Thread
server_thread = threading.Thread(target=start_local_server, daemon=True)
server_thread.start()

# --- Computer Vision Functions ---

def order_points(pts):
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect

def four_point_transform(image, pts):
    rect = order_points(pts)
    (tl, tr, br, bl) = rect
    widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    maxWidth = max(int(widthA), int(widthB))

    heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    maxHeight = max(int(heightA), int(heightB))

    dst = np.array([[0, 0], [maxWidth - 1, 0], [maxWidth - 1, maxHeight - 1], [0, maxHeight - 1]], dtype="float32")
    M = cv2.getPerspectiveTransform(rect, dst)
    return cv2.warpPerspective(image, M, (maxWidth, maxHeight))

def detect_edges_and_corners(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 50, 150)
    cnts, _ = cv2.findContours(edged, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]
    for c in cnts:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4:
            return approx.reshape(4, 2)
    h, w = image.shape[:2]
    return np.array([[10, 10], [w - 10, 10], [w - 10, h - 10], [10, h - 10]], dtype="float32")

def apply_filters(warped_img, filter_type="B&W Document Scan"):
    if len(warped_img.shape) == 3:
        gray = cv2.cvtColor(warped_img, cv2.COLOR_BGR2GRAY)
        rgb = cv2.cvtColor(warped_img, cv2.COLOR_BGR2RGB)
    else:
        gray = warped_img.copy()
        rgb = cv2.cvtColor(warped_img, cv2.COLOR_GRAY2RGB)

    if filter_type == "Original Image":
        return rgb

    elif filter_type == "Magic Color":
        enhanced = cv2.convertScaleAbs(rgb, alpha=1.2, beta=10)
        kernel = np.array([[0, -0.5, 0], [-0.5, 3, -0.5], [0, -0.5, 0]])
        return cv2.filter2D(enhanced, -1, kernel)

    elif filter_type == "B&W Document Scan":
        bg = cv2.GaussianBlur(gray, (0, 0), sigmaX=15, sigmaY=15)
        divided = cv2.divide(gray, bg, scale=255)
        scan_bw = cv2.convertScaleAbs(divided, alpha=1.3, beta=-25)
        return scan_bw

    elif filter_type == "Smooth Grayscale":
        return cv2.convertScaleAbs(gray, alpha=1.15, beta=5)

    return rgb

# --- Batch Page Selection Popup Dialog ---

class PageSelectionDialog:
    def __init__(self, parent, total_pages, doc_type="Page"):
        self.top = tk.Toplevel(parent)
        self.top.title(f"Batch Import - Select {doc_type}s")
        self.top.geometry("380x250")
        self.top.resizable(False, False)
        self.top.transient(parent)
        self.top.grab_set()

        self.total_pages = total_pages
        self.pages_to_import = []

        tk.Label(self.top, text=f"Total {doc_type}s Found: {total_pages}", font=("Arial", 11, "bold")).pack(pady=10)

        self.choice_var = tk.StringVar(value="all")

        rb_all = tk.Radiobutton(self.top, text=f"Import All {doc_type}s (1 - {total_pages})", variable=self.choice_var, value="all", command=self.toggle_entries)
        rb_all.pack(anchor=tk.W, padx=30, pady=5)

        rb_range = tk.Radiobutton(self.top, text="Import Specific Range:", variable=self.choice_var, value="range", command=self.toggle_entries)
        rb_range.pack(anchor=tk.W, padx=30, pady=5)

        range_frame = tk.Frame(self.top)
        range_frame.pack(pady=5)

        tk.Label(range_frame, text="From: ").pack(side=tk.LEFT)
        self.start_entry = tk.Entry(range_frame, width=5)
        self.start_entry.insert(0, "1")
        self.start_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(range_frame, text="To: ").pack(side=tk.LEFT)
        self.end_entry = tk.Entry(range_frame, width=5)
        self.end_entry.insert(0, str(total_pages))
        self.end_entry.pack(side=tk.LEFT, padx=5)

        self.toggle_entries()

        btn_frame = tk.Frame(self.top)
        btn_frame.pack(pady=15)
        tk.Button(btn_frame, text="Import", command=self.on_ok, bg="#10b981", fg="white", font=("Arial", 10, "bold"), width=10).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Cancel", command=self.on_cancel, bg="#ef4444", fg="white", font=("Arial", 10), width=10).pack(side=tk.LEFT, padx=10)

    def toggle_entries(self):
        state = tk.NORMAL if self.choice_var.get() == "range" else tk.DISABLED
        self.start_entry.config(state=state)
        self.end_entry.config(state=state)

    def on_ok(self):
        if self.choice_var.get() == "all":
            self.pages_to_import = list(range(self.total_pages))
        else:
            try:
                start = int(self.start_entry.get())
                end = int(self.end_entry.get())
                if 1 <= start <= end <= self.total_pages:
                    self.pages_to_import = list(range(start - 1, end))
                else:
                    messagebox.showerror("Invalid Range", f"Range 1 se {self.total_pages} ke beech honi chahiye!", parent=self.top)
                    return
            except ValueError:
                messagebox.showerror("Invalid Input", "Kripya valid numbers enter karein!", parent=self.top)
                return
        self.top.destroy()

    def on_cancel(self):
        self.pages_to_import = []
        self.top.destroy()

# --- GUI Application ---

class ProfessionalScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Multi-Format Batch Scanner & Full Document QR Viewer")
        self.root.geometry("1180x760")
        self.raw_warped_pages = []  # Unfiltered warped pages for re-filtering
        self.scanned_pages = []      # Filtered pages
        self.current_raw_image = None
        self.current_points = None
        self.current_qr_img = None
        self.setup_ui()

    def setup_ui(self):
        title_frame = tk.Frame(self.root, bg="#1e293b", pady=10)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="Advanced Multi-Format Batch Document Scanner & QR Engine", font=("Arial", 16, "bold"), fg="white", bg="#1e293b").pack()

        main_pane = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bd=2, relief="groove")
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left Editor Frame
        self.editor_frame = tk.Frame(main_pane, bg="#f1f5f9")
        main_pane.add(self.editor_frame, width=720)

        self.canvas = tk.Canvas(self.editor_frame, bg="#cbd5e1", width=650, height=480)
        self.canvas.pack(pady=10, padx=10)

        filter_frame = tk.Frame(self.editor_frame, bg="#f1f5f9")
        filter_frame.pack(pady=5)
        tk.Label(filter_frame, text="Filter: ", font=("Arial", 10, "bold"), bg="#f1f5f9").pack(side=tk.LEFT)
        self.filter_var = tk.StringVar(value="B&W Document Scan")
        ttk.Combobox(
            filter_frame,
            textvariable=self.filter_var,
            values=["Original Image", "Magic Color", "B&W Document Scan", "Smooth Grayscale"],
            state="readonly",
            width=18
        ).pack(side=tk.LEFT, padx=3)

        tk.Button(filter_frame, text="Apply Current Page", command=self.add_processed_page, bg="#2563eb", fg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(filter_frame, text="Apply to ALL Pages", command=self.apply_filter_to_all_pages, bg="#0284c7", fg="white", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # Right Sidebar Frame
        self.sidebar_frame = tk.Frame(main_pane, bg="#ffffff")
        main_pane.add(self.sidebar_frame, width=420)

        tk.Label(self.sidebar_frame, text="Scanned Pages Manager", font=("Arial", 12, "bold"), bg="#ffffff").pack(pady=5)
        self.pages_listbox = tk.Listbox(self.sidebar_frame, font=("Arial", 10), height=8)
        self.pages_listbox.pack(fill=tk.X, padx=10, pady=5)
        self.pages_listbox.bind('<<ListboxSelect>>', self.on_page_select)

        # QR Display Section
        tk.Label(self.sidebar_frame, text="Full Document QR Code Viewer", font=("Arial", 9, "bold"), bg="#ffffff").pack(pady=(5, 2))
        self.qr_canvas = tk.Canvas(self.sidebar_frame, bg="#f8fafc", width=180, height=180, highlightthickness=1, highlightbackground="#cbd5e1")
        self.qr_canvas.pack(pady=2)

        btn_box = tk.Frame(self.sidebar_frame, bg="#ffffff")
        btn_box.pack(fill=tk.X, pady=5, padx=10)

        tk.Button(btn_box, text="Import File (Image / PDF / PPT)", command=self.import_file, bg="#10b981", fg="white", font=("Arial", 10, "bold")).pack(fill=tk.X, pady=2)
        tk.Button(btn_box, text="Generate Full Document QR Code", command=self.generate_full_document_qr, bg="#8b5cf6", fg="white", font=("Arial", 10, "bold")).pack(fill=tk.X, pady=2)
        tk.Button(btn_box, text="Save Generated QR Code", command=self.save_qr_code, bg="#d97706", fg="white", font=("Arial", 9, "bold")).pack(fill=tk.X, pady=2)
        tk.Button(btn_box, text="Remove Selected Page", command=self.remove_selected_page, bg="#ef4444", fg="white", font=("Arial", 9)).pack(fill=tk.X, pady=2)
        tk.Button(btn_box, text="Export Multi-Page PDF", command=self.export_pdf, bg="#0f172a", fg="white", font=("Arial", 10, "bold")).pack(fill=tk.X, pady=6)

    def import_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("All Supported Files", "*.jpg *.png *.jpeg *.bmp *.pdf *.pptx *.ppt"),
                ("Image Files", "*.jpg *.png *.jpeg *.bmp"),
                ("PDF Documents", "*.pdf"),
                ("PowerPoint Presentations", "*.pptx *.ppt")
            ]
        )
        if not file_path:
            return

        ext = os.path.splitext(file_path)[1].lower()

        if ext in ['.jpg', '.png', '.jpeg', '.bmp']:
            self.load_image_file(file_path)
        elif ext == '.pdf':
            self.load_pdf_file_batch(file_path)
        elif ext in ['.pptx', '.ppt']:
            self.load_ppt_file_batch(file_path)

    def load_image_file(self, file_path):
        image = cv2.imread(file_path)
        if image is None:
            messagebox.showerror("Error", "Failed to load image!")
            return
        self.set_current_working_image(image)

    def load_pdf_file_batch(self, file_path):
        try:
            pdf_doc = fitz.open(file_path)
            total_pages = len(pdf_doc)
            if total_pages == 0:
                messagebox.showerror("Error", "PDF file is empty!")
                return

            dialog = PageSelectionDialog(self.root, total_pages, doc_type="PDF Page")
            self.root.wait_window(dialog.top)

            if not dialog.pages_to_import:
                return

            filter_choice = self.filter_var.get()
            added_count = 0

            for page_num in dialog.pages_to_import:
                page = pdf_doc[page_num]
                pix = page.get_pixmap(dpi=150)
                img_data = np.frombuffer(pix.samples, dtype=np.uint8)

                if pix.alpha:
                    img_data = img_data.reshape(pix.height, pix.width, 4)
                    image = cv2.cvtColor(img_data, cv2.COLOR_RGBA2BGR)
                else:
                    image = img_data.reshape(pix.height, pix.width, 3)
                    image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

                pts = detect_edges_and_corners(image)
                warped = four_point_transform(image, pts)
                final_processed = apply_filters(warped, filter_choice)

                self.raw_warped_pages.append(warped)
                self.scanned_pages.append(final_processed)

                page_idx = len(self.scanned_pages)
                self.pages_listbox.insert(tk.END, f"Page {page_idx} - PDF p.{page_num+1} [{filter_choice}]")
                added_count += 1

            if added_count > 0:
                self.pages_listbox.selection_clear(0, tk.END)
                self.pages_listbox.selection_set(tk.END)
                self.pages_listbox.activate(tk.END)
                self.display_image_on_canvas(self.scanned_pages[-1])
                messagebox.showinfo("Batch Import Success", f"Successfully imported & processed {added_count} PDF page(s)!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read PDF: {str(e)}")

    def load_ppt_file_batch(self, file_path):
        try:
            prs = Presentation(file_path)
            slides_images = []

            for i, slide in enumerate(prs.slides):
                slide_imgs = []
                for shape in slide.shapes:
                    if hasattr(shape, "image"):
                        image_bytes = shape.image.blob
                        pil_img = Image.open(io.BytesIO(image_bytes))
                        np_img = np.array(pil_img)
                        slide_imgs.append(np_img)
                if slide_imgs:
                    slides_images.append((i, slide_imgs[0]))

            if not slides_images:
                messagebox.showwarning("Warning", "PPT file me koi embedded image nahi mili!")
                return

            dialog = PageSelectionDialog(self.root, len(slides_images), doc_type="PPT Slide Image")
            self.root.wait_window(dialog.top)

            if not dialog.pages_to_import:
                return

            filter_choice = self.filter_var.get()
            added_count = 0

            for idx in dialog.pages_to_import:
                slide_num, raw_img = slides_images[idx]
                if len(raw_img.shape) == 3 and raw_img.shape[2] == 4:
                    image = cv2.cvtColor(raw_img, cv2.COLOR_RGBA2BGR)
                elif len(raw_img.shape) == 3:
                    image = cv2.cvtColor(raw_img, cv2.COLOR_RGB2BGR)
                else:
                    image = cv2.cvtColor(raw_img, cv2.COLOR_GRAY2BGR)

                pts = detect_edges_and_corners(image)
                warped = four_point_transform(image, pts)
                final_processed = apply_filters(warped, filter_choice)

                self.raw_warped_pages.append(warped)
                self.scanned_pages.append(final_processed)

                page_idx = len(self.scanned_pages)
                self.pages_listbox.insert(tk.END, f"Page {page_idx} - PPT s.{slide_num+1} [{filter_choice}]")
                added_count += 1

            if added_count > 0:
                self.pages_listbox.selection_clear(0, tk.END)
                self.pages_listbox.selection_set(tk.END)
                self.pages_listbox.activate(tk.END)
                self.display_image_on_canvas(self.scanned_pages[-1])
                messagebox.showinfo("Batch Import Success", f"Successfully imported {added_count} slide image(s) from PPT!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read PPT: {str(e)}")

    def set_current_working_image(self, image):
        self.current_raw_image = image
        self.current_points = detect_edges_and_corners(image)

        preview_img = image.copy()
        pts = self.current_points.reshape((-1, 1, 2)).astype(np.int32)
        cv2.polylines(preview_img, [pts], isClosed=True, color=(0, 0, 255), thickness=5)

        self.display_image_on_canvas(preview_img)

    def display_image_on_canvas(self, img):
        h, w = img.shape[:2]
        scale = min(650 / w, 480 / h)
        new_w, new_h = max(1, int(w * scale)), max(1, int(h * scale))
        resized = cv2.resize(img, (new_w, new_h))

        if len(resized.shape) == 2:
            rgb = cv2.cvtColor(resized, cv2.COLOR_GRAY2RGB)
        else:
            rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB) if img.ndim == 3 and img.shape[2] == 3 else resized

        self.preview_pil = Image.fromarray(rgb)
        self.preview_tk = ImageTk.PhotoImage(self.preview_pil)
        self.canvas.delete("all")
        self.canvas.create_image(325, 240, image=self.preview_tk, anchor=tk.CENTER)

    def add_processed_page(self):
        if self.current_raw_image is None:
            messagebox.showwarning("Warning", "Pehle ek File Import karein!")
            return

        warped = four_point_transform(self.current_raw_image, self.current_points)
        filter_choice = self.filter_var.get()
        final_processed = apply_filters(warped, filter_choice)

        self.raw_warped_pages.append(warped)
        self.scanned_pages.append(final_processed)

        page_num = len(self.scanned_pages)
        self.pages_listbox.insert(tk.END, f"Page {page_num} - Single [{filter_choice}]")

        self.pages_listbox.selection_clear(0, tk.END)
        self.pages_listbox.selection_set(tk.END)
        self.pages_listbox.activate(tk.END)

        self.display_image_on_canvas(final_processed)
        messagebox.showinfo("Success", f"Page {page_num} Scanned & Added!")

    def apply_filter_to_all_pages(self):
        if not self.scanned_pages:
            messagebox.showwarning("Warning", "Filter apply karne ke liye koi pages add nahi hain!")
            return

        filter_choice = self.filter_var.get()
        self.scanned_pages.clear()

        # Update Listbox & re-apply filter to all original warped pages
        for i, warped in enumerate(self.raw_warped_pages):
            re_processed = apply_filters(warped, filter_choice)
            self.scanned_pages.append(re_processed)

            # Preserve original label format in listbox
            old_label = self.pages_listbox.get(i)
            base_label = old_label.split('[')[0].strip() if '[' in old_label else old_label.split('-')[0].strip()
            self.pages_listbox.delete(i)
            self.pages_listbox.insert(i, f"{base_label} [{filter_choice}]")

        # Show current selection or last page preview
        selected = self.pages_listbox.curselection()
        active_idx = selected[0] if selected else len(self.scanned_pages) - 1

        self.pages_listbox.selection_clear(0, tk.END)
        self.pages_listbox.selection_set(active_idx)
        self.display_image_on_canvas(self.scanned_pages[active_idx])

        messagebox.showinfo("Bulk Filter Applied", f"'{filter_choice}' filter saare {len(self.scanned_pages)} pages par apply ho gaya hai!")

    def on_page_select(self, event):
        selected = self.pages_listbox.curselection()
        if selected:
            idx = selected[0]
            if idx < len(self.scanned_pages):
                self.display_image_on_canvas(self.scanned_pages[idx])

    def generate_full_document_qr(self):
        if not self.scanned_pages:
            messagebox.showwarning("Warning", "QR Code generate karne ke liye koi pages add nahi hain!")
            return

        try:
            # 1. Clear & save all pages into the server directory
            image_tags_html = ""
            total_pages = len(self.scanned_pages)

            for idx, selected_img in enumerate(self.scanned_pages):
                filename = f"doc_page_{idx + 1}.jpg"
                server_filepath = os.path.join(SERVER_DIR, filename)

                save_img = cv2.cvtColor(selected_img, cv2.COLOR_RGB2BGR) if len(selected_img.shape) == 3 else selected_img
                cv2.imwrite(server_filepath, save_img)

                image_tags_html += f"""
                <div class="page-card">
                    <img src="{filename}" alt="Page {idx + 1}">
                    <div class="page-num">Page {idx + 1} of {total_pages}</div>
                </div>
                """

            # 2. Build full document HTML reader page
            html_content = f"""
            <!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Scanned Multi-Page Document</title>
                <style>
                    body {{ background-color: #0f172a; margin: 0; padding: 15px; font-family: -apple-system, BlinkMacSystemFont, sans-serif; color: white; text-align: center; }}
                    .header {{ margin-bottom: 20px; font-size: 1.1rem; color: #38bdf8; font-weight: bold; }}
                    .page-card {{ margin: 0 auto 20px auto; max-width: 600px; background: white; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.4); padding: 8px; }}
                    img {{ width: 100%; height: auto; border-radius: 4px; display: block; }}
                    .page-num {{ color: #475569; font-size: 0.85rem; margin-top: 6px; font-weight: 600; text-transform: uppercase; }}
                </style>
            </head>
            <body>
                <div class="header">📄 Scanned Document ({total_pages} Pages)</div>
                {image_tags_html}
            </body>
            </html>
            """

            doc_html_path = os.path.join(SERVER_DIR, "index.html")
            with open(doc_html_path, "w", encoding="utf-8") as f:
                f.write(html_content)

            # 3. Create QR code pointing to index.html
            local_ip = get_local_ip()
            document_url = f"http://{local_ip}:{PORT}/index.html"

            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=6,
                border=2,
            )
            qr.add_data(document_url)
            qr.make(fit=True)

            self.current_qr_img = qr.make_image(fill_color="black", back_color="white")

            qr_display = self.current_qr_img.resize((170, 170))
            self.qr_tk = ImageTk.PhotoImage(qr_display)
            self.qr_canvas.delete("all")
            self.qr_canvas.create_image(90, 90, image=self.qr_tk, anchor=tk.CENTER)

            messagebox.showinfo("Full Document QR Ready", f"Full Document QR Generated!\n\nPhone se scan karne par poori file ke saare {total_pages} pages ek sath mobile screen par scroll karke dikhenge.\n\nLink: {document_url}")

        except Exception as e:
            messagebox.showerror("Error", f"QR Code generate nahi ho paya: {str(e)}")

    def save_qr_code(self):
        if self.current_qr_img is None:
            messagebox.showwarning("Warning", "Pehle 'Generate Full Document QR Code' click karke QR generate karein!")
            return

        save_qr_path = filedialog.asksaveasfilename(
            title="Save Generated QR Code",
            defaultextension=".png",
            filetypes=[("PNG Image", "*.png")]
        )
        if not save_qr_path:
            return

        try:
            self.current_qr_img.save(save_qr_path)
            messagebox.showinfo("Success", f"QR Code successfully save ho gaya:\n{save_qr_path}")
        except Exception as e:
            messagebox.showerror("Error", f"QR Code save nahi ho paya: {str(e)}")

    def remove_selected_page(self):
        selected = self.pages_listbox.curselection()
        if not selected:
            return
        idx = selected[0]
        self.raw_warped_pages.pop(idx)
        self.scanned_pages.pop(idx)
        self.pages_listbox.delete(idx)
        self.canvas.delete("all")
        self.qr_canvas.delete("all")
        self.current_qr_img = None

    def export_pdf(self):
        if not self.scanned_pages:
            messagebox.showwarning("Warning", "Export karne ke liye koi pages add nahi hain!")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF File", "*.pdf")])
        if not save_path:
            return
        try:
            c = canvas.Canvas(save_path, pagesize=letter)
            pdf_w, pdf_h = letter

            for img in self.scanned_pages:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_filename = tmp.name
                    save_img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR) if len(img.shape) == 3 else img
                    cv2.imwrite(tmp_filename, save_img)

                img_h, img_w = img.shape[:2]
                scale = min(pdf_w / img_w, pdf_h / img_h)
                render_w, render_h = img_w * scale, img_h * scale
                x_offset = (pdf_w - render_w) / 2
                y_offset = (pdf_h - render_h) / 2

                c.drawImage(tmp_filename, x_offset, y_offset, width=render_w, height=render_h)
                c.showPage()

                if os.path.exists(tmp_filename):
                    os.remove(tmp_filename)

            c.save()
            messagebox.showinfo("Export Successful", f"PDF Successfully Exported:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"PDF Save nahi ho paya: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ProfessionalScannerApp(root)
    root.mainloop()