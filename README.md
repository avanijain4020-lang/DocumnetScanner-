# 📄 Advanced Multi-Format Batch Document Scanner & QR Viewer

An all-in-one desktop application built with Python, OpenCV, and Tkinter to scan, process, and organize documents. Import images, PDFs, or PowerPoint presentations, auto-detect document edges, apply enhancement filters, and instantly share multi-page documents via a local web server and QR code reader.

---

## 🚀 Key Features

* **Multi-Format Import:** Batch import images (`.jpg`, `.png`, `.bmp`), multi-page PDFs (`.pdf`), and PowerPoint presentations (`.pptx`, `.ppt`).
* **Smart Edge & Corner Detection:** Automatic document boundary detection using OpenCV's contour analysis and 4-point perspective transformation.
* **Image Enhancement Filters:**
  * **B&W Document Scan:** Clean high-contrast black & white scan.
  * **Magic Color:** Enhanced color and sharpness.
  * **Smooth Grayscale:** Clean gray tone document.
  * **Original Image:** Unfiltered perspective view.
* **Bulk Operations:** Apply selected filters individually or to all pages in one click.
* **Interactive Multi-Page QR Code Viewer:**
  * Built-in local HTTP web server (`http.server` & `threading`).
  * Automatically generates an HTML reader page for all scanned pages.
  * Generates a QR code to view and scroll through all scanned pages instantly on a mobile phone (within the same Wi-Fi network).
* **Export & Management:**
  * Save individual QR codes as PNG images.
  * Reorder, remove, or manage individual pages.
  * Export all processed pages into a multi-page PDF document using ReportLab.
* **💻 Portable Executable (.exe):** Standalone Windows Desktop App available. No Python environment setup required!

---

## 🛠️ Tech Stack & Dependencies

* **Language & Runtime:** Python 3.12
* **GUI Framework:** Tkinter / ttk
* **Computer Vision & Processing:** OpenCV (`opencv-python`), NumPy, Pillow (`PIL`)
* **PDF & Presentation Libraries:** PyMuPDF (`pymupdf`), `python-pptx`, `reportlab`
* **Networking & Sharing:** `qrcode`, Standard Python HTTP Server (`http.server`), Socket, Threading
* **Packaging Tool:** PyInstaller

---

## 📥 Download Standalone Executable (Windows)

No Python installation needed! Just download and run:

1. Go to the [DocumentScanner Releases Page](https://github.com/avanijain4020-lang/DocumnetScanner-/releases/tag/v1.0.0).
2. Download `DocumentScanner.exe`.
3. Double-click `DocumentScanner.exe` to launch the application directly on Windows 10/11.

---

## 💻 Running from Source Code

### 1. Prerequisites
Ensure you have **Python 3.12** installed on your system.

### 2. Installation
Clone the repository and install the required Python packages:

```cmd
git clone [https://github.com/avanijain4020-lang/DocumnetScanner-.git](https://github.com/avanijain4020-lang/DocumnetScanner-.git)
cd DocumnetScanner-
py -3.12 -m pip install opencv-python numpy Pillow reportlab qrcode pymupdf python-pptx