# 📄 Advanced Multi-Format Batch Document Scanner & QR Viewer

An all-in-one desktop application built with **Python**, **OpenCV**, and **Tkinter** to scan, process, and organize documents. Import images, PDFs, or PowerPoint presentations, auto-detect document edges, apply enhancement filters, and instantly share multi-page documents via a local web server and QR code reader.

---

## ✨ Key Features

- **Multi-Format Import**: Batch import images (`.jpg`, `.png`, `.bmp`), multi-page PDFs (`.pdf`), and PowerPoint presentations (`.pptx`, `.ppt`).
- **Smart Edge & Corner Detection**: Automatic document boundary detection using OpenCV's contour analysis and 4-point perspective transformation.
- **Image Enhancement Filters**:
  - `B&W Document Scan` (Clean high-contrast black & white scan)
  - `Magic Color` (Enhanced color and sharpness)
  - `Smooth Grayscale` (Clean gray tone document)
  - `Original Image` (Unfiltered perspective view)
- **Bulk Operations**: Apply selected filters individually or to all pages in one click.
- **Interactive Multi-Page QR Code Viewer**:
  - Built-in local HTTP web server (`http.server` & `threading`).
  - Automatically generates an HTML reader page for all scanned pages.
  - Generates a QR code to view and scroll through all scanned pages instantly on a mobile phone (within the same Wi-Fi network).
- **Export & Management**:
  - Save individual QR codes as PNG images.
  - Reorder, remove, or manage individual pages.
  - Export all processed pages into a multi-page **PDF** document using ReportLab.

---

## 🛠️ Tech Stack & Dependencies

- **Programming Language**: Python 3.x
- **GUI Framework**: Tkinter / ttk
- **Computer Vision & Processing**: OpenCV (`opencv-python`), NumPy, Pillow (PIL)
- **PDF & Presentation Libraries**: PyMuPDF (`fitz`), `python-pptx`, ReportLab
- **Networking & Web**: `qrcode`, Standard Python HTTP Server (`http.server`), Socket, Threading

---

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have Python installed on your system.

### 2. Installation
Clone the repository and install the required Python packages:

```bash
git clone [https://github.com/avanijain4020-lang/DocumnetScanner-.git](https://github.com/avanijain4020-lang/DocumnetScanner-.git)
cd DocumnetScanner
pip install opencv-python numpy Pillow reportlab qrcode PyMuPDF python-pptx