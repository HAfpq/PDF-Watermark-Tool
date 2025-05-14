#PDF Watermarker GUI

**PDF Watermarker GUI** is a desktop application developed based on PyQt5 and ReportLab/PyPDF2, which aims to help users add text and image watermarks to PDF files in batches. It supports customizing watermark fonts, colors, sizes, transparency, positions and other parameters, and users can flexibly customize the style and layout of each watermark.

## ✨ Features

- 🗂️ **Batch processing**: Support selecting folders and adding watermarks to PDF files in batches.
- 📝 **Text Watermark**: You can add custom text watermarks and support setting text size, font, color and transparency.
- 🖼️ **Image Watermark**: Support uploading custom Logo images as watermarks.
- 🎨 **Customized watermark style**: including watermark transparency, text size, color, etc.
- 🔢 **Watermark quantity control**: Support setting the horizontal and vertical arrangement quantity of watermarks per page.
- 📐 **Watermark position and angle**: Support setting the watermark display position (center, four corners) and rotation angle.
- 🖥️ **Intuitive graphical interface**: User-friendly, easy to operate, suitable for users with no programming experience.
- 🚀 **Efficient processing**: Use multi-threading to process PDF files and support fast processing of large batches of files.

## 🛠️ Technology Stack

- **Python 3.x**
- **PyQt5** — for building graphical user interfaces
- **ReportLab** — for drawing text and graphics to PDF
- **PyPDF2** — for processing PDF file merging and editing
- **Pillow (PIL)** — for processing and displaying logo images
- **Multithreading** — Improve batch processing efficiency

## 📦 Install dependencies

Create a virtual environment in the project directory and install dependencies:

```bash
pip install -r requirements.txt        
