# Image Inserter @ Word files

This Python script inserts images into a Microsoft Word document. You can choose between **full-size** (1 image per page) or **half-size** (2 images per page). For half-size images, you’ll need to manually adjust their positions after insertion since they are added as floating shapes. You can also select your preferred **paper size**: Short, A4, or Long. The script automatically sets the correct layout, margins, and spacing based on your choices.

---

## Features

- Paper size options:
  - SHORT (8.5 x 11 in)
  - A4 (8.27 x 11.69 in)
  - LONG (8.5 x 13 in)
- Image layout options:
  - Full Size – one image per page
  - Half Size – two images per page
- Automatic formatting (margins, spacing, page breaks)
- Multiple pictures supported

---

## Requirements

- Python 3.12.4 or latest versions. [Download Python 3.12.4](https://www.python.org/downloads/release/python-3124/)
- Microsoft Word (Windows)
- Python Modules
  - `pywin32`
  - `tkinter`
  - `colorama`

 Or install the dependencies using pip @cmd
 
 ``
 pip install pywin32 colorama``
