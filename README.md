# OCR Measurement Extractor

This project provides a Python tool for automatically extracting measurement data from bitmap images and exporting the results to an Excel file.
The program was created to speed up receiving data for reproducibility test for PN-EN 54-23:2010.
The tool uses OCR to read numerical values from images, aggregates the extracted data, and generates a formatted Excel table.

---

## Features

The program:

- reads bitmap images from a specified directory
- preprocesses images to improve OCR accuracy
- extracts numerical values using OCR
- aggregates measurement data into a table
- exports the results to an Excel file
- highlights:
  - **maximum value (Qmax)** in red
  - **minimum value (Qmin)** in yellow
- applies table borders for readability
---

## Requirements

Python 3.9+

Required Python libraries:
 -opencv-python
 -pytesseract
 -pandas
 -openpyxl 
- tessaract 

## Installation
To install proper libraries type:
```
pip install opencv-python pytesseract pandas openpyxl
```
</br>
Tessaract needs to be installed from https://github.com/tesseract-ocr/tesseract. </br>
<u> </u>You can put a file path to it in the script, or you can add it to the environment variables and ommit that line </u>

## Project stucture
project </br>
│ </br>
├── ocr_analyse.py </br>
├── pliki/ </br>
│   ├── 1.bmp </br>
│   ├── 2.bmp </br>
│   └── ... </br>
│ </br>
└── results.xlsx
