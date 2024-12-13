# Excel Tool

A simple tool to upload an Excel file, replace `[variant]` placeholders with values and download the updated file made for Ovation.

## Features

- Upload an Excel file with placeholders.
- Replace `[variant]` placeholders with values from another sheet.
- Add new rows for each variant.
- Remove rows with `[variant]` placeholders.
- Download the updated file with `Final` added to the name.

## How to Use

1. Upload an Excel file with two sheets:
   - **Product**: Sheet with `[variant]` placeholders.
   - **Variant**: Sheet with variant values.
   
2. Click the **"Download bestand"** button to process the file.

3. The new file will be downloaded with `Final` added to the original name.

### Example:  
Upload `ProductList.xlsx` and the new file will be `ProductList Final.xlsx`.

## Technologies

- **HTML**: Structure.
- **CSS**: Styling.
- **JavaScript**: Uses `ExcelJS` for processing.

## Setup

To install `ExcelJS`:

```bash
npm install exceljs
