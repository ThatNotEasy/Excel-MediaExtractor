# Extract Excel Files

This Python script extracts tables, images, and embedded files (such as PDFs) from an Excel file and saves them in a specified output folder.

## Features
- Extracts tables from all sheets in the Excel file and saves them as CSV files.
- Extracts embedded images from sheets and saves them as PNG files.
- Extracts embedded PDF files from Excel and saves them separately.

## Requirements
Ensure you have the following Python libraries installed:

```sh
pip install pandas openpyxl pillow
```

## Usage

1. Place your Excel file in the project directory.
2. Update the `excel_file_path` variable with your file name.
3. Run the script:

```sh
python main.py
```

## Example Output Structure
```
output/
│── Sheet1.csv
│── Sheet2.csv
│── Sheet1_image_1.png
│── Sheet2_image_2.png
│── embedded_file_1.pdf
```

## License
This project is open-source and can be modified as needed.

## Author
Created by MrDodgerX
