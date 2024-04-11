Certainly! Here's a README file for your code:

---

# XML to XLSX Converter and Data Merger

This Python script is designed to convert XML files to XLSX format and merge data from multiple XLSX files into a single file.

## Features

- **XML to XLSX Conversion**: Converts XML files to XLSX format while preserving XML attributes.
- **Data Merging**: Merges data from multiple XLSX files based on specified columns.
- **Automatic Chunking**: Splits merged data into smaller chunks if the size exceeds a specified limit.

## Requirements

- Python 3.x
- `xml.etree.ElementTree` module
- `openpyxl` library
- `pandas` library

## Usage

### XML to XLSX Conversion

```python
from xml_to_xlsx import xml_folder_to_xlsx

# Specify the folder containing XML files and the folder to save XLSX files
xml_folder = "path/to/xml_folder"
xlsx_folder = "path/to/xlsx_folder"

# Convert XML files to XLSX
xlsx_files = xml_folder_to_xlsx(xml_folder, xlsx_folder)
```

### Data Merging

```python
from data_merger import merge_xlsx_files

# Specify the folder containing XLSX files
folder_path = "path/to/xlsx_folder"

# Merge data from XLSX files
merge_xlsx_files(folder_path)
```

## Folder Structure

- **xml_to_xlsx.py**: Contains functions for XML to XLSX conversion.
- **data_merger.py**: Contains functions for merging data from XLSX files.
- **test**: Example folder containing XML and XLSX files.

## How it Works

1. **XML to XLSX Conversion**: The script parses XML files and extracts attribute data. It then organizes the data into XLSX files, preserving the structure of the XML attributes.

2. **Data Merging**: The script reads multiple XLSX files, extracts specified columns, and merges them into a single DataFrame. If the merged data exceeds a specified size limit, it automatically splits the data into smaller chunks, each saved as a separate XLSX file.

## Limitations

- XML parsing errors may occur if the XML files are not well-formed.
- The script assumes consistent column structures across XLSX files for merging.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Feel free to customize this README according to your project's specific details and requirements!
