# Batch Excel (XLSX) and CSV to PDF Converter

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A powerful Python utility to batch convert Excel (.xlsx) and CSV (.csv) files into high-quality PDF documents, with robust options for styling, handling large files, and managing complex data structures.

## ✨ Features

- **🔄 Batch Conversion:** Process multiple `.xlsx` and `.csv` files automatically
- **📁 Recursive Search:** Optionally scan through subdirectories of the input folder
- **🏗️ Mirrored Output Structure:** Replicates input folder structure in output directory
- **📊 Excel Sheet Handling:** 
  - Combine all sheets into a single PDF document
  - Create separate PDF files for each sheet
- **🎨 Customizable PDF Styling:**
  - Control page size (A4, A3, letter) and orientation
  - Adjust margins, font size, and cell padding
  - Apply scaling factors for dense tables
- **⚡ Large File Handling:** Automatic table chunking for memory optimization
- **📝 Detailed Logging:** Comprehensive console output with timestamps
- **🛡️ Error Handling:** Graceful handling of common issues
- **🔧 CSV Encoding Detection:** Automatic detection of file encodings

## 📋 Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
  - [Directory Structure](#directory-structure)
  - [Configuration](#configuration)
  - [Running the Script](#running-the-script)
- [PDF Styling Options](#pdf-styling-options)
- [Handling Large Files](#handling-large-files)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## 🔧 Prerequisites

### Python Requirements
- **Python 3.7+** is required
- **pip** package installer

### System Dependencies for WeasyPrint

WeasyPrint requires several system libraries. Install them based on your operating system:

#### 🐧 Linux (Debian/Ubuntu)
```bash
sudo apt-get update
sudo apt-get install python3-dev python3-pip python3-setuptools python3-wheel \
    python3-cffi libcairo2 libpango-1.0-0 libpangocairo-1.0-0 \
    libgdk-pixbuf2.0-0 libffi-dev shared-mime-info
```

#### 🍎 macOS (using Homebrew)
```bash
brew install pango cairo libffi gdk-pixbuf
```

#### 🪟 Windows
1. Download the GTK+ runtime from [GTK's official Windows page](https://www.gtk.org/docs/installations/windows/)
2. Install the runtime environment
3. Add the GTK+ `bin` directory to your system's `PATH`

> **Note:** For detailed Windows installation instructions, refer to the [WeasyPrint documentation](https://doc.weasyprint.org/stable/install.html)

## 🚀 Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/SakibAhmedShuva/batch-excel-xlsx-csv-to-pdf-converter.git
   cd batch-excel-xlsx-csv-to-pdf-converter
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # macOS/Linux
   source venv/bin/activate
   ```

3. **Install system dependencies:** Follow the [Prerequisites](#prerequisites) section for your OS

4. **Install Python packages:**
   ```bash
   pip install -r requirements.txt
   ```

## 📖 Usage

### Directory Structure

```
your-project/
├── input_files/          # Your .xlsx and .csv files
│   ├── file1.xlsx
│   ├── file2.csv
│   └── subfolder/
│       └── file3.xlsx
├── output_pdf_files/     # Generated PDFs (created automatically)
└── app.py               # Main conversion script
```

### Configuration

Open `app.py` and modify the configuration variables at the top:

```python
# --- MAIN CONFIGURATION ---
INPUT_DIR = r"your/input/directory/path"  # ⚠️ CHANGE THIS
OUTPUT_PDF_DIR = "output_pdf_files"
RECURSIVE_SEARCH = True
CREATE_SUBFOLDERS_IN_OUTPUT = True
FILE_EXTENSIONS_TO_CONVERT = [".xlsx", ".csv"]

# --- EXCEL SHEET HANDLING ---
OPTION_COMBINE_SHEETS_INTO_ONE_PDF = False

# --- PDF STYLING ---
PAGE_SIZE = "A4 landscape"
PAGE_MARGIN = "0.4in"
BASE_FONT_SIZE_PT = 7
CELL_PADDING_PX = "2px 3px"
BODY_SCALE_FACTOR = 1.0

# --- LARGE FILE HANDLING ---
ENABLE_TABLE_CHUNKING = True
ROW_COUNT_THRESHOLD_FOR_CHUNKING = 1000
ROW_CHUNK_SIZE = 500
```

### Running the Script

```bash
python app.py
```

The script will display detailed progress information as it processes your files.

## 🎨 PDF Styling Options

| Setting | Description | Example Values |
|---------|-------------|----------------|
| `PAGE_SIZE` | Paper size and orientation | `"A4 landscape"`, `"A3 portrait"`, `"letter"` |
| `PAGE_MARGIN` | Margins around content | `"0.4in"`, `"1cm"`, `"10mm"` |
| `BASE_FONT_SIZE_PT` | Font size for table content | `6`, `7`, `8`, `10` |
| `CELL_PADDING_PX` | Padding inside table cells | `"2px 3px"`, `"1px 2px"` |
| `BODY_SCALE_FACTOR` | Overall content scaling | `1.0` (100%), `0.9` (90%) |

> **💡 Tip:** For tables with many columns, try reducing font size and using landscape orientation.

## ⚡ Handling Large Files

The script includes automatic table chunking for large datasets:

- **`ENABLE_TABLE_CHUNKING`**: Enable/disable chunking feature
- **`ROW_COUNT_THRESHOLD_FOR_CHUNKING`**: Minimum rows to trigger chunking
- **`ROW_CHUNK_SIZE`**: Rows per chunk

When chunking is active, large tables are split into smaller sections within the same PDF, with clear part indicators (e.g., "Sheet1 (Part 1 of 3)").

## 🔧 Troubleshooting

### Common Issues

#### `ModuleNotFoundError: No module named 'weasyprint'`
```bash
# Ensure virtual environment is activated, then:
pip install -r requirements.txt
```

#### WeasyPrint errors or hangs
- **Cause:** Missing system dependencies
- **Solution:** Carefully follow the [Prerequisites](#prerequisites) section
- **Windows users:** Ensure GTK+ is in your PATH

#### Script hangs on specific files
- Check console logs for the last processed file
- Try enabling table chunking with smaller chunk sizes
- Temporarily move problematic files to test others

#### "No files found to convert"
- Verify `INPUT_DIR` path is correct
- Check file extensions match `FILE_EXTENSIONS_TO_CONVERT`
- If `RECURSIVE_SEARCH = False`, ensure files are directly in `INPUT_DIR`

#### PDF content is cut off
- Reduce `BASE_FONT_SIZE_PT`
- Use smaller `CELL_PADDING_PX`
- Try `BODY_SCALE_FACTOR` < 1.0
- Consider larger `PAGE_SIZE` (e.g., "A3 landscape")

### Debug Tips

1. **Check file permissions:** Ensure files aren't locked by other applications
2. **Test with small files first:** Verify setup before processing large datasets
3. **Monitor console output:** Look for specific error messages and timestamps
4. **Try different configurations:** Experiment with styling options for better fit

## 🤝 Contributing

We welcome contributions! Here's how to get started:

1. **Fork** the repository
2. **Create** a feature branch: `git checkout -b feature/amazing-feature`
3. **Commit** your changes: `git commit -m 'Add amazing feature'`
4. **Push** to the branch: `git push origin feature/amazing-feature`
5. **Open** a Pull Request

### Development Guidelines

- Write clear, descriptive commit messages
- Add comments for complex logic
- Update documentation for new features
- Test with various file types and sizes

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Built with [pandas](https://pandas.pydata.org/) for data processing
- PDF generation powered by [WeasyPrint](https://weasyprint.org/)
- Excel support via [openpyxl](https://openpyxl.readthedocs.io/)

---

**⭐ Found this helpful?** Please star the repository and share with others!

**🐛 Found a bug?** Please open an issue with detailed information.

**💡 Have a feature request?** We'd love to hear your ideas!
