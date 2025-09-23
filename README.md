# File Data Extractor (Python 2.7.18)
Grab keywords to extract key information from files. Compatible with new versions of Python 3 and Python 2.7.18
This tool extracts data based on user-defined keywords (e.g., `Chapter`, `Element (VAL, ANGLE...)`) and saves the results as separate `.csv` or `.xlsx` files, each with its own filename for easy access.

## Prerequisites

- **Python 2.7.18** (Note: Python 2 is end-of-life; consider upgrading to Python 3)
- OpenPyXL (pre-configured in virtual environment)

## Setup Instructions

### 1. Create Virtual Environment
```bash
python -m venv myenv
```

### 2. Activate Virtual Environment
**Windows:**
```cmd
myenv\Scripts\activate
```

**macOS/Linux:**
```bash
source myenv/bin/activate
```

> OpenPyXL is already configured in this virtual environment (no additional installation required).

## Usage
1. Place your source files in the working directory
2. Run the extractor with your custom keywords
3. Find generated `.csv` or `.xlsx` files in output directory (each named by data category)

## Important Notes
- **Python 2.7.18 is end-of-life** (as of January 1, 2020)
- Always deactivate when finished:
  ```bash
  deactivate
  ```
- Recommended `.gitignore` entries:
  ```
  myenv/
  *.csv
  *.xlsx
  ```
