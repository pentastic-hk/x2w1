# Convert Excel to Word Tables

Convert Follow-up Plan to Detailed Findings Tables in Businesses Reporting Template.

Specifications provided by David, vibe-coded with Claude Sonnet.

## Installation

### Installing Python

You should install Python version 3 or above on your machine.

### Cloning the Repo

Then, clone this repo to your machine.
```sh
git clone https://github.com/pentastic-hk/x2w1.git
```

### Setting Up the Virtual Environment

Python virtual environment is used to keep this project isolated.
Run the following command to setup your virtual environment:
```sh
python3 -m venv .venv
```
This creates a folder `.venv` in your project folder.

Then activate your virtual environment with the following command:
```sh
source .venv/bin/active
```

Or the following command if you're using Windows:
```sh
.venv\scripts\Activate
```

### Installing Python Dependencies

Run the following command to install the project dependencies to your virtual environment:
```sh
pip install -r requirements.txt
```

## Usage

```sh
python3 excel_to_word_findings.py path/to/input.xlsx
python3 excel_to_word_findings.py path/to/input.xlsx output.docx
python3 excel_to_word_findings.py path/to/input.xlsx --sheet "Follow-up Items"
python3 excel_to_word_findings.py path/to/input.xlsx --debug
```

If you're too lazy to type the full path to your Excel file,
drag your file from your Files Explorer onto the command line.
