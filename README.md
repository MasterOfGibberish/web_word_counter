# Web Word Counter

A tool for crawling websites and generating word counts for translation cost estimation or content analysis.

## Features

- Crawls websites and extracts text content from pages
- Ignores duplicate pages and normalizes URLs
- Filters out common content like headers and footers that appear on multiple pages
- Calculates accurate word counts for text content
- Exports to Excel or Word format with comprehensive details
- Handles errors gracefully with multiple fallback mechanisms
- **User-friendly GUI interface** for easy operation

## Installation

1. Clone this repository:
```
git clone https://github.com/MasterOfGibberish/web_word_counter.git
cd web_word_counter
```

2. Install required packages:
```
pip install -r requirements.txt
```

## Python Version Compatibility

This tool is compatible with:
- Python 3.5+ (all features)
- Python 3.4+ (may require modifying f-strings)

If you experience syntax errors related to f-strings when running on older Python versions, the program has been updated to use the older string formatting style.

## Usage

### GUI Interface (Recommended)

To use the graphical user interface, simply run:
```
python run.py
```

This will open a user-friendly window where you can:
- Enter a website URL
- Set the maximum pages to crawl
- Choose the output format (Excel or Word)
- Configure additional options
- Monitor the crawling progress in real-time

![GUI Screenshot](screenshot.png)

### Command Line Interface

Basic usage:
```
python main.py https://example.com
```

Advanced options:
```
python main.py https://example.com -l 20 -o report.xlsx --format excel --overwrite
```

You can also use the launcher with arguments to access the command line interface:
```
python run.py https://example.com -l 20
```

### Command line options

- `url` - The website URL to crawl (required)
- `-l, --limit` - Maximum number of pages to crawl (default: 10)
- `-o, --output` - Output filename (default: website_text.xlsx)
- `-w, --wait` - Wait time in seconds for page loading (default: 5)
- `-t, --threads` - Number of threads to use for parallel processing (default: 4)
- `--format` - Output format: "excel" or "docx" (default: excel) 
- `--overwrite` - Overwrite existing output file if it exists

## Output

The tool generates either:
- An Excel file with two sheets:
  - Summary sheet with total page count and word count
  - Content sheet with all pages, their text content, and individual word counts
- A Word document with:
  - Summary information
  - Table of pages with their word counts
  - Full text content of each page

## Requirements

- Python 3.5 or higher recommended (3.4+ with modifications)
- Chrome browser (for web rendering)
- Python packages listed in requirements.txt
- tkinter and ttkthemes (for GUI interface)

## Troubleshooting

If you encounter errors like `SyntaxError: f-string expression part cannot include a backslash`, you are likely using an older version of Python (before 3.6). The program has been updated to use older string formatting, but if you still encounter issues, please:

1. Update to Python 3.6 or newer if possible
2. Or run with the `-t 1` option to disable multi-threading which may help with some compatibility issues

## License

MIT 