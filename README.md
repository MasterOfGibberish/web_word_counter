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
- `-w, --wait` - Wait time in seconds for page loading (default: 10)
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

- Python 3.6 or higher
- Chrome browser (for web rendering)
- Python packages listed in requirements.txt
- tkinter and ttkthemes (for GUI interface)

## License

MIT 