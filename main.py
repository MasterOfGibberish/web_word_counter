import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from docx import Document
from tqdm import tqdm
import re
import argparse
import time
import threading
import queue
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from collections import Counter
import difflib
import pandas as pd
import os
import datetime
import sys
import shutil

visited_urls = set()
common_content = Counter()
MAX_THREADS = 4  # Number of parallel threads to use

# Get the directory where the script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) or os.getcwd()

def normalize_url(url):
    """Normalize URL to prevent duplicates with different formats."""
    parsed = urlparse(url)
    # Remove trailing slashes, convert to lowercase
    normalized = f"{parsed.scheme}://{parsed.netloc.lower()}{parsed.path.rstrip('/')}"
    # Remove common index files
    if normalized.endswith(('/index.html', '/index.php', '/index.asp')):
        normalized = normalized[:-10]  # Remove the index filename
    # Remove query strings and fragments if needed
    normalized = normalized.split('?')[0].split('#')[0]
    return normalized

def setup_driver():
    """Set up and return a headless Chrome driver with optimized settings."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1280,720")  # Smaller window size
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-images")  # Don't load images
    chrome_options.add_argument("--blink-settings=imagesEnabled=false")
    chrome_options.add_argument("--disable-javascript")  # Optional: can break some sites
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(15)  # Maximum page load time
    return driver

def get_visible_text_selenium(driver, url, wait_time=5):
    """Get visible text from a page using Selenium with wait for page load."""
    try:
        driver.get(url)
        
        # Wait for page to load (body to be present)
        WebDriverWait(driver, wait_time).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
        
        # Additional wait for dynamic content - reduced
        time.sleep(1)  # Give extra time for JavaScript to render
        
        # Get the page source after JavaScript execution
        html = driver.page_source
        
        # Parse with BeautifulSoup
        soup = BeautifulSoup(html, "html.parser")
        
        # Extract and remove headers, navigation, and footers
        header_tags = soup.find_all(['header', 'nav', 'footer'])
        nav_items = soup.find_all(class_=re.compile(r'(header|nav|menu|navigation|footer)'))
        
        # Process header, footer, nav content for common content detection
        for element in header_tags + nav_items:
            text = ' '.join(element.stripped_strings)
            if text:
                common_content[text] += 1
                element.decompose()
        
        # Remove scripts, styles, nav, etc.
        for tag in soup(['script', 'style', 'noscript']):
            tag.decompose()
        
        # Get visible text only
        texts = soup.stripped_strings
        visible_text = '\n'.join(texts)
        
        # Get all links for crawling
        links = []
        for a in soup.find_all('a', href=True):
            href = a['href']
            full_url = urljoin(url, href)
            links.append(full_url)
            
        return visible_text, links
        
    except TimeoutException:
        print(f"Timeout waiting for page to load: {url}")
        return "", []
    except WebDriverException as e:
        print(f"WebDriver error for {url}: {str(e).split('\n')[0]}")
        return "", []
    except Exception as e:
        print(f"Error fetching {url}: {str(e).split('\n')[0]}")
        return "", []

def process_page(driver, url, wait_time, results, base_domain, progress_callback=None):
    """Process a single page with a specific driver."""
    if url in visited_urls:
        return
        
    url_domain = urlparse(url).netloc
    if url_domain != base_domain:
        return
        
    visited_urls.add(url)
    
    # Use the original URL format for fetching
    fetch_url = url
    visible_text, links = get_visible_text_selenium(driver, fetch_url, wait_time)
    
    if not visible_text:
        return [], []  # Empty text and links if there's an error
    
    if progress_callback:
        progress_callback()
    
    results.append((url, visible_text, 0))  # Temporary 0 for word count
    
    # Return new links to visit
    new_links = []
    for link in links:
        normalized_link = normalize_url(link)
        if normalized_link not in visited_urls:
            url_domain = urlparse(normalized_link).netloc
            if url_domain == base_domain:
                new_links.append(normalized_link)
    
    return new_links

def detect_common_text_chunks(text_list, min_pages=2, chunk_size=50):
    """Detect text chunks that appear on multiple pages."""
    common_chunks = []
    # Process texts to create chunks
    chunked_texts = []
    
    for text in text_list:
        words = re.findall(r'\w+', text.lower())
        # Use fewer chunks to speed up processing
        chunks = [' '.join(words[i:i+chunk_size]) for i in range(0, len(words), chunk_size*2)]
        chunked_texts.append(chunks)
    
    # Find chunks that appear in multiple pages
    all_chunks = []
    for chunks in chunked_texts:
        all_chunks.extend(chunks)
    
    # Count chunk occurrences
    chunk_counter = Counter(all_chunks)
    
    # Get chunks that appear in at least min_pages pages
    for chunk, count in chunk_counter.items():
        if count >= min_pages and len(chunk.split()) >= 5:  # At least 5 words
            common_chunks.append(chunk)
    
    return common_chunks

def remove_common_chunks(text, common_chunks):
    """Remove common text chunks from content."""
    if not common_chunks:
        return text
        
    clean_text = text
    for chunk in common_chunks:
        clean_text = clean_text.replace(chunk, '')
    
    # Clean up any resulting extra whitespace
    clean_text = re.sub(r'\s+', ' ', clean_text).strip()
    return clean_text

def count_words(text):
    # Remove special characters and split by whitespace
    words = re.findall(r'\w+', text.lower())
    return len(words)

def crawl_and_extract(base_url, limit=10, progress_bar=True, wait_time=5):
    """Crawl website and extract text with optimized multi-threading approach."""
    # Normalize the base URL
    normalized_base_url = normalize_url(base_url)
    base_domain = urlparse(normalized_base_url).netloc
    
    to_visit = queue.Queue()
    to_visit.put(normalized_base_url)
    
    extracted_texts = []
    raw_texts = []  # Store raw texts for common content detection
    total_word_count = 0
    visited_urls.clear()  # Clear global visited URLs
    common_content.clear()  # Clear global common content counter
    
    # Set up multiple drivers for parallel processing
    num_threads = min(MAX_THREADS, limit)
    drivers = []
    for _ in range(num_threads):
        try:
            drivers.append(setup_driver())
        except Exception as e:
            print(f"Error creating WebDriver: {e}")
            # If we can't create all drivers, just use what we have
            break
    
    if not drivers:
        print("Failed to create any WebDrivers. Exiting.")
        return [], 0
    
    try:
        pbar = tqdm(total=limit, desc="Pages crawled") if progress_bar else None
        page_count = 0
        
        # Function to update progress bar
        def update_progress():
            nonlocal page_count
            page_count += 1
            if pbar:
                pbar.update(1)
                
        # Process pages with a timeout
        max_time = 180  # Maximum 3 minutes total processing time
        start_time = time.time()
        
        # Process first few pages sequentially to get initial links
        driver = drivers[0]
        while not to_visit.empty() and page_count < min(5, limit) and time.time() - start_time < max_time:
            url = to_visit.get()
            if url in visited_urls:
                continue
                
            new_links = process_page(driver, url, wait_time, extracted_texts, 
                                     base_domain, update_progress)
            
            if new_links:
                for link in new_links:
                    if link not in visited_urls:
                        to_visit.put(link)
        
        # Create a shared queue for remaining work
        work_queue = queue.Queue()
        while not to_visit.empty() and page_count < limit:
            url = to_visit.get()
            if url not in visited_urls:
                work_queue.put(url)
        
        # If there's not much work, don't use threading
        if work_queue.qsize() < 3:
            while not work_queue.empty() and page_count < limit and time.time() - start_time < max_time:
                url = work_queue.get()
                if url in visited_urls:
                    continue
                    
                new_links = process_page(drivers[0], url, wait_time, extracted_texts, 
                                        base_domain, update_progress)
                
                if new_links:
                    for link in new_links:
                        if link not in visited_urls and page_count < limit:
                            work_queue.put(link)
        else:
            # Process remaining pages in parallel
            def worker(driver_idx):
                driver = drivers[driver_idx]
                while not work_queue.empty() and page_count < limit and time.time() - start_time < max_time:
                    try:
                        url = work_queue.get(block=False)
                    except queue.Empty:
                        break
                        
                    if url in visited_urls:
                        continue
                        
                    new_links = process_page(driver, url, wait_time, extracted_texts, 
                                           base_domain, update_progress)
                    
                    if new_links:
                        for link in new_links:
                            if link not in visited_urls and page_count < limit:
                                try:
                                    work_queue.put(link)
                                except:
                                    pass  # Queue might be full
            
            # Start worker threads
            threads = []
            for i in range(len(drivers)):
                t = threading.Thread(target=worker, args=(i,))
                t.daemon = True
                threads.append(t)
                t.start()
                
            # Wait for all threads to complete
            for t in threads:
                t.join(timeout=max(1, max_time - (time.time() - start_time)))
        
        if pbar:
            pbar.close()
            
        # Get the text content from results
        raw_texts = [text for _, text, _ in extracted_texts]
            
        # Post-processing: detect and remove common text chunks
        if len(raw_texts) > 1:
            print("Detecting common content across pages...")
            common_chunks = detect_common_text_chunks(raw_texts, min_pages=2)
            
            # Clean and update word counts
            cleaned_texts = []
            for i, (url, text, _) in enumerate(extracted_texts):
                cleaned_text = remove_common_chunks(text, common_chunks)
                word_count = count_words(cleaned_text)
                total_word_count += word_count
                cleaned_texts.append((url, cleaned_text, word_count))
            
            extracted_texts = cleaned_texts
    finally:
        # Always close the drivers
        for driver in drivers:
            try:
                driver.quit()
            except:
                pass
    
    return extracted_texts, total_word_count

def get_default_output_path(filename):
    """Get path in the same directory as the script."""
    # If filename is already an absolute path, return it as is
    if os.path.isabs(filename):
        return filename
    # Otherwise, join with script directory
    return os.path.join(SCRIPT_DIR, filename)

def ensure_directory_exists(filepath):
    """Make sure the directory for the file exists, create if it doesn't."""
    directory = os.path.dirname(filepath)
    
    # If directory is empty, use current directory
    if not directory:
        return True
        
    try:
        # Create directory and any parent directories
        os.makedirs(directory, exist_ok=True)
        return True
    except Exception as e:
        print(f"Warning: Could not create directory '{directory}': {e}")
        # Return False to indicate failure
        return False

def get_unique_filename(filename):
    """Generate a unique filename if the original already exists."""
    if not os.path.exists(filename):
        return filename
        
    base, ext = os.path.splitext(filename)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{base}_{timestamp}{ext}"

def check_file_exists(filename, overwrite=False):
    """Check if file exists and handle accordingly."""
    if os.path.exists(filename):
        if overwrite:
            try:
                # Try to remove the file first to ensure we can write to it
                os.remove(filename)
                print(f"Removed existing file '{filename}' for overwriting.")
            except Exception as e:
                print(f"Warning: Could not remove existing file '{filename}': {e}")
                # If failed to remove, create a unique name
                new_name = get_unique_filename(filename)
                print(f"Using alternative filename: {new_name}")
                return new_name
            return filename
        else:
            new_filename = get_unique_filename(filename)
            print(f"File '{filename}' already exists. Saving as '{new_filename}' instead.")
            return new_filename
    return filename

def save_to_excel(text_blocks, filename="website_text.xlsx", overwrite=False):
    """Save the extracted text and word counts to an Excel file."""
    # Make sure the filename is in the script directory if not absolute
    filename = get_default_output_path(filename)
    
    # Ensure directory exists
    if not ensure_directory_exists(filename):
        # If can't create directory, save to script directory instead
        filename = os.path.join(SCRIPT_DIR, os.path.basename(filename))
        print(f"Saving to script directory instead: {filename}")
        ensure_directory_exists(filename)  # Create script directory if needed
    
    # Check if file exists and get appropriate filename
    filename = check_file_exists(filename, overwrite)
    
    # Create DataFrame with the required columns
    data = []
    for url, text, word_count in text_blocks:
        data.append({
            "Page URL": url,
            "Text Content": text,
            "Word Count": word_count
        })
    
    df = pd.DataFrame(data)
    
    # Add summary row
    total_words = sum(block[2] for block in text_blocks)
    
    # Save to Excel
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Website Content', index=False)
            
            # Create summary sheet
            summary_data = {
                "Description": ["Total Pages", "Total Word Count"],
                "Value": [len(text_blocks), total_words]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            # Auto-adjust column widths
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns):
                    if sheet_name == 'Website Content':
                        # Set reasonable max width for text content
                        if col == 'Text Content':
                            worksheet.column_dimensions[chr(66)].width = 100  # Column B
                        elif col == 'Page URL':
                            worksheet.column_dimensions[chr(65)].width = 50   # Column A
                        else:
                            worksheet.column_dimensions[chr(67)].width = 15   # Column C
        
        print(f"Excel file saved successfully as '{filename}'")
        return total_words
    except Exception as e:
        print(f"Error saving to Excel: {e}")
        
        # Fallback to CSV if Excel fails
        try:
            csv_filename = os.path.splitext(filename)[0] + '.csv'
            csv_filename = check_file_exists(csv_filename, overwrite)
            pd.DataFrame(data).to_csv(csv_filename, index=False)
            print(f"Saved as CSV instead: {csv_filename}")
        except Exception as csv_error:
            print(f"Error saving to CSV: {csv_error}")
            # Last resort: save to a very basic text file
            txt_filename = os.path.join(SCRIPT_DIR, f"website_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            try:
                with open(txt_filename, 'w', encoding='utf-8') as f:
                    f.write(f"Total Pages: {len(text_blocks)}\n")
                    f.write(f"Total Word Count: {total_words}\n\n")
                    for url, _, word_count in text_blocks:
                        f.write(f"URL: {url}\nWord Count: {word_count}\n\n")
                print(f"Saved as text file instead: {txt_filename}")
            except Exception as txt_error:
                print(f"Failed to save data in any format. Error: {txt_error}")
        
        return total_words

def save_to_docx(text_blocks, filename="website_text.docx", overwrite=False):
    """Save the extracted text and word counts to a Word document."""
    # Make sure the filename is in the script directory if not absolute
    filename = get_default_output_path(filename)
    
    # Ensure directory exists
    if not ensure_directory_exists(filename):
        # If can't create directory, save to script directory instead
        filename = os.path.join(SCRIPT_DIR, os.path.basename(filename))
        print(f"Saving to script directory instead: {filename}")
        ensure_directory_exists(filename)  # Create script directory if needed
    
    # Check if file exists and get appropriate filename
    filename = check_file_exists(filename, overwrite)
    
    doc = Document()
    
    # Add summary at the beginning
    doc.add_heading("Website Text Extraction Summary", level=1)
    doc.add_paragraph(f"Total pages extracted: {len(text_blocks)}")
    
    total_words = sum(word_count for _, _, word_count in text_blocks)
    doc.add_paragraph(f"Total word count: {total_words}")
    
    # Add table of contents with word counts
    doc.add_heading("Pages and Word Counts", level=2)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    
    # Add header row
    header_cells = table.rows[0].cells
    header_cells[0].text = "Page #"
    header_cells[1].text = "URL"
    header_cells[2].text = "Word Count"
    
    # Add content rows
    for i, (url, _, word_count) in enumerate(text_blocks, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(i)
        row_cells[1].text = url
        row_cells[2].text = str(word_count)
    
    doc.add_page_break()
    
    # Add page content
    for i, (url, text, word_count) in enumerate(text_blocks, 1):
        doc.add_heading(f"Page {i}: {url}", level=2)
        doc.add_paragraph(f"Word count: {word_count}")
        doc.add_paragraph(text)
        doc.add_page_break()
    
    try:
        doc.save(filename)
        print(f"Word document saved successfully as '{filename}'")
    except Exception as e:
        print(f"Error saving Word document: {e}")
        
        # Try saving to the script directory with a different name
        backup_filename = os.path.join(SCRIPT_DIR, f"backup_{os.path.basename(filename)}")
        try:
            doc.save(backup_filename)
            print(f"Saved Word document to alternative location: {backup_filename}")
        except Exception as e2:
            print(f"Failed to save Word document as backup. Error: {e2}")
            
            # Last resort: save to a basic text file
            txt_filename = os.path.join(SCRIPT_DIR, f"website_data_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            try:
                with open(txt_filename, 'w', encoding='utf-8') as f:
                    f.write(f"Total Pages: {len(text_blocks)}\n")
                    f.write(f"Total Word Count: {total_words}\n\n")
                    for url, text_content, word_count in text_blocks:
                        f.write(f"URL: {url}\nWord Count: {word_count}\n\n")
                        f.write(f"Content Preview (first 500 chars): {text_content[:500]}...\n\n")
                print(f"Saved as text file instead: {txt_filename}")
            except Exception as txt_error:
                print(f"Failed to save data in any format. Error: {txt_error}")
    
    return total_words

def main():
    parser = argparse.ArgumentParser(description="Web crawler that extracts text and provides word count for translation quotes")
    parser.add_argument("url", help="The base URL to crawl (e.g., https://example.com)")
    parser.add_argument("-l", "--limit", type=int, default=10, help="Maximum number of pages to crawl (default: 10)")
    parser.add_argument("-o", "--output", default="website_text.xlsx", help="Output filename (default: website_text.xlsx)")
    parser.add_argument("-w", "--wait", type=int, default=5, help="Wait time in seconds for page loading (default: 5)")
    parser.add_argument("-t", "--threads", type=int, default=4, help="Number of threads to use (default: 4)")
    parser.add_argument("--format", choices=["excel", "docx"], default="excel", help="Output format (default: excel)")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing output file if it exists")
    
    args = parser.parse_args()
    
    # Update the global thread count
    global MAX_THREADS
    MAX_THREADS = max(1, min(8, args.threads))
    
    # If output path is not absolute, make it relative to the script directory
    if not os.path.isabs(args.output):
        args.output = get_default_output_path(args.output)
    
    print(f"Starting crawl of {args.url}...")
    print(f"Using {MAX_THREADS} threads for faster processing.")
    
    start_time = time.time()
    content, total_words = crawl_and_extract(args.url, limit=args.limit, wait_time=args.wait)
    end_time = time.time()
    
    print(f"Extracted text from {len(content)} pages in {end_time - start_time:.1f} seconds.")
    print(f"Total word count: {total_words}")
    
    # Choose the appropriate output format
    if args.format == "docx" or args.output.endswith(".docx"):
        output_file = args.output
        if not output_file.endswith(".docx"):
            output_file = os.path.splitext(output_file)[0] + ".docx"
        save_to_docx(content, output_file, args.overwrite)
    else:
        output_file = args.output
        if not output_file.endswith((".xlsx", ".xls")):
            output_file = os.path.splitext(output_file)[0] + ".xlsx"
        save_to_excel(content, output_file, args.overwrite)
    
    # Final verification message
    if os.path.exists(output_file):
        file_size = os.path.getsize(output_file) / 1024  # Size in KB
        print(f"SUCCESS: Output file '{output_file}' (size: {file_size:.1f} KB) has been created successfully.")
        print(f"Location: {os.path.abspath(output_file)}")
    else:
        print(f"WARNING: Could not verify that output file '{output_file}' was created.")
        print("Check the script directory for alternative output files that may have been created.")

# --- USAGE ---
if __name__ == "__main__":
    main()
