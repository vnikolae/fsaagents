# Write a Python class `SECFilingAgent` with the following features:

# 1. **Dependencies & Setup**  
#    - Uses `requests`, `pandas`, and `xlsxwriter`.  
#    - Defines a `User-Agent` header for SEC EDGAR API calls.

# 2. **CIK Lookup**  
#    - Downloads the ticker→CIK mapping from `https://www.sec.gov/files/company_tickers.json`.  
#    - Provides a `get_cik(ticker)` method that returns the zero-padded 10-digit CIK or raises an error.

# 3. **Historical Chunks Loader** (`_load_historical_entries`)  
#    - Fetches the main submissions JSON at `https://data.sec.gov/submissions/CIK{CIK}.json`.  
#    - Reads the `"files"` array of chunk metadata.  
#    - For each chunk file (e.g. `CIK0000789019-submissions-001.json`), downloads and parses its parallel arrays (`accessionNumber`, `form`, `filingDate`, `reportDate`, `primaryDocument`) into a flat list of dicts.  
#    - Annotates each entry with a `source` key equal to the chunk filename.

# 4. **Recent Window Loader** (`_load_recent_entries`)  
#    - Reuses the same submissions JSON to pull the `"recent"` section.  
#    - Extracts parallel arrays (`form`, `accessionNumber`, `reportDate`, `primaryDocument`) into dicts, marking `source = 'recent'`.

# 5. **Unified Retrieval** (`get_filings(ticker, form_type, years_back)`)  
#    - Validates `form_type` against a `VALID_FORM_TYPES` list.  
#    - Computes a cutoff date (`today - years_back`).  
#    - Calls both loaders, concatenates their outputs, filters by `form_type` and cutoff date.  
#    - Builds an SEC-Archive URL using `accessionNumber` and `document`.  
#    - Deduplicates entries (favoring recent ones) and sorts by `filingDate` descending.  

# 6. **Export to Excel** (`save_to_excel(filings, filename)`)  
#    - Takes the list of dicts, converts it to a pandas DataFrame.  
#    - Converts the `link` column into `=HYPERLINK(...)` formulas.  
#    - Writes to an Excel file with auto-adjusted column widths.  

# Include comprehensive inline comments explaining each step. At the end, show example usage in a standard Python script (e.g. Spyder), fetching the last 10 years of 10-Ks for “MSFT” and saving them to `MSFT_10Ks_combined_10yrs.xlsx`.  

# Install dependencies if needed
# pip install pandas requests xlsxwriter

import requests
import pandas as pd
from typing import List, Dict
from datetime import datetime, timedelta

class SECFilingAgent:
    """
    Agent to fetch both historical and recent SEC filings by form type over a number of years,
    and export them to Excel with clickable hyperlinks and source annotations.
    """
    # Supported SEC forms
    VALID_FORM_TYPES = [
        '10-K', '10-Q', '8-K', '20-F', '6-K', 'S-1', 'S-3',
        '13F', '424B2', 'DEF 14A', 'SC 13G', 'SC 13D', 'N-PORT', 'N-CSR'
    ]

    # URLs and headers
    SEC_CIK_LOOKUP_URL = 'https://www.sec.gov/files/company_tickers.json'
    EDGAR_SUBMISSIONS_URL = 'https://data.sec.gov/submissions/CIK{cik}.json'
    CHUNKS_BASE_URL = 'https://data.sec.gov/submissions/'
    ARCHIVE_BASE = 'https://www.sec.gov/Archives/edgar/data/'
    HEADERS = {'User-Agent': 'Your Name your.email@example.com'}

    def __init__(self):
        # Load mapping of ticker -> zero-padded CIK
        self.cik_map = self._load_cik_map()

    def _load_cik_map(self) -> Dict[str, str]:
        """
        Download the ticker-to-CIK mapping from the SEC and build a dict.
        """
        resp = requests.get(self.SEC_CIK_LOOKUP_URL, headers=self.HEADERS)
        resp.raise_for_status()
        data = resp.json()
        return {v['ticker'].upper(): str(v['cik_str']).zfill(10) for v in data.values()}

    def get_cik(self, ticker: str) -> str:
        """
        Return the 10-digit CIK for a given stock ticker.
        Raises ValueError if ticker is unknown.
        """
        tk = ticker.upper()
        if tk not in self.cik_map:
            raise ValueError(f"Unknown ticker: {ticker}")
        return self.cik_map[tk]

    def _load_historical_entries(self, cik: str) -> List[Dict]:
        """
        Fetch and parse all EDGAR JSON chunk files for a given CIK.
        Each chunk contains parallel arrays of filing metadata (accessionNumber, form, dates, etc.).
        Returns a list of dicts with keys: form, accessionNumber, filingDate, reportDate, document, source_chunk.
        """
        url = self.EDGAR_SUBMISSIONS_URL.format(cik=cik)
        resp = requests.get(url, headers=self.HEADERS)
        resp.raise_for_status()
        summary = resp.json().get('filings', {}).get('files', [])

        entries = []
        for chunk in summary:
            chunk_name = chunk.get('name')
            if not chunk_name:
                continue
            chunk_url = f"{self.CHUNKS_BASE_URL}{chunk_name}"
            r = requests.get(chunk_url, headers=self.HEADERS)
            r.raise_for_status()
            data = r.json()

            # Parse parallel arrays
            if 'accessionNumber' in data and isinstance(data['accessionNumber'], list):
                forms = data.get('form', [])
                accession_nums = data['accessionNumber']
                filing_dates = data.get('filingDate', [])
                report_dates = data.get('reportDate', [])
                primary_docs = data.get('primaryDocument', [])

                for i, acc in enumerate(accession_nums):
                    entries.append({
                        'form': forms[i] if i < len(forms) else None,
                        'accessionNumber': acc,
                        'filingDate': filing_dates[i] if i < len(filing_dates) else None,
                        'reportDate': report_dates[i] if i < len(report_dates) else None,
                        'document': primary_docs[i] if i < len(primary_docs) else None,
                        'source': chunk_name  # annotate which chunk this came from
                    })
            # Fallback: nested 'filings' list
            elif 'filings' in data and isinstance(data['filings'], list):
                for entry in data['filings']:
                    entry['source'] = chunk_name
                    entries.append(entry)
        return entries

    def _load_recent_entries(self, cik: str) -> List[Dict]:
        """
        Fetch the 'recent' window of filings for a given CIK.
        Returns a list of dicts with keys: form, accessionNumber, reportDate, document, source='recent'.
        """
        url = self.EDGAR_SUBMISSIONS_URL.format(cik=cik)
        resp = requests.get(url, headers=self.HEADERS)
        resp.raise_for_status()
        recent = resp.json().get('filings', {}).get('recent', {})

        entries = []
        forms = recent.get('form', [])
        accessions = recent.get('accessionNumber', [])
        report_dates = recent.get('reportDate', [])
        primary_docs = recent.get('primaryDocument', [])

        for i, frm in enumerate(forms):
            entries.append({
                'form': frm,
                'accessionNumber': accessions[i],
                'filingDate': report_dates[i],  # use reportDate for cutoff
                'reportDate': report_dates[i],
                'document': primary_docs[i],
                'source': 'recent'
            })
        return entries

    def get_filings(
        self,
        ticker: str,
        form_type: str,
        years_back: int = 5
    ) -> List[Dict]:
        """
        Retrieve all filings of `form_type` within the last `years_back` years,
        combining both historical and recent sources and annotating each entry with its source.
        """
        if form_type not in self.VALID_FORM_TYPES:
            raise ValueError(f"Form {form_type} not supported.")

        # Compute cutoff date
        cutoff = datetime.utcnow().date() - timedelta(days=365 * years_back)
        cik = self.get_cik(ticker)

        # Load two data sources
        recent_entries = self._load_recent_entries(cik)
        hist_entries = self._load_historical_entries(cik)

        # Combine and filter
        combined = []
        for entry in recent_entries + hist_entries:
            if entry.get('form') != form_type:
                continue
            date_str = entry.get('filingDate')
            if not date_str:
                continue
            filing_date = datetime.strptime(date_str.split('T')[0], '%Y-%m-%d').date()
            if filing_date < cutoff:
                continue
            # Build link using the document name (historical) or primaryDocument (recent)
            acc = entry['accessionNumber']
            acc_nodash = acc.replace('-', '')
            link = f"{self.ARCHIVE_BASE}{int(cik)}/{acc_nodash}/{entry['document']}"

            combined.append({
                'accessionNumber': acc,
                'filingDate': date_str,
                'document': entry['document'],
                'link': link,
                'source': entry['source']
            })

        # Deduplicate on accessionNumber, preferring recent if overlap
        unique = {e['accessionNumber']: e for e in combined}
        # Sort by filingDate descending
        result = sorted(unique.values(), key=lambda x: x['filingDate'], reverse=True)
        return result

    def save_to_excel(self, filings: List[Dict], filename: str = 'filings.xlsx') -> None:
        """
        Save the list of filings to an Excel workbook, converting links to clickable formulas.
        """
        df = pd.DataFrame(filings)
        if df.empty:
            print("No filings to save.")
            return
        # Convert link column to Excel HYPERLINK formulas
        df['link'] = df['link'].apply(lambda u: f'=HYPERLINK("{u}", "{u}")')

        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            sheet_name = filename[:31]
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                max_len = df[col].astype(str).map(len).max() + 2
                ws.set_column(i, i, max_len)

        print(f"Saved {len(filings)} filings to {filename}")

# Example usage in Spyder
agent = SECFilingAgent()
filings_list = agent.get_filings('NFLX', '10-K', years_back=10)
agent.save_to_excel(filings_list, 'NFLX_10Ks_combined_10yrs.xlsx')


filings_list_10K = agent.get_filings('NFLX', '10-K', years_back=5)
filings_list_10Q = agent.get_filings('NFLX', '10-Q', years_back=5)
filings_list_8K = agent.get_filings('NFLX', '8-K', years_back=5)

print(filings_list_10K)

# ...existing code...

# def download_filing_file(filing: dict, content_key: str = "file_content") -> None:
#     """
#     Download the file from the 'link' URL in the filing dict and save its content
#     as a new key (default: 'file_content') in the same dictionary.
#     """
#     url = filing.get("link")
#     if not url or not isinstance(url, str):
#         filing[content_key] = None
#         return
#     # Remove Excel HYPERLINK formula if present
#     if url.startswith('=HYPERLINK('):
#         # Extract the actual URL from the formula
#         url = url.split('"')[1]
#     try:
#         resp = requests.get(url, headers=SECFilingAgent.HEADERS)
#         resp.raise_for_status()
#         filing[content_key] = resp.content
#     except Exception as e:
#         filing[content_key] = None
#         print(f"Failed to download {url}: {e}")

# # Example usage:
# for f in filings_list:
#     download_filing_file(f)


# Integrate load_pages to store markdown-formatted pages in the filing dict
# Assume md (markdownify) and load_pages are available in the environment
# If not, user should define/import them appropriately
from markdownify import markdownify as md

def load_pages(file: str) -> list:
    """
    Load the pages from a 10k filing.

    Args:
        file: The path to the HTML 10k filing

    Returns:
        List[str]: A list of markdown-formatted pages
    """
    # Read the file in as a string
    data = open(file, encoding="latin-1").read()

    # Convert to markdown. This removes a lot of the extra HTML
    # formatting that can be token-heavy.
    markdown_document = md(data, strip=["a", "b", "i", "u", "code", "pre"])

    # Split the document into pages
    return [page.strip() for page in markdown_document.split("\n---\n")]

def download_filing_file_markdown(filing: dict, content_key: str = "markdown_pages") -> None:
    """
    Download the file from the 'link' URL in the filing dict, convert its HTML content
    to markdown-formatted pages, and save as a new key (default: 'markdown_pages') in the dict.
    """
    url = filing.get("link")
    if not url or not isinstance(url, str):
        filing[content_key] = None
        return
    # Remove Excel HYPERLINK formula if present
    if url.startswith('=HYPERLINK('):
        url = url.split('"')[1]
    try:
        resp = requests.get(url, headers=SECFilingAgent.HEADERS)
        resp.raise_for_status()
        # Convert HTML content to string using latin-1 encoding
        html_str = resp.content.decode("latin-1", errors="replace")
        # Convert to markdown
        markdown_document = md(html_str, strip=["a", "b", "i", "u", "code", "pre"])
        # Split into pages
        pages = [page.strip() for page in markdown_document.split("\n---\n")]
        filing[content_key] = pages
    except Exception as e:
        filing[content_key] = None
        print(f"Failed to download or convert {url}: {e}")

# Example usage:
for f in filings_list:
    download_filing_file_markdown(f)

# Print each markdown page of the first filing, with page numbers and separators
first_filing = filings_list[0] if filings_list else None
if first_filing and "markdown_pages" in first_filing and first_filing["markdown_pages"]:
    for idx, page in enumerate(first_filing["markdown_pages"], start=1):
        print(f"\n{'-'*20} page {idx} {'-'*20}\n")
        print(page)
else:
    print("No markdown content available for the first filing.")


import os

# Save filings_list as a text file in the home directory
# Save to the root of the current repository
home_dir = os.path.dirname(os.path.abspath(__file__))
txt_path = os.path.join(home_dir, "filings_list.txt")
with open(txt_path, "w", encoding="utf-8") as f_txt:
    for filing in filings_list:
        # Write each filing as a line of key: value pairs
        line = "; ".join(f"{k}: {v}" for k, v in filing.items())
        f_txt.write(line + "\n")
print(f"filings_list saved as text to {txt_path}")
