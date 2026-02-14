"""Fetch a sample option OI Excel file for structure analysis."""

from data.fetcher import get_oi_index, download_oi_excel
from pathlib import Path

# Get 2026 OI index
entries = get_oi_index('2026')

# Find most recent entry with IndexOptions
target_entry = None
for entry in entries:
    if 'IndexOptions' in entry and entry['IndexOptions']:
        target_entry = entry
        break

if target_entry:
    print(f"Found: {target_entry['TradeDate']}")
    print(f"Path: {target_entry['IndexOptions']}")

    # Download the file
    content = download_oi_excel(target_entry['IndexOptions'])
    print(f"Downloaded: {len(content)} bytes")

    # Save to analyze
    output_path = Path("cache/oi/sample_nk225op.xlsx")
    print(f"Saved to: {output_path}")
else:
    print("No IndexOptions found")
