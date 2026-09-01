# Google Reviews Tracker

A simple Python-based tool to track Google Business Profile reviews over time. Detects new reviews, lost reviews, and maintains historical data in Excel format.

## Features

- **NEW Reviews**: Reviews that appear for the first time
- **CURRENT Reviews**: Reviews present in the latest input
- **LOST Reviews**: Reviews that were previously seen but are missing from the latest input
- **DUPLICATES**: Detects duplicate occurrences within the input file
- **Historical Tracking**: Maintains complete history of all reviews ever seen

## File Structure

```
ReviewReport/
├── review_tracker.py      # Main script
├── all-reviews.txt        # Input file (paste Google reviews here)
├── google_reviews.xlsx    # Output Excel file (auto-generated)
└── README.md              # This file
```

## How to Use

### 1. Prepare Input File

Copy reviews from your Google Business Profile and paste them into `all-reviews.txt`.

The file should be in the format:
```
Reviews

All

Replied

Unreplied

Reviewer Nameopen_in_new
X reviews • Y photos
starstarstarstarstar 5 days ago New
Review text here...
```

### 2. Run the Script

```bash
python review_tracker.py
```

Or specify a custom input file:
```bash
python review_tracker.py your-reviews.txt
```

### 3. View Results

Open `google_reviews.xlsx` to see the results.

## Excel Sheets

The Excel file contains multiple sheets:

### All Reviews
Historical master list of all reviews ever seen.
- **Review Key**: SHA-256 hash (unique identifier)
- **Reviewer Name**: Name of the reviewer
- **Rating**: Star rating (1-5)
- **Review Date**: Date from Google (e.g., "5 days ago")
- **Review Message**: The review text
- **Reply Status**: New/Replied/Unreplied (if available)
- **First Seen**: Date when first detected
- **Last Seen**: Date when last seen
- **Status**: NEW/CURRENT/LOST

### Current
Reviews present in the latest input file.

### New
Reviews detected as NEW during the latest run.

### Lost
Reviews that were previously known but are missing from the latest input.

### Duplicates
Duplicate occurrences detected in the current input file.

## Status Meanings

- **NEW**: Review appeared for the first time in this run
- **CURRENT**: Review exists in both historical data and current input
- **LOST**: Review existed previously but is missing from current input

## Review Key

Each review is identified by a SHA-256 hash based on:
- Reviewer Name
- Rating
- Review Message

This ensures that even if the review text changes slightly, it can still be tracked.

## Workflow

1. **First Run**: All reviews are marked as NEW
2. **Subsequent Runs**: 
   - Existing reviews become CURRENT
   - New reviews are marked as NEW
   - Missing reviews are marked as LOST (but kept in history)
3. **Lost Review Returns**: If a LOST review appears again, it becomes CURRENT (historical data preserved)

## Example Console Output

```
Google Review Tracker
---------------------

Parsed 172 reviews from all-reviews.txt

Previous reviews (historical) : 172
Current reviews (in input)   : 170

NEW reviews      : 2
CURRENT reviews  : 168
LOST reviews     : 2
DUPLICATES       : 0

Excel updated:
  google_reviews.xlsx

Sheets created:
  - All Reviews (historical master list)
  - Current (reviews in latest input)
  - New (newly discovered reviews)
  - Lost (reviews missing from latest input)
```

## Requirements

- Python 3.x
- openpyxl library

Install dependencies:
```bash
pip install openpyxl
```

## Important Notes

- **LOST reviews are never deleted** - they remain in the All Reviews sheet for historical tracking
- **No duplicate master records** - the same review key will only have one row in All Reviews
- **First Seen is preserved** - when a review becomes CURRENT again, its original First Seen date is kept
- **Run regularly** - run this script weekly or whenever you update your Google reviews to maintain accurate tracking

## Troubleshooting

**No reviews parsed**: Ensure your input file follows the Google Business Profile format with `open_in_new` in the reviewer name line.

**Excel file not updating**: Delete the existing `google_reviews.xlsx` and run again to start fresh.

**Incorrect status detection**: The script compares Review Keys. If the same review has slightly different text, it may be detected as new. This is intentional to catch modified reviews.
