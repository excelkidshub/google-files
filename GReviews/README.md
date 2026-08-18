# Google Business Profile Reviews Backup System

A Python application to automatically download, track, and monitor Google Business Profile reviews with historical comparison and missing review detection.

## What This Application Does

This application helps you maintain a permanent historical record of your Google Business Profile reviews by:

1. **Downloading all current reviews** from your Google Business Profile using the official Google API
2. **Saving reviews to Excel** with organized, readable sheets
3. **Maintaining historical records** - every review ever seen is preserved permanently
4. **Detecting missing reviews** - compares current reviews against historical data to identify reviews that were previously seen but are no longer returned by Google
5. **Tracking review changes** - detects when reviewers edit their reviews (rating or text changes)
6. **Creating automatic backups** - timestamped backups and snapshots protect your data
7. **Safe error handling** - if the Google API fails, historical data is never modified

## Project Structure

```
D:\Git_ExcelKidsHub\google-files\GReviews\
├── google_reviews.py          # Main Python script
├── config.json               # Configuration file (update with your IDs)
├── requirements.txt           # Python dependencies
├── run_reviews.bat           # Easy-to-run batch file
├── README.md                 # This file
├── .gitignore               # Excludes credentials from git
├── credentials/
│   ├── README.txt           # OAuth setup instructions
│   ├── client_secret.json   # Your OAuth credentials (you provide)
│   └── token.json           # Auto-generated OAuth token
├── data/
│   ├── google_reviews.xlsx  # Main Excel report
│   ├── backup/              # Automatic backups
│   └── snapshots/           # Weekly snapshots
├── reports/                 # Additional reports
└── logs/
    └── reviews.log          # Application logs
```

## Prerequisites

- **Python 3.8 or higher** installed on your computer
- **Google Business Profile** account with reviews
- **Google Cloud Project** with appropriate APIs enabled

## Setup Instructions

### Required One-Time Setup

#### Step 1: Create Google Cloud Project

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one

#### Step 2: Enable Required APIs

1. Navigate to "APIs & Services" > "Library"
2. Search for and enable:
   - **My Business Account Management API**
   - **Business Profile Performance API**

#### Step 3: Create OAuth 2.0 Credentials

1. Go to [Google Cloud Console Credentials](https://console.cloud.google.com/apis/credentials)
2. Click "Create Credentials" > "OAuth client ID"
3. **Application type**: Desktop app
4. **Name**: Google Reviews Backup
5. Click "Create"
6. Download the JSON file
7. Rename it to `client_secret.json`
8. Place it in the `credentials/` folder

#### Step 4: Configure OAuth Consent Screen (if prompted)

1. Configure OAuth consent screen (External user type)
2. Add required scope: `https://www.googleapis.com/auth/business.manage`
3. Add your Google account as a test user
4. You can use "Testing" mode for personal use

**IMPORTANT:** You do NOT need to manually find or enter account IDs or location IDs. The application will discover them automatically.

### First Run

1. Double-click `run_reviews.bat`
2. The script will:
   - Install dependencies automatically if needed
   - Open Google OAuth authentication in your browser
   - Sign in with the Google account that manages ExcelKidsHub
   - Discover your available Google Business Profile accounts
   - Discover your business locations
   - Display the available businesses/locations
   - Let you select ExcelKidsHub (or auto-select if only one)
   - Automatically save the selected account_id and location_id to config.json
   - Download all available reviews
   - Create the first historical baseline
   - Generate `google_reviews.xlsx`

**Expected first-run output:**
```
========================================
ExcelKidsHub Google Reviews Backup
========================================

Connecting to Google Business Profile...
Discovering your Google Business Profile accounts and locations...

Found 1 Google Business Profile account(s):
  1. accounts/1234567890 (ID: 1234567890)

Automatically selected: accounts/1234567890

Found 1 location(s) in this account:
  1. ExcelKidsHub (ID: 9876543210)

Automatically selected: ExcelKidsHub

Configuration saved to config.json

Downloading reviews...
Page 1...
Successfully downloaded: 125 reviews

FIRST RUN DETECTED
Creating baseline...
Baseline created successfully.

Current reviews: 125
New reviews: 125
Potentially missing: 0
Returned: 0

Excel report:
D:\Git_ExcelKidsHub\google-files\GReviews\data\google_reviews.xlsx

========================================
Completed successfully
======================
```

### Weekly Runs

Simply double-click `run_reviews.bat` each week.

The script will:
- Use the saved account_id and location_id from config.json
- Download current reviews
- Compare with historical data
- Detect new, missing, and returned reviews
- Update the Excel report
- Create automatic backups

No manual configuration needed after the first run.

## Excel Report Structure

The main Excel file (`data/google_reviews.xlsx`) contains these sheets:

### 1. Current Reviews

Reviews returned by Google during the current run.

**Columns:**
- Review ID
- Reviewer Name
- Rating
- Review Text
- Creation Date
- Update Date
- Owner Reply
- Owner Reply Date
- Review URL
- First Seen
- Last Seen
- Current Status

### 2. Review History

Permanent historical database of all reviews ever seen.

**Important:** Reviews are NEVER deleted from this sheet, even if Google no longer returns them.

**Columns:**
- Review ID
- Reviewer Name
- Rating
- Review Text
- Creation Date
- Update Date
- Owner Reply
- Owner Reply Date
- First Seen
- Last Seen
- Current Status (PRESENT/MISSING/RETURNED)
- Times Seen

### 3. Missing Reviews

Reviews that were previously seen but are not currently returned by Google.

**Important:** These are labeled "Potentially Missing" because the API result may change temporarily.

**Columns:**
- Review ID
- Reviewer Name
- Rating
- Original Review Date
- Review Text
- First Seen
- Last Seen
- Current Status
- Times Seen
- Days Since Last Seen

**Formatting:**
- Header row frozen
- Auto-filter enabled
- MISSING status highlighted in red
- Long review text wrapped
- Dates formatted properly

### 4. Run Summary

Quick overview of the current run.

**Shows:**
- Run Date
- Location
- Current reviews found
- New reviews
- Previously known reviews
- Potentially missing reviews
- Returned reviews
- Total historical records
- Run status (BASELINE CREATED or COMPARISON COMPLETED)

### 5. Review Changes (Optional)

Tracks when reviewers edit their reviews.

**Columns:**
- Review ID
- Reviewer Name
- Field Changed (rating or review_text)
- Old Value
- New Value
- Change Date

## Understanding "Potentially Missing Reviews"

**Why "Potentially Missing"?**

The Google API result may change temporarily due to:
- API delays or caching
- Google's internal processing
- Temporary visibility changes

A review is only marked "Potentially Missing" when:
- It exists in your historical database
- It was previously observed by this application
- Its Review ID is not present in the current Google API result

**What to do:**

1. **Don't panic** - Wait a few days and run the script again
2. **Check manually** - Look at your Google Business Profile in the browser
3. **Contact Google** - If reviews are genuinely missing, contact Google Business Profile support

**Status Tracking:**
- **PRESENT**: Review is currently returned by Google
- **MISSING**: Review was seen before but not in current results
- **RETURNED**: Review was missing but has appeared again

## Backup and Safety Features

### Automatic Backups

Before modifying the historical Excel file, the script creates a timestamped backup:

```
data/backup/google_reviews_backup_2026-08-16_140500.xlsx
```

### Weekly Snapshots

One snapshot per day in the snapshots folder:

```
data/snapshots/google_reviews_2026-08-16.xlsx
```

### Error Safety

**Critical Safety Feature:** If the Google API request fails (authentication error, network error, etc.), the script:

- **Stops immediately**
- **Does NOT modify the historical database**
- **Does NOT mark reviews as missing**
- Reports: "Google review retrieval failed. Historical data was not modified."

This prevents false missing review reports due to API failures.

## Logging

All operations are logged to `logs/reviews.log`:

- Start time
- Google account/location used
- Number of reviews downloaded
- Pagination information
- New/existing/missing/returned review counts
- Errors
- Completion time

**Security:** OAuth secrets, access tokens, and passwords are NEVER logged.

## Troubleshooting

### Google API Authentication Fails

**Symptoms:**
- Error: "Credentials file not found"
- Error: "Invalid credentials"
- Browser authentication fails
- Error: "Permission denied" during account discovery

**Solutions:**

1. **Check credentials file:**
   - Ensure `credentials/client_secret.json` exists
   - Verify it's a valid OAuth client ID JSON file

2. **Re-create credentials:**
   - Go to Google Cloud Console
   - Delete the old OAuth client ID
   - Create a new one
   - Download and replace `client_secret.json`

3. **Check OAuth consent screen:**
   - Ensure your Google account is added as a test user
   - Verify the required scope is added

4. **Verify APIs are enabled:**
   - Ensure both "My Business Account Management API" and "Business Profile Performance API" are enabled
   - Check in Google Cloud Console > APIs & Services > Library

5. **Delete token.json:**
   - Delete `credentials/token.json`
   - Run the script again to re-authenticate

### "No historical records" but file exists

**Solution:** The Excel file may be corrupted. Restore from a backup in `data/backup/`.

### Script crashes with import errors

**Solution:** Install dependencies:
```bash
pip install -r requirements.txt
```

### Virtual environment issues

**Solution:** Delete the `venv` folder and run `run_reviews.bat` again to recreate it.

## Restoring from Backup

If you need to restore the historical database:

1. Go to `data/backup/`
2. Find the backup file you want to restore (e.g., `google_reviews_backup_2026-08-16_140500.xlsx`)
3. Copy it to `data/google_reviews.xlsx`
4. Run the script again

## Security Best Practices

- **Never commit** `credentials/client_secret.json` to version control
- **Never commit** `credentials/token.json` to version control
- **Never share** your OAuth credentials
- The `.gitignore` file is configured to exclude these files
- Keep your Google Cloud project secure with appropriate permissions

## Weekly Workflow

Simply double-click `run_reviews.bat` each week. That's it!

The script will automatically:
- Authenticate with Google (if needed)
- Download current reviews
- Compare with historical data
- Detect new, missing, and returned reviews
- Update the Excel report
- Create automatic backups

Review the console output and check the "Missing Reviews" sheet if any are reported.

## Technical Details

### Missing Review Detection Logic

The application uses **Review ID comparison**, not count comparison:

- ✅ **Correct:** Compare unique Google Review IDs
- ❌ **Incorrect:** Compare total review counts

A review is only marked missing when:
- Its ID exists in historical database
- It was previously observed
- Its ID is not in current API results

### Duplicate Protection

- Google Review ID is used as the unique primary key
- Running the script twice on the same day does not create duplicates
- Existing reviews update their "Last Seen" timestamp and increment "Times Seen"

### Review Edit Handling

When a reviewer edits their review:
- Same Review ID is maintained
- Review text and rating are updated
- First Seen date is preserved
- Last Seen date is updated
- Change is recorded in "Review Changes" sheet

## Support

For issues with:
- **Google API:** Check Google Cloud Console and API documentation
- **OAuth:** Follow the credentials setup instructions in `credentials/README.txt`
- **Python:** Ensure Python 3.8+ is installed
- **Excel:** Ensure you have a modern Excel version that supports .xlsx files

## License

This is a personal/business utility for ExcelKidsHub.
