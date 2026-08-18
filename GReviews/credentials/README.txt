CREDENTIALS SETUP INSTRUCTIONS
==============================

This directory must contain your Google OAuth credentials for the Google Business Profile API.

STEP 1: Create Google Cloud Project
-----------------------------------
1. Go to https://console.cloud.google.com/
2. Create a new project or select an existing one

STEP 2: Enable Required APIs
----------------------------
Enable these APIs in Google Cloud Console:

1. Go to "APIs & Services" > "Library"
2. Search for and enable:
   - **My Business Account Management API**
   - **Business Profile Performance API**

STEP 3: Create OAuth 2.0 Credentials
------------------------------------
1. Go to https://console.cloud.google.com/apis/credentials
2. Click "Create Credentials" > "OAuth client ID"
3. Application type: "Desktop app"
4. Name: "Google Reviews Backup"
5. Click "Create"
6. Download the JSON file
7. Rename it to "client_secret.json"
8. Place it in this directory (credentials/)

STEP 4: Configure OAuth Consent Screen
--------------------------------------
If prompted:
1. Configure OAuth consent screen (External user type)
2. Add required scope:
   - https://www.googleapis.com/auth/business.manage
3. Add your Google account as a test user
4. You can use "Testing" mode for personal use

IMPORTANT: NO MANUAL IDs NEEDED
--------------------------------
You do NOT need to manually find or enter account IDs or location IDs.

The application will automatically:
- Discover your Google Business Profile accounts
- Discover your business locations
- Let you select ExcelKidsHub
- Save the IDs to config.json automatically

IMPORTANT SECURITY NOTES:
------------------------
- Never commit client_secret.json to version control
- Never share your OAuth credentials
- The .gitignore file is configured to ignore these files
- Never expose credentials in logs or output

FILE LOCATION:
--------------
Place your client_secret.json here:
D:\Git_ExcelKidsHub\google-files\GReviews\credentials\client_secret.json

The script will automatically create token.json in this directory after successful authentication.
