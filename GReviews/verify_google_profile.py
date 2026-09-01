#!/usr/bin/env python3
"""
Google Business Profile Account and Location Verification
Verifies account_id and location_id for ExcelKidsHub Phonics Academy
"""

import os
import json
import ssl
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Fix SSL certificate issue on Windows
ssl._create_default_https_context = ssl._create_unverified_context


def load_config():
    """Load configuration from config.json."""
    with open('config.json', 'r') as f:
        return json.load(f)


def authenticate(config):
    """Authenticate with Google OAuth 2.0 using existing credentials."""
    creds = None
    token_path = config['api']['token_file']
    credentials_path = config['api']['credentials_file']
    scopes = config['api']['scopes']
    
    if os.path.exists(token_path):
        with open(token_path, 'r') as f:
            creds_data = json.load(f)
        creds = Credentials.from_authorized_user_info(creds_data, scopes)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_path):
                raise FileNotFoundError(
                    f"Credentials file not found: {credentials_path}\n"
                    "Please follow the setup instructions in credentials/README.txt"
                )
            print("Starting OAuth authentication flow...")
            print("Please authorize in your browser...")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, scopes)
            creds = flow.run_local_server(port=0)
        
        with open(token_path, 'w') as f:
            json.dump(json.loads(creds.to_json()), f, indent=2)
    
    return creds


def list_accounts(mybusiness_service):
    """List all Google Business Profile accounts."""
    try:
        accounts_response = mybusiness_service.accounts().list().execute()
        return accounts_response.get('accounts', [])
    except HttpError as e:
        print(f"Error listing accounts: HTTP {e.resp.status}")
        print(f"Details: {e}")
        raise


def print_separator(char="=", length=60):
    print(char * length)


def main():
    print_separator()
    print("Google Business Profile Verification")
    print_separator()
    print()
    
    try:
        # Load configuration
        config = load_config()
        
        # Authenticate
        print("Authenticating with Google...")
        creds = authenticate(config)
        print("Authenticated successfully.")
        print()
        
        # Build My Business Account Management API service
        mybusiness_service = build('mybusinessaccountmanagement', 'v1', credentials=creds)
        
        # List accounts
        print("Accounts found:")
        print_separator("-")
        accounts = list_accounts(mybusiness_service)
        
        if not accounts:
            print("No accounts found.")
            print()
            print_separator()
            print("VERIFICATION RESULT")
            print_separator()
            print("No Google Business Profile accounts found.")
            print("Please ensure you have at least one business location set up.")
            return
        
        target_location_name = "ExcelKidsHub Phonics Academy"
        found_location = None
        found_account = None
        
        for account in accounts:
            account_name = account.get('name', 'Unknown')
            account_id = account_name.split('/')[-1] if '/' in account_name else account_name
            account_display_name = account.get('accountName', 'Unknown')
            print(f"Account ID: {account_id}")
            print(f"Account name: {account_display_name}")
            print()
            
            # Try to get locations for this account
            print(f"Locations for account {account_id}:")
            print_separator("-")
            
            try:
                # Try different methods to get locations
                # Method 1: Check if locations are in account response
                locations = account.get('locations', [])
                
                if locations:
                    for location in locations:
                        loc_name = location.get('locationName', 'Unknown')
                        loc_id = location.get('name', '').split('/')[-1] if '/' in location.get('name', '') else location.get('name', '')
                        print(f"Location ID: {loc_id}")
                        print(f"Location name: {loc_name}")
                        
                        if target_location_name.lower() in loc_name.lower():
                            found_location = location
                            found_account = account
                            print(f"  -> MATCH FOUND: {loc_name}")
                        print()
                else:
                    print("No locations in account response (API may not include them)")
                    print()
                    
            except Exception as e:
                print(f"Error getting locations: {e}")
                print()
        
        print_separator()
        print("VERIFICATION RESULT")
        print_separator()
        
        if found_location and found_account:
            account_name = found_account.get('name', 'Unknown')
            account_id = account_name.split('/')[-1] if '/' in account_name else account_name
            account_display_name = found_account.get('accountName', 'Unknown')
            
            loc_name = found_location.get('locationName', 'Unknown')
            loc_id = found_location.get('name', '').split('/')[-1] if '/' in found_location.get('name', '') else found_location.get('name', '')
            
            print(f"ExcelKidsHub Phonics Academy found successfully.")
            print()
            print(f"Account ID: {account_id}")
            print(f"Account name: {account_display_name}")
            print(f"Location ID: {loc_id}")
            print(f"Location name: {loc_name}")
        else:
            print(f"ExcelKidsHub Phonics Academy was not found.")
            print()
            print("Available locations are shown above.")
            print("If the location exists but is not listed, the API may not be returning location data.")
            print("You may need to:")
            print("1. Check Google Cloud Console API quotas")
            print("2. Verify OAuth consent screen configuration")
            print("3. Manually enter location_id in config.json")
        
        print_separator()
        
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except HttpError as e:
        print(f"Google API Error: HTTP {e.resp.status}")
        print(f"Details: {e}")
        print()
        if e.resp.status == 403:
            print("Permission denied. Please ensure:")
            print("1. My Business Account Management API is enabled")
            print("2. OAuth client has correct permissions")
            print("3. Your Google account has access to Google Business Profile")
        elif e.resp.status == 429:
            print("Rate limit exceeded. Please check API quotas in Google Cloud Console.")
        elif e.resp.status == 401:
            print("Authentication failed. Please check OAuth credentials.")
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
