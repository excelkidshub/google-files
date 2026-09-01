#!/usr/bin/env python3
"""
Standalone test for Google Business Profile Reviews API
Tests accounts.locations.reviews.list operation
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


def print_separator(char="=", length=60):
    print(char * length)


def main():
    print_separator()
    print("Google Business Profile Review Test")
    print_separator()
    print()
    
    try:
        # Load configuration
        config = load_config()
        
        account_id = config['google_business_profile']['account_id']
        location_id = config['google_business_profile']['location_id']
        location_name = config['google_business_profile']['location_name']
        
        print(f"Account ID: {account_id}")
        print(f"Location ID: {location_id}")
        print(f"Location Name: {location_name}")
        print()
        
        # Authenticate
        print("Authenticating with Google...")
        creds = authenticate(config)
        print("Authenticated successfully.")
        print()
        
        # Try to build the My Business API (Reviews API)
        print("Building My Business API service...")
        try:
            service = build('mybusiness', 'v4', credentials=creds)
            print(f"Service built successfully: {service}")
            print()
        except Exception as e:
            print(f"Error building mybusiness v4 service: {e}")
            print()
            print("Trying alternative API names...")
            
            # Try different API names and versions
            api_variations = [
                ('mybusiness', 'v4'),
                ('mybusinessaccountmanagement', 'v1'),
                ('businessprofileperformance', 'v1'),
                ('mybusinessbusinessinformation', 'v1'),
                ('mybusinessplaceinsights', 'v1'),
                ('mybusinesslodging', 'v1'),
                ('mybusinessverifications', 'v1'),
                ('mybusinessqa', 'v1'),
            ]
            
            for api_name, version in api_variations:
                try:
                    print(f"Trying {api_name} {version}...")
                    service = build(api_name, version, credentials=creds)
                    print(f"Success with {api_name} {version}")
                    print()
                    break
                except:
                    print(f"Failed with {api_name} {version}")
            else:
                raise Exception("No working API found")
        
        # Inspect the service structure
        print("Service structure inspection:")
        print(f"Service type: {type(service)}")
        print()
        
        # Check available methods
        print("Top-level methods:")
        for method in dir(service):
            if not method.startswith('_'):
                print(f"  - {method}")
        print()
        
        # Try to access accounts resource
        try:
            accounts = service.accounts()
            print("Methods on service.accounts():")
            for method in dir(accounts):
                if not method.startswith('_'):
                    print(f"  - {method}")
            print()
        except AttributeError as e:
            print(f"Error accessing accounts: {e}")
            print()
        
        # Try to access locations resource directly from service
        try:
            locations = service.locations()
            print("Methods on service.locations():")
            for method in dir(locations):
                if not method.startswith('_'):
                    print(f"  - {method}")
            print()
            
            # Try to access reviews resource
            try:
                reviews = service.locations().reviews()
                print("Methods on service.locations().reviews():")
                for method in dir(reviews):
                    if not method.startswith('_'):
                        print(f"  - {method}")
                print()
            except AttributeError as e:
                print(f"service.locations().reviews() not available: {e}")
                print()
                
        except AttributeError as e:
            print(f"Error accessing locations: {e}")
            print()
        
        print_separator()
        print("Attempting to retrieve reviews")
        print_separator()
        print()
        
        # Try to retrieve reviews using the correct API structure
        reviews = []
        page_token = None
        page_count = 0
        
        print("Attempting different API approaches...")
        print()
        
        # Approach 1: Try using accounts().list() to get location with reviews
        try:
            print("Approach 1: Using accounts().list() with readMask...")
            request = service.accounts().list()
            response = request.execute()
            
            if 'accounts' in response:
                for account in response['accounts']:
                    account_name = account.get('name')
                    print(f"  Account: {account_name}")
                    print(f"  Account keys: {list(account.keys())}")
                    
                    # Check if reviews are in the account response
                    if 'reviews' in account:
                        print(f"  Found reviews in account response!")
                        reviews.extend(account['reviews'])
            print()
        except Exception as e:
            print(f"  Approach 1 failed: {e}")
            print()
        
        # Approach 2: Try using accounts().get() with readMask
        try:
            print("Approach 2: Using accounts().get() with readMask...")
            request = service.accounts().get(
                name=f"accounts/{account_id}",
                readMask="name,accountName,reviews"
            )
            response = request.execute()
            
            print(f"  Response keys: {list(response.keys())}")
            
            if 'reviews' in response:
                print(f"  Found reviews in account response!")
                reviews.extend(response['reviews'])
            print()
        except Exception as e:
            print(f"  Approach 2 failed: {e}")
            print()
        
        # Approach 3: Try direct REST API call using http
        try:
            print("Approach 3: Using direct HTTP request to Reviews API...")
            import google.auth.transport.requests as http_request
            
            http = http_request.AuthorizedSession(creds)
            
            # Try the My Business API v4 reviews endpoint
            url = f"https://mybusiness.googleapis.com/v4/accounts/{account_id}/locations/{location_id}/reviews"
            response = http.get(url)
            
            print(f"  HTTP Status: {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                print(f"  Response keys: {list(data.keys())}")
                
                if 'reviews' in data:
                    reviews.extend(data['reviews'])
                    print(f"  Retrieved {len(data['reviews'])} reviews via direct HTTP")
                    
                    # Handle pagination for direct HTTP
                    while 'nextPageToken' in data:
                        next_page_token = data['nextPageToken']
                        print(f"  Fetching next page with token: {next_page_token[:20]}...")
                        
                        paginated_url = f"{url}?pageToken={next_page_token}"
                        response = http.get(paginated_url)
                        
                        if response.status_code == 200:
                            data = response.json()
                            if 'reviews' in data:
                                reviews.extend(data['reviews'])
                                print(f"  Retrieved {len(data['reviews'])} more reviews")
                            else:
                                break
                        else:
                            break
            else:
                print(f"  HTTP Error: {response.text}")
            print()
        except Exception as e:
            print(f"  Approach 3 failed: {e}")
            import traceback
            traceback.print_exc()
            print()
        
        # Approach 4: Try Business Profile Performance API for review metrics
        try:
            print("Approach 4: Using Business Profile Performance API...")
            service_perf = build('businessprofileperformance', 'v1', credentials=creds)
            
            print("Methods on service_perf.locations():")
            for method in dir(service_perf.locations()):
                if not method.startswith('_'):
                    print(f"  - {method}")
            print()
            
            # Try to get review metrics
            request = service_perf.locations().getDailyMetricsTimeSeries(
                name=f"locations/{location_id}",
                dailyMetrics=["REVIEWS"],
                startDate="2020-01-01",
                endDate="2026-12-31"
            )
            
            response = request.execute()
            print(f"  Response keys: {list(response.keys())}")
            print()
        except Exception as e:
            print(f"  Approach 4 failed: {e}")
            print()
                
        except AttributeError as e:
            print(f"AttributeError: {e}")
            print("The reviews resource or list method may not be available in this API version.")
        except HttpError as e:
            print(f"Google API HTTP Error: {e.resp.status}")
            print(f"Details: {e}")
        except Exception as e:
            print(f"Error retrieving reviews: {e}")
            import traceback
            traceback.print_exc()
        
        print()
        print_separator()
        print("TEST RESULTS")
        print_separator()
        print(f"Total reviews downloaded: {len(reviews)}")
        print(f"Pages retrieved: {page_count}")
        print()
        
        if reviews:
            print("Sample review data:")
            if len(reviews) > 0:
                first_review = reviews[0]
                print(f"  Review ID: {first_review.get('name', 'N/A')}")
                print(f"  Reviewer: {first_review.get('reviewer', {}).get('displayName', 'N/A')}")
                print(f"  Rating: {first_review.get('starRating', 'N/A')}")
                print(f"  Comment: {first_review.get('comment', 'N/A')[:100]}...")
        
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
