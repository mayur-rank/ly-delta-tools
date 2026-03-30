import sys
import os
import time

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from ly_premium_fetcher import NSEDataFetcher
    
    print("Initializing Fetcher...")
    fetcher = NSEDataFetcher()
    
    print("Fetching Nifty Spot...")
    spot = fetcher.get_nifty_spot()
    print(f"Nifty Spot: {spot}")
    
    if spot:
        print("Fetching ATM Premium...")
        strike, expiry, ce, pe = fetcher.get_atm_premium(spot)
        print(f"ATM Strike: {strike}")
        print(f"Expiry: {expiry}")
        print(f"CE LTP: {ce}")
        print(f"PE LTP: {pe}")
        print(f"Total Premium: {ce + pe}")
        print("SUCCESS")
    else:
        print("FAILED to fetch spot price.")
        
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
