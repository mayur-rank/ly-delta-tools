import pywinauto
import time
import os
import sys

# --- CONFIGURATION ---
TARGET_WINDOW = "Downloads"
LOG_FILE = "trades_detected.log"
POLLING_SPEED = 0.5

def start_scraper():
    print(f"--- Fast Scraper Started: '{TARGET_WINDOW}' ---")
    
    # 1. Create Log File Immediately so you can see it
    with open(LOG_FILE, "a") as log:
        log.write(f"\n--- SESSION STARTED AT {time.ctime()} ---\n")
    print(f"Log file created: {os.path.abspath(LOG_FILE)}")

    processed_items = set()
    
    try:
        app = pywinauto.Application(backend="uia").connect(title_re=f".*{TARGET_WINDOW}.*", timeout=5)
        window = app.window(title_re=f".*{TARGET_WINDOW}.*")
        items_view = window.child_window(title="Items View", control_type="List")
        
        print("Gathering existing items...")
        for item in items_view.items():
            processed_items.add(item.window_text())
        
        print(f"Tracking {len(processed_items)} items. GO AHEAD AND ADD A FILE TO DOWNLOADS NOW!")

        loop_count = 0
        while True:
            try:
                # HEARTBEAT: Show we are alive every 5 seconds
                loop_count += 1
                if loop_count % 10 == 0:
                    print(".", end="", flush=True)

                current_items = items_view.items()
                for item in current_items:
                    text = item.window_text()
                    
                    if text not in processed_items:
                        processed_items.add(text)
                        message = f"\n[{time.strftime('%H:%M:%S')}] NEW ITEM DETECTED: {text}"
                        print(message)
                        with open(LOG_FILE, "a") as log:
                            log.write(message + "\n")
                            
            except Exception as loop_error:
                print(f"\nWindow connection lost: {loop_error}")
                break
            
            time.sleep(POLLING_SPEED) 
            
    except KeyboardInterrupt:
        print("\nScraper stopped by user. Goodbye!")
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")

if __name__ == "__main__":
    start_scraper()
