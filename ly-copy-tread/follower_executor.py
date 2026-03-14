import socket
import json
import win32gui
import win32api
import win32con
import time
import pyautogui

# --- CONFIGURATION ---
LISTEN_IP = "0.0.0.0"  # Listens on all network interfaces
PORT = 5555            # Must match the port on the Leader PC
ODIN_TITLE = "ODIN Client Ver 10.0.5.0 [Powered by SynapseWave] Jainam Broking Ltd" # Updated from screenshot

# Optimization: Reduce PyAutoGUI's internal safety delays
pyautogui.PAUSE = 0.01 

def get_odin_hwnd():
    """Locates the ODIN window handle using fuzzy title matching."""
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if ODIN_TITLE.lower() in title.lower():
                windows.append(hwnd)
        return True

    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows[0] if windows else None

def execute_trade(side, scrip_name, qty):
    """
    Performs the 3-step execution:
    1. Search Strike (Ctrl+S)
    2. Open Order Window (F1/F2)
    3. Fill & Submit (Qty + Enter)
    """
    hwnd = get_odin_hwnd()
    if not hwnd:
        print("ALERT: ODIN Diet window not found. Please open the terminal.")
        return

    try:
        # STEP 1: SEARCH & SET STRIKE (Ctrl + S)
        # This ensures the follower is on the exact same contract as the leader
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, ord('S'), 0)
        time.sleep(0.02)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, ord('S'), 0)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        
        time.sleep(0.05) # Wait for search bar to focus
        pyautogui.typewrite(scrip_name)
        pyautogui.press('enter')
        time.sleep(0.05) # Wait for scrip to be selected in Market Watch

        # STEP 2: TRIGGER ORDER ENTRY
        # F1 = Buy, F2 = Sell
        v_key = win32con.VK_F1 if side.upper() == "BUY" else win32con.VK_F2
        win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, v_key, 0)
        win32api.PostMessage(hwnd, win32con.WM_KEYUP, v_key, 0)
        
        time.sleep(0.08) # Wait for Buy/Sell window to pop up

        # STEP 3: FILL QUANTITY AND SUBMIT
        pyautogui.typewrite(str(int(qty)))
        pyautogui.press('enter')
        
        print(f"SUCCESS: {side} {qty} units of {scrip_name}")

    except Exception as e:
        print(f"EXECUTION ERROR: {e}")

def start_follower():
    """Initializes the UDP listener."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind((LISTEN_IP, PORT))
    
    print(f"Follower active on port {PORT}. Waiting for Leader signals...")
    
    while True:
        data, addr = sock.recvfrom(1024)
        try:
            trade = json.loads(data.decode())
            print(f"Incoming Signal: {trade}")
            
            # Extract data from the JSON packet sent by PC 1
            execute_trade(
                side=trade.get('side'),
                scrip_name=trade.get('symbol'),
                qty=trade.get('qty')
            )
        except Exception as e:
            print(f"Packet Error: {e}")

if __name__ == "__main__":
    start_follower()