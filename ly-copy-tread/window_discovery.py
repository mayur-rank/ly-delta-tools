import pywinauto
from pywinauto import Desktop
import sys
import os

def list_windows():
    print("\n--- Listing All Visible Windows ---")
    windows = Desktop(backend="win32").windows()
    for i, w in enumerate(windows):
        try:
            text = w.window_text()
            if text:
                print(f"[{i}] Title: {text} | Class: {w.class_name()}")
        except:
            pass
    print("-" * 35)

def inspect_window(title_keyword):
    log_file = "window_layout.txt"
    print(f"\n--- Deep Inspecting Window: '{title_keyword}' ---")
    
    for backend in ["uia", "win32"]:
        print(f"Trying backend: {backend}...")
        try:
            app = pywinauto.Application(backend=backend).connect(title_re=f".*{title_keyword}.*", timeout=3)
            main_win = app.window(title_re=f".*{title_keyword}.*")
            
            print(f"SUCCESS! Connected using {backend}.")
            
            # Find all sub-windows (dialogs) inside
            children = main_win.children()
            print(f"\nFound {len(children)} sub-elements/windows inside.")
            for c in children:
                try:
                    name = c.window_text()
                    if name:
                        print(f"   >> Sub-Window Found: '{name}'")
                except:
                    pass

            print(f"\nSaving FULL LAYOUT to {log_file}...")
            with open(log_file, "w", encoding="utf-8") as f:
                original_stdout = sys.stdout
                sys.stdout = f
                try:
                    main_win.print_control_identifiers()
                finally:
                    sys.stdout = original_stdout
            
            print(f"DONE! Please check {os.path.abspath(log_file)} for details.")
            return
        except Exception as e:
            print(f"Failed with {backend}: {e}")
    
    print("\n[!] Could not connect to window. Make sure it's open and visible.")

if __name__ == "__main__":
    list_windows()
    keyword = input("\nEnter a keyword from the list above: ")
    if keyword:
        inspect_window(keyword)
