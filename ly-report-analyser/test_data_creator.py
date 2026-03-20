import win32com.client
import os
from datetime import datetime, timedelta
import random
import pythoncom

def generate_test_data(folder_name="v_test_data", num_days=5):
    """
    Generates a set of Excel files simulating daily trading logs using win32com.
    Injects patterns like drops at 09:18:00 AM and 03:25:05 PM.
    """
    abs_folder = os.path.abspath(folder_name)
    if not os.path.exists(abs_folder):
        os.makedirs(abs_folder)
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        start_date = datetime.now() - timedelta(days=num_days)
        types = ["BSC", "NSC"]
        
        for day_offset in range(num_days):
            current_date = start_date + timedelta(days=day_offset)
            if current_date.weekday() >= 5: continue # Skip weekends
            
            date_str = current_date.strftime('%Y-%m-%d')
            weekday_str = current_date.strftime('%A')
            
            for t in types:
                filename = os.path.join(abs_folder, f"Logs_{t}_{date_str}.xlsx")
                wb = excel.Workbooks.Add()
                sheet = wb.ActiveSheet
                sheet.Name = date_str
                
                # Metadata Row 1
                sheet.Cells(1, 1).Value = f"Date: {date_str}"
                sheet.Cells(1, 2).Value = f"Day: {weekday_str}"
                sheet.Cells(1, 3).Value = f"Type: {t}"
                
                # Headers Row 3
                headers = ["DateTime", "Premium", "Difference"]
                for i, h in enumerate(headers, 1):
                    sheet.Cells(3, i).Value = h
                
                # Data generation (9:15 to 3:30, 1 min intervals for speed)
                curr_time = current_date.replace(hour=9, minute=15, second=0, microsecond=0)
                end_time = current_date.replace(hour=15, minute=30, second=0, microsecond=0)
                
                row = 4
                base_val = random.uniform(100, 500)
                prev_val = base_val
                
                while curr_time <= end_time:
                    time_str = curr_time.strftime('%I:%M:%S %p')
                    
                    # Pattern Injection
                    change = random.uniform(-0.1, 0.1)
                    if curr_time.hour == 9 and curr_time.minute == 18:
                        change = random.uniform(-4.0, -2.0)
                    if curr_time.hour == 15 and curr_time.minute == 25:
                        change = random.uniform(-3.0, -2.0)
                    
                    val = round(prev_val + change, 2)
                    diff = round(val - prev_val, 2)
                    
                    sheet.Cells(row, 1).Value = time_str
                    sheet.Cells(row, 2).Value = val
                    sheet.Cells(row, 3).Value = diff
                    
                    prev_val = val
                    curr_time += timedelta(minutes=1) # 1 min for faster test data creation
                    row += 1
                
                wb.SaveAs(filename)
                wb.Close()
                print(f"Generated {filename}")
                
        excel.Quit()
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    generate_test_data()
