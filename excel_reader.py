import win32com.client
import pythoncom

class ExcelReader:
    def __init__(self):
        self.wb_name = None
        self.sheet_name = None
        self.cells = []

    def set_config(self, wb_name, sheet_name, cell1, cell2, cell3):
        self.wb_name = wb_name
        self.sheet_name = sheet_name
        self.cells = [cell1, cell2, cell3]

    def append_row(self, sheet_name, row_data, header_data=None, metadata_data=None):
        """Append a row of data to the specified sheet. Creates sheet and adds metadata/header if needed."""
        if not self.wb_name:
            return False

        target_wb = self.wb_name.replace("/", "\\").lower()
        is_full_path = "\\" in target_wb or ":" in target_wb

        pythoncom.CoInitialize()
        try:
            excel = None
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
            except:
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                except:
                    return False
            
            if not excel:
                return False

            # Find the workbook
            wb = None
            for w in excel.Workbooks:
                fullname = w.FullName.lower()
                name = w.Name.lower()
                if is_full_path:
                    if fullname == target_wb:
                        wb = w
                        break
                else:
                    if name == target_wb:
                        wb = w
                        break
            
            if not wb:
                try:
                    wb = excel.Workbooks(self.wb_name)
                except:
                    return False

            # Find or create sheet
            sheet = None
            try:
                sheet = wb.Sheets(sheet_name)
                newly_created = False
            except:
                try:
                    # Create new sheet at the end
                    sheet = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
                    sheet.Name = sheet_name
                    newly_created = True
                except:
                    return False

            if not sheet:
                return False

            # If newly created, write metadata and header
            if newly_created:
                if metadata_data:
                    for idx, val in enumerate(metadata_data, 1):
                        sheet.Cells(1, idx).Value = val
                if header_data:
                    for idx, val in enumerate(header_data, 1):
                        sheet.Cells(3, idx).Value = val
                last_row = 4
            else:
                # Find the last used row in column A
                try:
                    # xlUp depends on the Excel instance, using constant -4162
                    last_row = sheet.Cells(sheet.Rows.Count, 1).End(-4162).Row + 1
                    # Ensure we don't overwrite header if it's the first data row
                    if last_row < 4:
                        last_row = 4
                except:
                    # Fallback if End fails
                    last_row = 4
                    while sheet.Cells(last_row, 1).Value:
                        last_row += 1

            # Write row data
            for idx, val in enumerate(row_data, 1):
                sheet.Cells(last_row, idx).Value = val
            
            return True

        except Exception as e:
            print(f"Error appending to Excel: {e}")
            return False
        finally:
            pythoncom.CoUninitialize()

    def read_cells(self):
        if not self.wb_name or not self.cells:
            return "Set", "Excel", "Config"

        # Normalize path for matching
        target_wb = self.wb_name.replace("/", "\\").lower()
        
        # Determine if we are looking for a full path or just a filename
        is_full_path = "\\" in target_wb or ":" in target_wb

        max_retries = 3
        for attempt in range(max_retries):
            try:
                pythoncom.CoInitialize()
                
                excel = None
                # Try to get existing object first
                try:
                    excel = win32com.client.GetActiveObject("Excel.Application")
                except:
                    # Fallback to Dispatch (can sometimes find it when GetActiveObject fails)
                    try:
                        excel = win32com.client.Dispatch("Excel.Application")
                    except:
                        return "Err", "Exc", "Not Found"
                
                if not excel:
                    return "Err", "Exc", "Not Found"

                # Find the workbook
                wb = None
                try:
                    for w in excel.Workbooks:
                        fullname = w.FullName.lower()
                        name = w.Name.lower()
                        
                        if is_full_path:
                            if fullname == target_wb:
                                wb = w
                                break
                        else:
                            if name == target_wb:
                                wb = w
                                break
                except:
                    # Excel might be busy or in a state where Workbooks is inaccessible
                    import time
                    time.sleep(0.05)
                    continue
                
                if not wb:
                    # If not found in open workbooks, try to see if we can just grab it by name
                    # (Sometimes hidden or special workbooks don't show in the list but are accessible)
                    try:
                        wb = excel.Workbooks(self.wb_name)
                    except:
                        return "Err", "WB", "Not Open"

                # Use sheet name or active sheet
                sheet = None
                try:
                    if self.sheet_name:
                        sheet = wb.Sheets(self.sheet_name)
                    else:
                        sheet = wb.ActiveSheet
                except:
                    return "Err", "Sheet", "Not Found"

                if not sheet:
                    return "Err", "Sheet", "Not Found"

                # Reading cells with error handling for Each cell
                # If Odin is updating, some cells might be temporarily "Busy"
                try:
                    val1 = sheet.Range(self.cells[0]).Value
                    val2 = sheet.Range(self.cells[1]).Value
                    val3 = sheet.Range(self.cells[2]).Value
                    
                    return (
                        "" if val1 is None else val1,
                        "" if val2 is None else val2,
                        "" if val3 is None else val3
                    )
                except:
                    # Cell range might be busy
                    raise Exception("Busy")

            except Exception as e:
                error_code = getattr(e, 'hresult', 0)
                # COM Busy codes
                if error_code in [-2147418111, -2147417851] or "busy" in str(e).lower():
                    import time
                    time.sleep(0.05)
                    continue
                else:
                    return "Err", "Exc", "Busy/Fail"
            finally:
                pythoncom.CoUninitialize()
        
        return "Err", "Exc", "Busy"
