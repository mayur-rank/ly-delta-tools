import socket
import xlwings as xw
import time
import json
import argparse

def start_leader(receiver_ip, excel_file, sheet_name):
    # Setup UDP Socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    
    print(f"Connecting to Excel: {excel_file} (Sheet: {sheet_name})...")
    try:
        wb = xw.Book(excel_file)
        sheet = wb.sheets[sheet_name]
    except Exception as e:
        print(f"FAILED to open Excel or Sheet: {e}")
        return
    
    last_processed_row = 0
    PORT = 5555
    print(f"Leader is broadcasting to {receiver_ip}:{PORT}...")

    while True:
        try:
            # Check the last row with data in Column A
            current_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            
            if current_row > last_processed_row and last_processed_row != 0:
                # Grab the new trade data
                # Assuming: Col A=Symbol, Col B=Side(BUY/SELL), Col C=Qty
                trade_data = sheet.range(f"A{current_row}:C{current_row}").value
                
                payload = {
                    "symbol": trade_data[0],
                    "side": trade_data[1],
                    "qty": trade_data[2]
                }
                
                # Send via UDP
                sock.sendto(json.dumps(payload).encode(), (receiver_ip, PORT))
                print(f"Sent: {payload}")
                
            last_processed_row = current_row
            time.sleep(0.01) # 10ms poll rate
            
        except Exception as e:
            print(f"Error: {e}")
            time.sleep(1)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Leader Trade Sender")
    parser.add_argument("--ip", type=str, default="192.168.1.XX", help="Receiver IP address")
    parser.add_argument("--excel", type=str, default="LiveTrade.xlsx", help="Excel file name")
    parser.add_argument("--sheet", type=str, default="Sheet1", help="Excel sheet name")
    
    args = parser.parse_args()
    start_leader(args.ip, args.excel, args.sheet)