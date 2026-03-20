import socket
import struct
import time
import urllib.request
import email.utils
from datetime import datetime, timezone, timedelta

def get_ntp_time(host="pool.ntp.org"):
    try:
        # Standard NTP packet is 48 bytes
        # Mode 3 = Client, Version 3
        # Reference: https://stackoverflow.com/questions/12664295/ntp-client-in-python
        port = 123
        buf = 1024
        address = (host, port)
        # Mode 3 (Client), Version 3
        msg = b'\x1b' + 47 * b'\0'
        
        # Connect to server
        client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        client.settimeout(2)
        client.sendto(msg, address)
        msg, address = client.recvfrom(buf)
        
        # Unpack the response
        # 10th element is the transmit timestamp (seconds since 1900)
        t = struct.unpack("!12I", msg)[10]
        t -= 2208988800 # Time difference between 1900 and 1970
        return t
    except Exception:
        return None

def get_http_time():
    try:
        # Use google.com as it's highly available and often whitelisted
        response = urllib.request.urlopen('http://www.google.com', timeout=2)
        date_str = response.headers['Date']
        # 'Fri, 20 Mar 2026 06:43:41 GMT'
        dt = email.utils.parsedate_to_datetime(date_str)
        return dt.timestamp()
    except Exception:
        return None

class TimeSyncer:
    def __init__(self):
        self.offset = 0
        self.last_sync_time = 0
        self.is_synced = False
        self.sync_source = "System"

    def sync(self):
        # Try NTP first
        ntp_time = get_ntp_time()
        if ntp_time:
            self.offset = ntp_time - time.time()
            self.is_synced = True
            self.sync_source = "NTP"
            self.last_sync_time = time.time()
            return True
        
        # Fallback to HTTP
        http_time = get_http_time()
        if http_time:
            self.offset = http_time - time.time()
            self.is_synced = True
            self.sync_source = "HTTP"
            self.last_sync_time = time.time()
            return True
        
        # Keep old offset if still "fresh" (e.g. within 1 hour)
        # Otherwise, we are not really synced anymore
        if time.time() - self.last_sync_time > 3600:
            # self.is_synced = False
            pass
            
        return False

    def get_current_time(self):
        # Returns (current_timestamp, is_synced, source)
        return time.time() + self.offset, self.is_synced, self.sync_source

if __name__ == "__main__":
    syncer = TimeSyncer()
    print("Syncing...")
    if syncer.sync():
        t, s, src = syncer.get_current_time()
        print(f"Synced from {src}! Offset: {syncer.offset:.2f}s")
        print(f"Real Time: {datetime.fromtimestamp(t).strftime('%H:%M:%S %p')}")
    else:
        print("Sync failed. Using system time.")
