import socket
import struct
import time
import urllib.request
import email.utils
from datetime import datetime, timezone, timedelta

def get_ntp_time(host="pool.ntp.org"):
    try:
        # Standard NTP packet is 48 bytes
        port = 123
        buf = 1024
        address = (host, port)
        # Mode 3 (Client), Version 3
        msg = b'\x1b' + 47 * b'\0'
        
        # Connect to server
        client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        client.settimeout(2.5) # Slightly longer timeout for overseas servers
        
        t_send = time.time()
        client.sendto(msg, address)
        msg, address = client.recvfrom(buf)
        t_recv = time.time()
        
        # Unpack the response
        # Indices 10 (seconds) and 11 (fraction) are the transmit timestamp (server send)
        unpacked = struct.unpack("!12I", msg)
        t_secs = unpacked[10] - 2208988800 # Time difference 1900 to 1970
        t_frac = unpacked[11] / (2**32)
        t_server_xmit = t_secs + t_frac
        
        # RTT compensation: Assume symmetric network delay
        # Offset = (ServerTime + RTT/2) - ClientTimeAtRecv
        rtt = t_recv - t_send
        offset = t_server_xmit + (rtt / 2.0) - t_recv
        
        return offset
    except Exception:
        return None

def get_http_time():
    """Fallback: High-precision HTTP time via WorldTimeAPI."""
    try:
        # Asia/Kolkata for IST
        url = "http://worldtimeapi.org/api/timezone/Asia/Kolkata"
        t_send = time.time()
        with urllib.request.urlopen(url, timeout=3) as response:
            data = json.loads(response.read().decode())
            # datetime: "2026-03-30T18:19:00.123456+05:30"
            dt_str = data.get('datetime')
            # Extract fractional seconds (WorldTimeAPI provides microseconds)
            dt = datetime.fromisoformat(dt_str)
            t_server = dt.timestamp()
            t_recv = time.time()
            
            rtt = t_recv - t_send
            offset = t_server + (rtt / 2.0) - t_recv
            return offset, rtt
    except Exception as e:
        print(f"WorldTimeAPI Fallback Error: {e}")
        return None, 0

class TimeSyncer:
    def __init__(self):
        self.offset = 0.0
        self.last_sync_time = 0.0
        self.is_synced = False
        self.sync_source = "System"
        self.last_rtt = 0.0
        self.manual_offset = 0.0 # Fine-tuning in seconds
        self.servers = [
            "in.pool.ntp.org",
            "time.google.com",
            "time.cloudflare.com",
            "pool.ntp.org"
        ]

    def set_manual_adjustment(self, ms_offset):
        """Sets a manual fine-tuning offset in milliseconds."""
        self.manual_offset = float(ms_offset) / 1000.0

    def sync(self):
        # Multi-sample strategy: Take up to 5 samples per server and pick the best one
        samples = []
        for server in self.servers:
            server_samples = []
            for _ in range(3): # Take 3 quick samples per server
                # We need a slightly modified get_ntp_time that returns (offset, rtt)
                result = self._get_ntp_sample(server)
                if result:
                    server_samples.append(result)
            
            if server_samples:
                # Pick the sample with the lowest RTT for this server
                best_sample = min(server_samples, key=lambda x: x[1])
                samples.append((best_sample[0], server, best_sample[1]))
                # If we have a good sample from a primary server, we can stop
                if best_sample[1] < 0.1: # 100ms RTT is good enough
                    break
        
        if samples:
            # use the sample from the most reliable server (first in list that worked)
            self.offset = samples[0][0]
            self.is_synced = True
            self.sync_source = "NTP"
            self.last_rtt = samples[0][2] # Best sample's RTT
            self.last_sync_time = time.time()
            return True
        
        # Fallback to HTTP (High-Precision fallback)
        offset, rtt = get_http_time()
        if offset is not None:
            self.offset = offset
            self.is_synced = True
            self.sync_source = "HTTP (WorldTimeAPI)"
            self.last_rtt = rtt
            self.last_sync_time = time.time()
            return True
        
        return False

    def _get_ntp_sample(self, host):
        """Returns (offset, rtt) for a single NTP request."""
        try:
            port = 123
            buf = 1024
            address = (host, port)
            msg = b'\x1b' + 47 * b'\0'
            client = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            client.settimeout(1.5)
            
            t_send = time.time()
            client.sendto(msg, address)
            msg, address = client.recvfrom(buf)
            t_recv = time.time()
            
            unpacked = struct.unpack("!12I", msg)
            t_secs = unpacked[10] - 2208988800
            t_frac = unpacked[11] / (2**32)
            t_server_xmit = t_secs + t_frac
            
            rtt = t_recv - t_send
            offset = t_server_xmit + (rtt / 2.0) - t_recv
            return offset, rtt
        except Exception:
            return None

    def get_current_time(self):
        # Returns (current_timestamp_including_manual, is_synced, source)
        return time.time() + self.offset + self.manual_offset, self.is_synced, self.sync_source

if __name__ == "__main__":
    syncer = TimeSyncer()
    print("Syncing...")
    if syncer.sync():
        t, s, src = syncer.get_current_time()
        print(f"Synced from {src}! Offset: {syncer.offset:.2f}s")
        print(f"Real Time: {datetime.fromtimestamp(t).strftime('%H:%M:%S %p')}")
    else:
        print("Sync failed. Using system time.")
