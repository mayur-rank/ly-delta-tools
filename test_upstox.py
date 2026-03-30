import urllib.request
import json
import gzip
import zlib

def decompress(body, encoding):
    if encoding == 'gzip':
        return gzip.decompress(body)
    elif encoding == 'deflate':
        return zlib.decompress(body, -zlib.MAX_WBITS)
    return body

url = "https://service.upstox.com/option-analytics-tool/open/v1/strategy-chains?assetKey=NSE_INDEX%7CNifty+50&strategyChainType=PC_CHAIN"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "application/json",
    "Referer": "https://upstox.com/option-chain/nifty/"
}

try:
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req) as response:
        encoding = response.info().get('Content-Encoding')
        body = response.read()
        body = decompress(body, encoding)
        data = json.loads(body.decode('utf-8'))
        print(json.dumps(data, indent=2)[:2000])
except Exception as e:
    print(f"Error: {e}")
