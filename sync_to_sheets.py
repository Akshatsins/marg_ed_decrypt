"""
Marg EDE → Google Sheets Auto Sync
Runs daily at 9:30 AM via GitHub Actions.
Fetches Marg API, decrypts, and writes fresh data to Google Sheet.
"""

import base64, json, os, sys, zlib
from Crypto.Cipher import AES
import gspread
from google.oauth2.service_account import Credentials
import requests

# ─── Config ──────────────────────────────────────────────────────────────────

SHEET_URL   = "https://docs.google.com/spreadsheets/d/1qbSn3wdmdAS9CIkm_X3NAvAKn8m5oe7chNCzX7GQz6E/edit"
SHEET_ID    = "1qbSn3wdmdAS9CIkm_X3NAvAKn8m5oe7chNCzX7GQz6E"

API_URL     = "https://corporate.margerp.com/api/eOnlineData/MargCorporateEDE"
COMPANY     = "RAKESHCORPORATE3"
MARGKEY     = "1DVQEFIU2CL4RBIILE9VDA8QCK2X2BSNX3CV"
COMPANYID   = "8820"
DECKEY      = "BXOW2BIS1IGD"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ─── Crypto ──────────────────────────────────────────────────────────────────

def decrypt_marg(data: str, key: str) -> str:
    clean = data.replace(" ", "").replace("\n", "").replace("\r", "")
    enc = base64.b64decode(clean)
    kp = (key.encode() + b"\x00" * 16)[:16]
    from Crypto.Cipher import AES
    cipher = AES.new(kp, AES.MODE_CBC, iv=kp)
    plain = cipher.decrypt(enc)
    return plain[:-plain[-1]].decode("utf-8")

def decompress_marg(b64: str) -> str:
    return zlib.decompress(base64.b64decode(b64), -15).decode("utf-8-sig")

def decrypt_and_parse(payload: str, key: str) -> dict:
    return json.loads(decompress_marg(decrypt_marg(payload, key)))

# ─── Fetch from Marg API ─────────────────────────────────────────────────────

def fetch_data() -> dict:
    print("📡 Calling Marg API...")
    resp = requests.post(
        API_URL,
        json={
            "CompanyCode": COMPANY,
            "Datetime":    "",
            "MargKey":     MARGKEY,
            "Index":       "0",
            "CompanyID":   COMPANYID,
            "APIType":     "2",
        },
        timeout=60,
    )
    resp.raise_for_status()
    raw = resp.json()

    if isinstance(raw, str):
        encrypted = raw
    elif isinstance(raw, dict):
        encrypted = (raw.get("Data") or raw.get("data") or
                     raw.get("Result") or raw.get("result") or
                     raw.get("EData") or raw.get("edata"))
    else:
        encrypted = raw

    if not isinstance(encrypted, str):
        raise ValueError(f"Unexpected API response: {str(raw)[:200]}")

    print("🔓 Decrypting...")
    return decrypt_and_parse(encrypted, DECKEY)

# ─── Upload to Google Sheets ──────────────────────────────────────────────────

def get_all_sections(data: dict) -> dict:
    """Extract all data sections as {sheet_name: [rows]}"""
    sections = {}

    # Structure A: {"Details": {"pro_N": [...], "Party": [...], ...}}
    if "Details" in data and isinstance(data["Details"], dict):
        for k, rows in data["Details"].items():
            if isinstance(rows, list) and rows:
                sections[k] = rows

    # Structure B: flat top-level lists
    TOP_KEYS = ["Masters", "MDis", "Party", "Product", "SaleType", "Stock"]
    for k in TOP_KEYS:
        rows = data.get(k)
        if isinstance(rows, list) and rows:
            sections[k] = rows

    # Fallback
    if not sections:
        for k, v in data.items():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                sections[k] = v

    return sections

def rows_to_grid(rows: list) -> list:
    """Convert list of dicts to 2D list (headers + data rows)"""
    if not rows:
        return []
    headers = list(dict.fromkeys(col for r in rows if isinstance(r, dict) for col in r.keys()))
    grid = [headers]
    for row in rows:
        grid.append([str(row.get(h, "")) for h in headers])
    return grid

def upload_to_sheets(data: dict):
    print("🔑 Authenticating with Google Sheets...")

    # Load credentials from env variable (GitHub Secret) or local file
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
        # Local testing: load from file
        with open("marg-sync-45e7d5f142d3.json") as f:
            creds_dict = json.load(f)

    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(SHEET_ID)

    sections = get_all_sections(data)
    if not sections:
        print("⚠️  No data sections found.")
        return

    # Get existing worksheets
    existing = {ws.title: ws for ws in sh.worksheets()}

    for name, rows in sections.items():
        grid = rows_to_grid(rows)
        if not grid:
            continue

        print(f"  📝 Writing '{name}': {len(rows)} records, {len(grid[0])} columns...")

        if name in existing:
            ws = existing[name]
            ws.clear()
        else:
            ws = sh.add_worksheet(title=name, rows=len(grid) + 10, cols=len(grid[0]) + 2)

        # Write all data in one API call (faster, fewer quota hits)
        ws.update(grid, value_input_option="RAW")

        # Bold the header row
        ws.format("1:1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.122, "green": 0.306, "blue": 0.475},
        })

    print(f"\n✅ Done! {len(sections)} sheet(s) updated in Google Sheets.")
    print(f"   🔗 {SHEET_URL}")

# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    try:
        data = fetch_data()
        upload_to_sheets(data)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback; traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
