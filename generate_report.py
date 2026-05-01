"""
SMS Campaign — Consolidated Excel Report Generator
Generates a detailed, styled Excel report with:
  1. Summary Dashboard
  2. SmartCard by LO (day-wise detail)
  3. SmartCard raw data per LO (separate sheets)
  4. Renewal 30 Days summary by district
  5. Renewal 30 Days raw data

Supports reading data from:
  - Local filesystem (default)
  - Google Drive shared folder (via service account)

Usage:
  python generate_report.py                          # local mode (default)
  python generate_report.py --source drive            # drive mode (uses config.json)
  python generate_report.py --source drive --folder-id <ID> --key service_account.json
"""

import os
import io
import json
import argparse
import tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from collections import OrderedDict
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SC_DIR = os.path.join(BASE_DIR, "SmartCard")
RENEWAL_FILE = os.path.join(BASE_DIR, "Renewal_30days_30th_April.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "SMS_Consolidated_Report.xlsx")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

# --- Styling constants ---
DARK_BLUE = "1B2A4A"
MED_BLUE = "2D4373"
LIGHT_BLUE = "D6E4F0"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
GREEN = "27AE60"
CYAN = "17A2B8"
ORANGE = "E67E22"
RED = "E74C3C"

header_font = Font(name="Calibri", bold=True, color=WHITE, size=12)
header_fill = PatternFill(start_color=DARK_BLUE, end_color=DARK_BLUE, fill_type="solid")
sub_header_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
sub_header_fill = PatternFill(start_color=MED_BLUE, end_color=MED_BLUE, fill_type="solid")
title_font = Font(name="Calibri", bold=True, size=16, color=DARK_BLUE)
subtitle_font = Font(name="Calibri", bold=True, size=13, color=MED_BLUE)
bold_font = Font(name="Calibri", bold=True, size=11)
normal_font = Font(name="Calibri", size=11)
number_font = Font(name="Calibri", size=11)
total_fill = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
total_font = Font(name="Calibri", bold=True, size=11, color=DARK_BLUE)
alt_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
center = Alignment(horizontal="center", vertical="center")
left_align = Alignment(horizontal="left", vertical="center")
right_align = Alignment(horizontal="right", vertical="center")

# Mapping for LO file names to readable names
LO_NAME_MAP = OrderedDict([
    ("LO1_Makali", "LO1 – Makali"),
    ("LO2_Kengeri", "LO2 – Kengeri"),
    ("LO3_Marathalli", "LO3 – Marathalli"),
    ("LO4_HBR_Layout", "LO4 – HBR Layout"),
    ("LO5_RT_Nagar", "LO5 – RT Nagar"),
    ("LO6_Chandapura", "LO6 – Chandapura"),
    ("LO7_Shanti Nagar", "LO7 – Shanti Nagar"),
    ("LO8_Jakkur", "LO8 – Jakkur"),
    ("LO_Bng_Rural_Davanahalli", "LO – Devanahalli (Bng Rural)"),
    ("LO_Mysore", "LO – Mysore"),
    ("Mysore", "LO – Mysore"),
])

RENEWAL_LOC_MAP = {
    'BGM': 'Belagavi', 'DVG': 'Davanagere', 'MYS': 'Mysuru', 'DWR': 'Dharwad',
    'BLU': 'Ballari', 'HVR': 'Haveri', 'BGK': 'Bagalkot', 'KPL': 'Koppal',
    'UTK': 'Uttara Kannada', 'TMK': 'Tumakuru', 'CTD': 'Chitradurga', 'KLR': 'Kolar',
    'SMG': 'Shivamogga', 'DKN': 'Dakshina Kannada', 'RCR': 'Raichur', 'BDR': 'Bidar',
    'BJR': 'Vijayapura', 'CKM': 'Chikkamagaluru', 'GDG': 'Gadag', 'HSN': 'Hassan',
    'KLB': 'Kalaburagi', 'CMJ': 'Chamarajanagar', 'VIJ': 'Vijayanagara', 'BLY': 'Bengaluru Rural',
    'UDP': 'Udupi', 'MND': 'Mandya', 'CBP': 'Chikkaballapur', 'YDR': 'Yadgir',
    'RMR': 'Ramanagara', 'BLR': 'Bengaluru Urban', '100': 'Unknown (100)', 'MDK': 'Kodagu',
}

DAY_FOLDERS = ["28th April", "29th April", "30th April"]


# ---------------------------------------------------------------------------
#  Google Drive Reader
# ---------------------------------------------------------------------------

class GoogleDriveReader:
    """Reads files from a shared Google Drive folder using a service account."""

    def __init__(self, service_account_key_path):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        creds = service_account.Credentials.from_service_account_file(
            service_account_key_path,
            scopes=['https://www.googleapis.com/auth/drive.readonly']
        )
        self.service = build('drive', 'v3', credentials=creds)
        print("   ✅ Google Drive authenticated successfully")

    def list_folders(self, parent_id):
        """List sub-folders inside a folder. Returns {name: id}."""
        results = self.service.files().list(
            q=f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id, name)",
            pageSize=100
        ).execute()
        return {f['name']: f['id'] for f in results.get('files', [])}

    def list_files(self, parent_id, extension=".xlsx"):
        """List files inside a folder. Returns [{name, id}]."""
        results = self.service.files().list(
            q=f"'{parent_id}' in parents and trashed=false",
            fields="files(id, name, mimeType)",
            pageSize=200
        ).execute()
        files = results.get('files', [])
        if extension:
            files = [f for f in files if f['name'].endswith(extension)]
        return files

    def download_file(self, file_id):
        """Download a file and return its content as bytes."""
        from googleapiclient.http import MediaIoBaseDownload
        request = self.service.files().get_media(fileId=file_id)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buffer.seek(0)
        return buffer

    def find_folder(self, parent_id, folder_name):
        """Find a specific subfolder by name. Returns folder ID or None."""
        folders = self.list_folders(parent_id)
        return folders.get(folder_name)


def load_config():
    """Load config.json if it exists."""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return {}


# ---------------------------------------------------------------------------
#  Styling helpers
# ---------------------------------------------------------------------------

def style_header_row(ws, row, max_col, font=header_font, fill=header_fill):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = font
        cell.fill = fill
        cell.alignment = center
        cell.border = thin_border


def auto_width(ws, min_width=12, max_width=40):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        lengths = []
        for cell in col_cells:
            if cell.value:
                lengths.append(len(str(cell.value)))
        if lengths:
            best = min(max(max(lengths) + 2, min_width), max_width)
            ws.column_dimensions[col_letter].width = best


def apply_number_format(ws, row, col):
    ws.cell(row=row, column=col).number_format = '#,##0'


def _parse_xlsx_from_source(source):
    """Open an openpyxl workbook from a file path (str) or BytesIO buffer."""
    if isinstance(source, io.BytesIO):
        return openpyxl.load_workbook(source, read_only=True)
    return openpyxl.load_workbook(source, read_only=True)


def read_smartcard_data_local():
    """Read all SmartCard Excel files from local filesystem."""
    data = {}  # {lo_name: {day: count, 'raw': {day: [rows]}}}

    for day in DAY_FOLDERS:
        day_path = os.path.join(SC_DIR, day)
        if not os.path.isdir(day_path):
            continue
        for fname in sorted(os.listdir(day_path)):
            if not fname.endswith(".xlsx") or fname.startswith("."):
                continue
            key = fname.replace(".xlsx", "")
            lo_name = LO_NAME_MAP.get(key, key)

            if lo_name not in data:
                data[lo_name] = {"raw": {}}

            fpath = os.path.join(day_path, fname)
            wb = openpyxl.load_workbook(fpath, read_only=True)
            ws = wb.active
            rows = []
            count = 0
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue  # skip header
                rows.append(row)
                count += 1
            wb.close()

            data[lo_name][day] = count
            data[lo_name].setdefault("raw", {})[day] = rows

    return data


def read_smartcard_data_drive(drive, folder_id):
    """Read all SmartCard Excel files from Google Drive."""
    data = {}  # {lo_name: {day: count, 'raw': {day: [rows]}}}

    # Find the SmartCard subfolder
    sc_folder_id = drive.find_folder(folder_id, "SmartCard")
    if not sc_folder_id:
        print("   ⚠️  'SmartCard' folder not found in Drive. Skipping.")
        return data

    # List day sub-folders
    day_folders = drive.list_folders(sc_folder_id)
    print(f"   Found day folders: {list(day_folders.keys())}")

    for day in DAY_FOLDERS:
        day_folder_id = day_folders.get(day)
        if not day_folder_id:
            print(f"   ⚠️  '{day}' folder not found. Skipping.")
            continue

        files = drive.list_files(day_folder_id, extension=".xlsx")
        print(f"   📂 {day}: {len(files)} files")

        for f in sorted(files, key=lambda x: x['name']):
            fname = f['name']
            key = fname.replace(".xlsx", "")
            lo_name = LO_NAME_MAP.get(key, key)

            if lo_name not in data:
                data[lo_name] = {"raw": {}}

            print(f"      ⬇️  Downloading {fname}...")
            file_buffer = drive.download_file(f['id'])
            wb = _parse_xlsx_from_source(file_buffer)
            ws = wb.active
            rows = []
            count = 0
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue  # skip header
                rows.append(row)
                count += 1
            wb.close()

            data[lo_name][day] = count
            data[lo_name].setdefault("raw", {})[day] = rows

    return data


def read_smartcard_data(source="local", drive=None, folder_id=None):
    """Read SmartCard data from local or Drive source."""
    if source == "drive":
        return read_smartcard_data_drive(drive, folder_id)
    return read_smartcard_data_local()


def read_renewal_data_local():
    """Read Renewal 30 days Excel from local filesystem."""
    wb = openpyxl.load_workbook(RENEWAL_FILE, read_only=True)
    ws = wb.active
    headers = None
    rows = []
    district_counts = {}

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = list(row)
            continue
        rows.append(list(row))
        code = row[3]  # application_code_suffix
        if code and len(str(code)) > 6:
            loc_code = str(code)[4:7]
            dist_name = RENEWAL_LOC_MAP.get(loc_code, loc_code)
            district_counts[dist_name] = district_counts.get(dist_name, 0) + 1

    wb.close()
    sorted_districts = sorted(district_counts.items(), key=lambda x: -x[1])
    return headers, rows, sorted_districts


def read_renewal_data_drive(drive, folder_id):
    """Read Renewal 30 days Excel from Google Drive."""
    # Look for a file matching 'Renewal' in the root folder
    files = drive.list_files(folder_id, extension=".xlsx")
    renewal_file = None
    for f in files:
        if 'renewal' in f['name'].lower() or 'Renewal' in f['name']:
            renewal_file = f
            break

    if not renewal_file:
        print("   ⚠️  No Renewal file found in Drive root. Skipping.")
        return [], [], []

    print(f"   ⬇️  Downloading {renewal_file['name']}...")
    file_buffer = drive.download_file(renewal_file['id'])
    wb = _parse_xlsx_from_source(file_buffer)
    ws = wb.active
    headers = None
    rows = []
    district_counts = {}

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = list(row)
            continue
        rows.append(list(row))
        code = row[3]  # application_code_suffix
        if code and len(str(code)) > 6:
            loc_code = str(code)[4:7]
            dist_name = RENEWAL_LOC_MAP.get(loc_code, loc_code)
            district_counts[dist_name] = district_counts.get(dist_name, 0) + 1

    wb.close()
    sorted_districts = sorted(district_counts.items(), key=lambda x: -x[1])
    return headers, rows, sorted_districts


def read_renewal_data(source="local", drive=None, folder_id=None):
    """Read Renewal data from local or Drive source."""
    if source == "drive":
        return read_renewal_data_drive(drive, folder_id)
    return read_renewal_data_local()


def create_summary_sheet(wb, sc_data, rn_districts, rn_total):
    """Create the Summary Dashboard sheet."""
    ws = wb.active
    ws.title = "Summary Dashboard"
    ws.sheet_properties.tabColor = DARK_BLUE

    # Title
    ws.merge_cells("A1:G1")
    ws["A1"] = "📱 SMS Campaign — Consolidated Report"
    ws["A1"].font = Font(name="Calibri", bold=True, size=18, color=DARK_BLUE)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:G2")
    ws["A2"] = f"Report Generated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}"
    ws["A2"].font = Font(name="Calibri", size=11, color="888888")
    ws["A2"].alignment = center

    # --- Overall Summary ---
    row = 4
    ws.merge_cells(f"A{row}:G{row}")
    ws[f"A{row}"] = "OVERALL SUMMARY"
    ws[f"A{row}"].font = subtitle_font
    ws.row_dimensions[row].height = 28

    row = 5
    headers_list = ["Campaign", "Total SMS Sent", "Days", "Locations/Districts", "Date Range", "SMS Success Rate", "Status"]
    for c, h in enumerate(headers_list, 1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, len(headers_list))

    sc_grand = sum(sum(v for k, v in lo.items() if k in DAY_FOLDERS) for lo in sc_data.values())

    row = 6
    sc_row = ["SmartCard", sc_grand, 3, f"{len(sc_data)} LOs", "28–30 Apr 2026", "100%", "✅ Completed"]
    for c, v in enumerate(sc_row, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = normal_font
        cell.alignment = center
        cell.border = thin_border
        if c == 2:
            cell.number_format = '#,##0'

    row = 7
    rn_row = ["Renewal 30 Days", rn_total, 1, f"{len(rn_districts)} Districts", "30 Apr 2026", "100%", "✅ Completed"]
    for c, v in enumerate(rn_row, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = normal_font
        cell.alignment = center
        cell.border = thin_border
        if c == 2:
            cell.number_format = '#,##0'
        if row % 2 == 1:
            cell.fill = alt_fill

    row = 8
    total_row = ["GRAND TOTAL", sc_grand + rn_total, "—", "—", "28–30 Apr 2026", "100%", "✅ All Done"]
    for c, v in enumerate(total_row, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = total_font
        cell.fill = total_fill
        cell.alignment = center
        cell.border = thin_border
        if c == 2:
            cell.number_format = '#,##0'

    # --- SmartCard Summary by LO ---
    row = 10
    ws.merge_cells(f"A{row}:F{row}")
    ws[f"A{row}"] = "💳 SMARTCARD SMS — BY LABOUR OFFICE"
    ws[f"A{row}"].font = subtitle_font
    ws.row_dimensions[row].height = 28

    row = 11
    sc_headers = ["#", "Labour Office", "28th April", "29th April", "30th April", "Total"]
    for c, h in enumerate(sc_headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, len(sc_headers))

    row = 12
    grand_totals = {d: 0 for d in DAY_FOLDERS}
    grand_total_all = 0
    for idx, (lo_name, lo_data) in enumerate(sorted(sc_data.items()), 1):
        ws.cell(row=row, column=1, value=idx).font = normal_font
        ws.cell(row=row, column=2, value=lo_name).font = bold_font
        lo_total = 0
        for c, day in enumerate(DAY_FOLDERS, 3):
            cnt = lo_data.get(day, 0)
            cell = ws.cell(row=row, column=c, value=cnt)
            cell.number_format = '#,##0'
            cell.font = number_font
            cell.alignment = right_align
            grand_totals[day] += cnt
            lo_total += cnt
        cell = ws.cell(row=row, column=6, value=lo_total)
        cell.number_format = '#,##0'
        cell.font = total_font
        cell.alignment = right_align
        grand_total_all += lo_total

        for c in range(1, 7):
            ws.cell(row=row, column=c).border = thin_border
            if c == 1:
                ws.cell(row=row, column=c).alignment = center
            if idx % 2 == 0:
                ws.cell(row=row, column=c).fill = alt_fill
        row += 1

    # Grand total row
    ws.cell(row=row, column=1, value="").font = normal_font
    ws.cell(row=row, column=2, value="GRAND TOTAL").font = total_font
    for c, day in enumerate(DAY_FOLDERS, 3):
        cell = ws.cell(row=row, column=c, value=grand_totals[day])
        cell.number_format = '#,##0'
        cell.font = total_font
        cell.alignment = right_align
    cell = ws.cell(row=row, column=6, value=grand_total_all)
    cell.number_format = '#,##0'
    cell.font = total_font
    cell.alignment = right_align
    for c in range(1, 7):
        ws.cell(row=row, column=c).fill = total_fill
        ws.cell(row=row, column=c).border = thin_border

    # --- Renewal Summary by District ---
    row += 2
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "🔄 RENEWAL 30 DAYS SMS — BY DISTRICT"
    ws[f"A{row}"].font = subtitle_font
    ws.row_dimensions[row].height = 28

    row += 1
    rn_headers = ["#", "District", "SMS Count", "% Share", "Status"]
    for c, h in enumerate(rn_headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, len(rn_headers))

    row += 1
    rn_total_count = sum(c for _, c in rn_districts)
    for idx, (dist, cnt) in enumerate(rn_districts, 1):
        ws.cell(row=row, column=1, value=idx).font = normal_font
        ws.cell(row=row, column=2, value=dist).font = normal_font
        cell = ws.cell(row=row, column=3, value=cnt)
        cell.number_format = '#,##0'
        cell.font = number_font
        pct = round(cnt / rn_total_count * 100, 1) if rn_total_count else 0
        ws.cell(row=row, column=4, value=f"{pct}%").font = normal_font
        ws.cell(row=row, column=5, value="✅ Sent").font = normal_font
        for c in range(1, 6):
            ws.cell(row=row, column=c).border = thin_border
            ws.cell(row=row, column=c).alignment = center if c in (1, 4, 5) else left_align if c == 2 else right_align
            if idx % 2 == 0:
                ws.cell(row=row, column=c).fill = alt_fill
        row += 1

    # Total
    ws.cell(row=row, column=2, value="GRAND TOTAL").font = total_font
    cell = ws.cell(row=row, column=3, value=rn_total_count)
    cell.number_format = '#,##0'
    cell.font = total_font
    ws.cell(row=row, column=4, value="100%").font = total_font
    ws.cell(row=row, column=5, value="✅").font = total_font
    for c in range(1, 6):
        ws.cell(row=row, column=c).fill = total_fill
        ws.cell(row=row, column=c).border = thin_border

    auto_width(ws)
    ws.freeze_panes = "A6"


def create_smartcard_lo_sheets(wb, sc_data):
    """Create a detailed sheet per LO with day-wise raw data."""
    for lo_name, lo_data in sorted(sc_data.items()):
        # Sheet name max 31 chars
        sheet_name = lo_name[:31].replace("–", "-")
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_properties.tabColor = "6366F1"

        # Title
        ws.merge_cells("A1:D1")
        ws["A1"] = f"💳 {lo_name} — SmartCard SMS Detail"
        ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=DARK_BLUE)
        ws.row_dimensions[1].height = 32

        row = 3
        for day in DAY_FOLDERS:
            raw_rows = lo_data.get("raw", {}).get(day, [])
            count = lo_data.get(day, 0)

            # Day header
            ws.merge_cells(f"A{row}:D{row}")
            ws[f"A{row}"] = f"📅 {day} — {count:,} SMS"
            ws[f"A{row}"].font = Font(name="Calibri", bold=True, size=12, color=WHITE)
            ws[f"A{row}"].fill = PatternFill(start_color=MED_BLUE, end_color=MED_BLUE, fill_type="solid")
            ws[f"A{row}"].alignment = left_align
            ws.row_dimensions[row].height = 26
            row += 1

            # Column headers
            col_headers = ["#", "Mobile No", "Location", "SMS Content / Address"]
            for c, h in enumerate(col_headers, 1):
                ws.cell(row=row, column=c, value=h)
            style_header_row(ws, row, len(col_headers), sub_header_font, sub_header_fill)
            row += 1

            # Data rows
            for idx, r in enumerate(raw_rows, 1):
                ws.cell(row=row, column=1, value=idx).font = normal_font
                ws.cell(row=row, column=1).alignment = center
                mobile = r[0] if r[0] else ""
                ws.cell(row=row, column=2, value=str(mobile)).font = normal_font
                ws.cell(row=row, column=3, value=r[1] if len(r) > 1 else "").font = normal_font
                ws.cell(row=row, column=4, value=r[2] if len(r) > 2 else "").font = normal_font
                for c in range(1, 5):
                    ws.cell(row=row, column=c).border = thin_border
                    if idx % 2 == 0:
                        ws.cell(row=row, column=c).fill = alt_fill
                row += 1

            row += 1  # gap between days

        auto_width(ws, min_width=14)
        ws.freeze_panes = "A4"


def create_renewal_raw_sheet(wb, headers, rows):
    """Create Renewal 30 Days raw data sheet."""
    ws = wb.create_sheet(title="Renewal 30D Raw Data")
    ws.sheet_properties.tabColor = "22D3EE"

    ws.merge_cells("A1:K1")
    ws["A1"] = "🔄 Renewal 30 Days — Raw SMS Data (30th April 2026)"
    ws["A1"].font = Font(name="Calibri", bold=True, size=14, color=DARK_BLUE)
    ws.row_dimensions[1].height = 32

    # Headers
    row = 3
    for c, h in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=h)
    style_header_row(ws, row, len(headers))

    # Data (limit to first 50000 to avoid huge files)
    max_rows = min(len(rows), 50000)
    for idx, r in enumerate(rows[:max_rows]):
        data_row = row + 1 + idx
        for c, v in enumerate(r, 1):
            cell = ws.cell(row=data_row, column=c, value=v)
            cell.font = normal_font
            cell.border = thin_border
            if idx % 2 == 1:
                cell.fill = alt_fill

    # Add count note
    note_row = row + max_rows + 2
    ws.cell(row=note_row, column=1, value=f"Total records: {len(rows):,}").font = bold_font

    auto_width(ws, min_width=12, max_width=30)
    ws.freeze_panes = "A4"


def parse_args():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description="SMS Campaign — Consolidated Excel Report Generator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_report.py                                    # local mode
  python generate_report.py --source drive                     # drive mode (uses config.json)
  python generate_report.py --source drive --folder-id ABC123  # drive mode with folder ID
        """
    )
    parser.add_argument(
        '--source', choices=['local', 'drive'], default=None,
        help='Data source: "local" (filesystem) or "drive" (Google Drive). '
             'Defaults to config.json value, or "local" if not set.'
    )
    parser.add_argument(
        '--folder-id', default=None,
        help='Google Drive folder ID containing SmartCard/ and Renewal files.'
    )
    parser.add_argument(
        '--key', default=None,
        help='Path to Google Service Account JSON key file.'
    )
    return parser.parse_args()


def main():
    args = parse_args()
    config = load_config()

    # Resolve source — CLI overrides config, default is "local"
    source = args.source or config.get('source', 'local')
    folder_id = args.folder_id or config.get('drive_folder_id')
    key_path = args.key or config.get('service_account_key', 'service_account.json')

    print("📊 Generating SMS Consolidated Report...")
    print("=" * 50)
    print(f"   Source: {source.upper()}")

    drive = None
    if source == 'drive':
        if not folder_id or folder_id == 'YOUR_FOLDER_ID_HERE':
            print("❌ Error: Drive folder ID not set.")
            print("   Set it in config.json or pass --folder-id <ID>")
            return
        if not os.path.exists(key_path):
            print(f"❌ Error: Service account key not found at: {key_path}")
            print("   Download it from Google Cloud Console and place it in this directory.")
            print("   Or pass --key <path_to_key.json>")
            return
        print(f"   🔑 Using service account: {key_path}")
        print(f"   📁 Drive folder ID: {folder_id}")
        drive = GoogleDriveReader(key_path)

    # Read data
    print("\n📂 Reading SmartCard data...")
    sc_data = read_smartcard_data(source=source, drive=drive, folder_id=folder_id)
    print(f"   Found {len(sc_data)} Labour Offices")

    print("📂 Reading Renewal 30 Days data...")
    rn_headers, rn_rows, rn_districts = read_renewal_data(source=source, drive=drive, folder_id=folder_id)
    rn_total = len(rn_rows)
    print(f"   Found {rn_total:,} records across {len(rn_districts)} districts")

    # Create workbook
    wb = openpyxl.Workbook()

    print("\n📝 Creating Summary Dashboard...")
    create_summary_sheet(wb, sc_data, rn_districts, rn_total)

    print("📝 Creating SmartCard LO detail sheets...")
    create_smartcard_lo_sheets(wb, sc_data)

    print("📝 Creating Renewal raw data sheet...")
    create_renewal_raw_sheet(wb, rn_headers, rn_rows)

    # Save
    print(f"\n💾 Saving to: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)
    print(f"✅ Report generated successfully!")
    print(f"   File: {OUTPUT_FILE}")
    print(f"   Size: {os.path.getsize(OUTPUT_FILE) / (1024*1024):.1f} MB")


if __name__ == "__main__":
    main()
