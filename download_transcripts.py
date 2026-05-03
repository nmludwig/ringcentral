#!/usr/bin/env python3
"""
Ludwig's RingCentral ACE Transcript Downloader
======================================
Downloads all RingSense call transcripts from a customer account.

Saves to the same folder as this script:
    transcripts_[customer]_[date].xlsx   -- Excel spreadsheet
    transcripts_[customer]_[date].pdf    -- Formatted PDF

Run:
    python3 download_transcripts.py

The script will ask for everything it needs.
Credentials are never saved to disk.
"""

import sys
import os
import re
import json
import time
import subprocess
from datetime import datetime
from datetime import timezone
from getpass import getpass
from pathlib import Path


# ── Auto-install all dependencies ────────────────────────────────────────────
def install(import_name, pip_name=None):
    """Install a package if not already present."""
    try:
        __import__(import_name)
    except ImportError:
        pkg = pip_name or import_name
        print("  Installing " + pkg + "...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", pkg, "-q"],
            check=True
        )

print("")
print("Checking dependencies...")
install("requests")
install("openpyxl")
install("reportlab")
install("PIL", "pillow")
print("")

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Terminal colors ───────────────────────────────────────────────────────────
G    = "\033[92m"   # green
Y    = "\033[93m"   # yellow
R    = "\033[91m"   # red
B    = "\033[94m"   # blue
BOLD = "\033[1m"
DIM  = "\033[2m"
W    = "\033[0m"    # reset

def ok(msg):     print("  " + G + "OK" + W + "  " + str(msg))
def info(msg):   print("  " + B + "->" + W + "  " + str(msg))
def warn(msg):   print("  " + Y + "!!" + W + "  " + str(msg))
def err(msg):    print("  " + R + "XX" + W + "  " + str(msg))
def header(msg): print("\n" + BOLD + str(msg) + W)
def rule():      print(DIM + "-" * 56 + W)


# ── Banner ────────────────────────────────────────────────────────────────────
def print_banner():
    print("")
    print(BOLD + "+======================================================+" + W)
    print(BOLD + "|  RingCentral ACE Transcript Downloader               |" + W)
    print(BOLD + "|  Downloads all call transcripts to Excel + PDF       |" + W)
    print(BOLD + "+======================================================+" + W)
    print("")


# ── Prompts ───────────────────────────────────────────────────────────────────
def ask(prompt, default=None, secret=False):
    suffix = (" [" + default + "]: ") if default else ": "
    display = "  " + prompt + suffix
    val = getpass(display) if secret else input(display).strip()
    if not val and default:
        return default
    return val.strip()


def ask_choice(prompt, choices):
    print("\n  " + prompt)
    for i, c in enumerate(choices, 1):
        print("    " + B + str(i) + W + ". " + c)
    while True:
        val = input("\n  Enter number (1-" + str(len(choices)) + "): ").strip()
        if val.isdigit() and 1 <= int(val) <= len(choices):
            return int(val) - 1
        err("Please enter a number between 1 and " + str(len(choices)))


def ask_date(prompt, default):
    while True:
        val = ask(prompt, default)
        if re.match(r"^\d{4}-\d{2}-\d{2}$", val):
            return val
        err("Please enter date as YYYY-MM-DD  e.g. 2026-01-01")


def confirm(prompt):
    return input("  " + prompt + " (y/n): ").strip().lower() in ("y", "yes")


# ── Step 1: Collect credentials ───────────────────────────────────────────────
def collect_credentials():
    header("STEP 1 -- Customer & Credentials")
    rule()
    print("""
  You need three things from developers.ringcentral.com:

    Client ID      -- from your app Credentials tab
    Client Secret  -- from your app Credentials tab
    JWT Token      -- from your app Auth tab (Create JWT)

  These are only used during this session.
  They are never saved to disk.
""")
    return {
        "customer_name": ask("Customer company name"),
        "client_id":     ask("Client ID"),
        "client_secret": ask("Client Secret", secret=True),
        "jwt_token":     ask("JWT Token",     secret=True),
    }


# ── Step 2: Authenticate ──────────────────────────────────────────────────────
def authenticate(creds):
    header("STEP 2 -- Authenticating with RingCentral")
    rule()
    info("Connecting...")

    try:
        resp = requests.post(
            "https://platform.ringcentral.com/restapi/oauth/token",
            auth=(creds["client_id"], creds["client_secret"]),
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                "assertion":  creds["jwt_token"],
            },
            timeout=15,
        )
        resp.raise_for_status()

    except requests.exceptions.HTTPError as e:
        code = e.response.status_code if e.response else "?"
        if code == 401:
            err("Authentication failed")
            err("Check your Client ID, Client Secret, and JWT Token")
            err("Make sure you regenerated the JWT after adding the RingSense scope")
        elif code == 400:
            err("JWT Token may be expired")
            err("Go to developers.ringcentral.com and create a new JWT Token")
        else:
            err("HTTP error " + str(code) + ": " + str(e))
        sys.exit(1)

    except requests.exceptions.ConnectionError:
        err("Cannot connect -- check your internet connection")
        sys.exit(1)

    data  = resp.json()
    token = data["access_token"]
    scope = data.get("scope", "")

    ok("Authenticated successfully")
    info("Scopes: " + scope)

    if "RingSense" not in scope and "ringsense" not in scope.lower():
        warn("RingSense scope not found in token")
        warn("Transcripts may not be available")
        warn("Ensure the RingSense scope was approved by RingCentral")
        warn("and that you regenerated the JWT token after it was approved")
        print("")
        if not confirm("Continue anyway?"):
            sys.exit(0)

    return token


# ── Step 3: Find account ID ───────────────────────────────────────────────────
def get_account_id(token):
    info("Looking up Account ID...")
    resp = requests.get(
        "https://platform.ringcentral.com/restapi/v1.0/account/~/call-log",
        headers={"Authorization": "Bearer " + token},
        params={"perPage": 1},
        timeout=15,
    )
    resp.raise_for_status()
    match = re.search(r"account/(\d+)", resp.json().get("uri", ""))
    if match:
        account_id = match.group(1)
        ok("Account ID: " + account_id)
        return account_id
    err("Could not determine Account ID")
    sys.exit(1)


# ── Step 4: Choose date range ─────────────────────────────────────────────────
def choose_date_range():
    header("STEP 3 -- Date Range")
    rule()
    print("\n  For a POC, 30 days is a good starting point.\n")

    choice = ask_choice("How far back do you want to pull transcripts?", [
        "Last 7 days",
        "Last 30 days",
        "Last 90 days",
        "Custom range -- I will enter dates",
    ])

    today = datetime.now()

    if choice == 0:
        date_from = today.replace(day=max(1, today.day - 7)).strftime("%Y-%m-%d")
    elif choice == 1:
        m = today.month - 1 or 12
        y = today.year if today.month > 1 else today.year - 1
        date_from = today.replace(year=y, month=m, day=1).strftime("%Y-%m-%d")
    elif choice == 2:
        d = today
        for _ in range(3):
            m = d.month - 1 or 12
            y = d.year if d.month > 1 else d.year - 1
            d = d.replace(year=y, month=m, day=1)
        date_from = d.strftime("%Y-%m-%d")
    else:
        date_from = ask_date("Start date (YYYY-MM-DD)", "2026-01-01")

    date_to = today.strftime("%Y-%m-%d")
    ok("Date range: " + date_from + " to " + date_to)
    return date_from + "T00:00:00Z", date_to + "T23:59:59Z"


# ── Step 5: Download call log ─────────────────────────────────────────────────
def download_call_logs(token, account_id, date_from, date_to):
    header("STEP 4 -- Downloading Call Log")
    rule()
    info("Fetching recorded calls from " + date_from[:10] + " to " + date_to[:10] + "...")

    records = []
    page    = 1

    while True:
        print("    Fetching page " + str(page) + "...", end="\r")
        resp = requests.get(
            "https://platform.ringcentral.com/restapi/v1.0/account/" + account_id + "/call-log",
            headers={"Authorization": "Bearer " + token},
            params={
                "view":          "Detailed",
                "dateFrom":      date_from,
                "dateTo":        date_to,
                "type":          "Voice",
                "withRecording": "true",
                "perPage":       250,
                "page":          page,
            },
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        records.extend(data.get("records", []))
        if not data.get("navigation", {}).get("nextPage"):
            break
        page += 1
        time.sleep(0.25)

    print("")
    ok("Found " + str(len(records)) + " recorded calls")
    return records


# ── Step 6: Fetch RingSense transcripts ───────────────────────────────────────
def fetch_transcripts(token, call_logs):
    header("STEP 5 -- Fetching RingSense Transcripts")
    rule()
    total = len(call_logs)
    info("Fetching transcripts for " + str(total) + " calls...")
    info("Takes about 1.5 seconds per call -- please be patient")
    print("")

    records          = []
    with_transcripts = 0

    for i, call in enumerate(call_logs):
        recording_id = call.get("recording", {}).get("id")
        if not recording_id:
            continue

        # Progress bar
        pct    = int((i / max(total, 1)) * 100)
        filled = int(40 * pct / 100)
        bar    = G + ("#" * filled) + DIM + ("." * (40 - filled)) + W
        print("    [" + bar + "] " + str(pct).rjust(3) + "%  " + str(i + 1) + "/" + str(total), end="\r")

        # Fetch insights
        insights = None
        url = (
            "https://platform.ringcentral.com/ai/ringsense/v1/public"
            "/accounts/~/domains/pbx/records/" + recording_id + "/insights"
        )

        while True:
            try:
                resp = requests.get(
                    url,
                    headers={"Authorization": "Bearer " + token},
                    timeout=30,
                )
                if resp.status_code == 429:
                    print("")
                    warn("Rate limit reached -- waiting 65 seconds...")
                    time.sleep(65)
                    continue
                if resp.status_code in (404, 403):
                    break
                if resp.status_code == 200:
                    insights = resp.json()
                    break
                break
            except Exception:
                break

        if insights:
            with_transcripts += 1

        # Real speaker names from speakerInfo
        speaker_map = {}
        if insights:
            for sp in insights.get("speakerInfo", []):
                sid  = sp.get("speakerId", "")
                name = sp.get("name", "") or sp.get("phoneNumber", sid)
                if sid and name:
                    speaker_map[sid] = name

        # Build readable transcript text
        utterances = (insights or {}).get("insights", {}).get("Transcript", [])
        lines      = []
        for u in utterances:
            sid   = u.get("speakerId", "?")
            name  = speaker_map.get(sid, sid)
            txt   = u.get("text", "").strip()
            start = u.get("start", 0)
            mm    = str(int(start // 60)).zfill(2)
            ss    = str(int(start % 60)).zfill(2)
            lines.append("[" + mm + ":" + ss + "] " + name + ": " + txt)

        transcript_text = "\n".join(lines)

        # AI Summary and Sentiment
        summary_list   = (insights or {}).get("insights", {}).get("Summary",   [])
        sentiment_list = (insights or {}).get("insights", {}).get("Sentiment", [])
        summary        = summary_list[0].get("value",   "") if summary_list   else ""
        sentiment      = sentiment_list[0].get("value", "") if sentiment_list else ""

        rec = call.get("recording", {})
        records.append({
            "call_id":         call.get("id", ""),
            "start_time":      call.get("startTime", ""),
            "duration_sec":    call.get("duration", 0),
            "direction":       call.get("direction", ""),
            "from_number":     call.get("from", {}).get("phoneNumber", ""),
            "from_name":       call.get("from", {}).get("name", ""),
            "to_number":       call.get("to", {}).get("phoneNumber", ""),
            "to_name":         call.get("to", {}).get("name", ""),
            "recording_id":    rec.get("id", ""),
            "has_transcript":  insights is not None,
            "sentiment":       sentiment,
            "summary":         summary,
            "transcript":      transcript_text,
        })

        time.sleep(1.5)

    print("")
    ok(str(with_transcripts) + " of " + str(total) + " calls have transcripts")

    if with_transcripts == 0:
        warn("No transcripts found. Common reasons:")
        warn("  * RingSense scope not yet approved by RingCentral")
        warn("  * JWT token was not regenerated after scope approval")
        warn("  * Extensions do not have RingSense licenses assigned")

    return records, with_transcripts


# ── Build Excel spreadsheet ───────────────────────────────────────────────────
def build_excel(records, customer_name, date_from, date_to, out_path):
    wb = openpyxl.Workbook()

    # ── Colors ────────────────────────────────────────────────────────────────
    RC_ORANGE  = "FF6A00"
    DARK       = "1A1A1A"
    WHITE      = "FFFFFF"
    LIGHT_GREY = "F5F5F5"
    BLUE_LIGHT = "E6F1FB"
    GREEN_LIGHT= "E8F8EF"
    RED_LIGHT  = "FEECEC"
    BORDER_CLR = "CCCCCC"

    thin = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Sheet 1: Summary ──────────────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Summary"

    # Title row
    ws_sum.merge_cells("A1:H1")
    title_cell = ws_sum["A1"]
    title_cell.value = "RingCentral ACE Transcript Report -- " + customer_name
    title_cell.font       = Font(name="Calibri", size=16, bold=True, color=WHITE)
    title_cell.fill       = PatternFill("solid", fgColor=RC_ORANGE)
    title_cell.alignment  = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    # Subtitle
    ws_sum.merge_cells("A2:H2")
    sub_cell = ws_sum["A2"]
    sub_cell.value = (date_from[:10] + " to " + date_to[:10] +
                      "   |   Generated " + datetime.now().strftime("%B %d, %Y") +
                      "   |   RingCentral AI Conversation Expert   |   Confidential")
    sub_cell.font      = Font(name="Calibri", size=10, color="555555")
    sub_cell.fill      = PatternFill("solid", fgColor=LIGHT_GREY)
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    # Stats row
    with_trans = len([r for r in records if r["has_transcript"]])
    total_sec  = sum(r["duration_sec"] for r in records)
    hrs, rem   = divmod(total_sec, 3600)
    mins       = rem // 60

    stats = [
        ("Total Calls",         str(len(records))),
        ("With Transcripts",    str(with_trans)),
        ("Without Transcripts", str(len(records) - with_trans)),
        ("Total Talk Time",     str(hrs) + "h " + str(mins) + "m"),
        ("Date From",           date_from[:10]),
        ("Date To",             date_to[:10]),
    ]

    ws_sum.append([])  # row 3 blank
    for row_idx, (label, value) in enumerate(stats, start=4):
        ws_sum.cell(row=row_idx, column=1).value = label
        ws_sum.cell(row=row_idx, column=1).font  = Font(name="Calibri", size=11, bold=True, color=DARK)
        ws_sum.cell(row=row_idx, column=1).fill  = PatternFill("solid", fgColor=LIGHT_GREY)
        ws_sum.cell(row=row_idx, column=2).value = value
        ws_sum.cell(row=row_idx, column=2).font  = Font(name="Calibri", size=11, color=DARK)

    ws_sum.column_dimensions["A"].width = 24
    ws_sum.column_dimensions["B"].width = 20

    # ── Sheet 2: All Calls ────────────────────────────────────────────────────
    ws_calls = wb.create_sheet("All Calls")

    headers = [
        "Date", "Time", "Direction", "Duration",
        "From Name", "From Number", "To Name", "To Number",
        "Has Transcript", "Sentiment", "AI Summary"
    ]
    col_widths = [14, 10, 12, 10, 22, 18, 22, 18, 14, 12, 60]

    # Header row
    for col, (hdr, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws_calls.cell(row=1, column=col)
        cell.value     = hdr
        cell.font      = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=RC_ORANGE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
        ws_calls.column_dimensions[get_column_letter(col)].width = width

    ws_calls.row_dimensions[1].height = 24

    # Data rows
    for row_idx, rec in enumerate(records, start=2):
        # Parse date/time
        try:
            dt       = datetime.fromisoformat(rec["start_time"].replace("Z", "+00:00"))
            date_str = dt.strftime("%Y-%m-%d")
            time_str = dt.strftime("%I:%M %p")
        except Exception:
            date_str = rec["start_time"][:10]
            time_str = ""

        dur_m, dur_s = divmod(int(rec["duration_sec"]), 60)
        dur_str      = str(dur_m) + "m " + str(dur_s) + "s"

        row_data = [
            date_str,
            time_str,
            rec["direction"],
            dur_str,
            rec["from_name"],
            rec["from_number"],
            rec["to_name"],
            rec["to_number"],
            "Yes" if rec["has_transcript"] else "No",
            rec["sentiment"],
            rec["summary"],
        ]

        # Row background based on sentiment
        sl = rec["sentiment"].lower()
        if "positive"  in sl: row_bg = GREEN_LIGHT
        elif "negative" in sl: row_bg = RED_LIGHT
        else:                  row_bg = WHITE if row_idx % 2 == 0 else LIGHT_GREY

        for col, val in enumerate(row_data, start=1):
            cell           = ws_calls.cell(row=row_idx, column=col)
            cell.value     = val
            cell.font      = Font(name="Calibri", size=10)
            cell.fill      = PatternFill("solid", fgColor=row_bg)
            cell.alignment = Alignment(vertical="top", wrap_text=(col == 11))
            cell.border    = border

        ws_calls.row_dimensions[row_idx].height = 30 if rec["summary"] else 18

    ws_calls.freeze_panes = "A2"

    # ── Sheet 3: Full Transcripts ─────────────────────────────────────────────
    ws_trans = wb.create_sheet("Transcripts")
    ws_trans.column_dimensions["A"].width = 18
    ws_trans.column_dimensions["B"].width = 90

    trans_records = [r for r in records if r["transcript"]]
    current_row   = 1

    for idx, rec in enumerate(trans_records, start=1):
        # Call header
        try:
            dt       = datetime.fromisoformat(rec["start_time"].replace("Z", "+00:00"))
            date_str = dt.strftime("%B %d, %Y")
            time_str = dt.strftime("%I:%M %p")
        except Exception:
            date_str = rec["start_time"][:10]
            time_str = ""

        from_name = rec["from_name"] or rec["from_number"] or "Unknown"
        to_name   = rec["to_name"]   or rec["to_number"]   or "Unknown"
        dur_m, dur_s = divmod(int(rec["duration_sec"]), 60)

        # Call title cell
        ws_trans.merge_cells(
            start_row=current_row, start_column=1,
            end_row=current_row,   end_column=2
        )
        title_val = ("CALL " + str(idx) + " OF " + str(len(trans_records)) + "   |   " +
                     rec["direction"] + "   |   " + date_str + "   " + time_str + "   |   " +
                     str(dur_m) + "m " + str(dur_s) + "s" + "   |   From: " + from_name +
                     "   To: " + to_name)
        cell = ws_trans.cell(row=current_row, column=1)
        cell.value     = title_val
        cell.font      = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=RC_ORANGE)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws_trans.row_dimensions[current_row].height = 22
        current_row += 1

        # Sentiment + summary
        sl = rec["sentiment"].lower()
        if "positive"  in sl: sent_bg = GREEN_LIGHT
        elif "negative" in sl: sent_bg = RED_LIGHT
        else:                  sent_bg = LIGHT_GREY

        sent_cell = ws_trans.cell(row=current_row, column=1)
        sent_cell.value     = "Sentiment"
        sent_cell.font      = Font(name="Calibri", size=9, bold=True)
        sent_cell.fill      = PatternFill("solid", fgColor=sent_bg)
        sent_cell.alignment = Alignment(vertical="center")

        sent_val = ws_trans.cell(row=current_row, column=2)
        sent_val.value     = rec["sentiment"] or "Neutral"
        sent_val.font      = Font(name="Calibri", size=9)
        sent_val.fill      = PatternFill("solid", fgColor=sent_bg)
        ws_trans.row_dimensions[current_row].height = 16
        current_row += 1

        if rec["summary"]:
            sum_lbl = ws_trans.cell(row=current_row, column=1)
            sum_lbl.value     = "AI Summary"
            sum_lbl.font      = Font(name="Calibri", size=9, bold=True)
            sum_lbl.fill      = PatternFill("solid", fgColor="FFF4EC")
            sum_lbl.alignment = Alignment(vertical="top")

            sum_val = ws_trans.cell(row=current_row, column=2)
            sum_val.value     = rec["summary"]
            sum_val.font      = Font(name="Calibri", size=9)
            sum_val.fill      = PatternFill("solid", fgColor="FFF4EC")
            sum_val.alignment = Alignment(vertical="top", wrap_text=True)
            # Estimate row height based on text length
            ws_trans.row_dimensions[current_row].height = max(30, min(90, len(rec["summary"]) // 4))
            current_row += 1

        # Transcript lines
        for line in rec["transcript"].split("\n"):
            if not line.strip():
                continue
            # Parse [MM:SS] Speaker: text
            m = re.match(r"^(\[\d+:\d+\])\s+(.+?):\s+(.+)$", line)
            if m:
                ts_str  = m.group(1)
                speaker = m.group(2)
                text    = m.group(3)

                ts_cell = ws_trans.cell(row=current_row, column=1)
                ts_cell.value     = ts_str + "  " + speaker
                ts_cell.font      = Font(name="Calibri", size=9, bold=True, color="185FA5")
                ts_cell.alignment = Alignment(vertical="top")

                txt_cell = ws_trans.cell(row=current_row, column=2)
                txt_cell.value     = text
                txt_cell.font      = Font(name="Calibri", size=9)
                txt_cell.alignment = Alignment(vertical="top", wrap_text=True)
                ws_trans.row_dimensions[current_row].height = max(15, min(60, len(text) // 8))
            else:
                merged_cell = ws_trans.cell(row=current_row, column=1)
                merged_cell.value = line
                merged_cell.font  = Font(name="Calibri", size=9, italic=True, color="888888")

            current_row += 1

        # Spacer between calls
        current_row += 1

    wb.save(str(out_path))


# ── Build PDF ─────────────────────────────────────────────────────────────────
def build_pdf(records, customer_name, date_from, date_to, out_path):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    HRFlowable, Table, TableStyle, KeepTogether)
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib.colors import HexColor, white
    from reportlab.lib.enums import TA_CENTER
    from reportlab.pdfgen import canvas as pdfcanvas

    RC_ORANGE = HexColor("#FF6A00")
    DARK      = HexColor("#1A1A1A")
    MUTED     = HexColor("#7A7A7A")
    MID       = HexColor("#4A4A4A")
    RULE      = HexColor("#E8E8E8")
    BG_GREY   = HexColor("#F5F5F5")
    BG_ORANGE = HexColor("#FFF4EC")
    BG_BLUE   = HexColor("#EAF3FB")
    BG_GREEN  = HexColor("#E8F8EF")
    BG_RED    = HexColor("#FEECEC")
    TX_BLUE   = HexColor("#185FA5")
    TX_GREEN  = HexColor("#1A7A45")
    TX_RED    = HexColor("#C0392B")
    TX_GREY   = HexColor("#555555")
    PURPLE    = HexColor("#6B3FA0")

    def ps(name, **kw):
        s = ParagraphStyle(name)
        for k, v in kw.items():
            setattr(s, k, v)
        return s

    call_num_s   = ps("CN",  fontSize=8,   fontName="Helvetica-Bold", textColor=RC_ORANGE, spaceAfter=2)
    call_title_s = ps("CT",  fontSize=13,  fontName="Helvetica-Bold", textColor=DARK,      spaceAfter=4,  leading=16)
    meta_s       = ps("ME",  fontSize=8.5, fontName="Helvetica",      textColor=MUTED,     spaceAfter=3)
    sum_lbl_s    = ps("SL",  fontSize=7.5, fontName="Helvetica-Bold", textColor=MUTED,     spaceBefore=10, spaceAfter=3)
    sum_txt_s    = ps("ST",  fontSize=9.5, fontName="Helvetica",      textColor=MID,       leading=14,    spaceAfter=4,  leftIndent=10, rightIndent=10)
    tr_lbl_s     = ps("TL",  fontSize=7.5, fontName="Helvetica-Bold", textColor=MUTED,     spaceBefore=10, spaceAfter=4)
    sp_a_s       = ps("SA",  fontSize=8.5, fontName="Helvetica-Bold", textColor=TX_BLUE,   spaceAfter=1)
    sp_b_s       = ps("SB",  fontSize=8.5, fontName="Helvetica-Bold", textColor=PURPLE,    spaceAfter=1)
    utt_s        = ps("UT",  fontSize=9.5, fontName="Helvetica",      textColor=DARK,      leading=14,    spaceAfter=7,  leftIndent=14)
    no_tr_s      = ps("NT",  fontSize=9,   fontName="Helvetica-Oblique", textColor=MUTED,  spaceAfter=4)
    cover_big_s  = ps("CB",  fontSize=26,  fontName="Helvetica-Bold", textColor=DARK,      alignment=TA_CENTER, spaceAfter=6)
    cover_sub_s  = ps("CS",  fontSize=13,  fontName="Helvetica",      textColor=MUTED,     alignment=TA_CENTER, spaceAfter=20)
    stat_s       = ps("SS",  fontSize=18,  fontName="Helvetica-Bold", textColor=DARK,      alignment=TA_CENTER, leading=22)
    stat_or_s    = ps("SO",  fontSize=18,  fontName="Helvetica-Bold", textColor=RC_ORANGE, alignment=TA_CENTER, leading=22)
    gen_s        = ps("GS",  fontSize=7.5, fontName="Helvetica",      textColor=MUTED,     alignment=TA_CENTER)
    hdr_txt_s    = ps("HT",  fontSize=22,  fontName="Helvetica-Bold", textColor=white,     alignment=TA_CENTER)

    class NumCanvas(pdfcanvas.Canvas):
        def __init__(self, *args, **kwargs):
            pdfcanvas.Canvas.__init__(self, *args, **kwargs)
            self._pages = []
        def showPage(self):
            self._pages.append(dict(self.__dict__))
            self._startPage()
        def save(self):
            n = len(self._pages)
            for state in self._pages:
                self.__dict__.update(state)
                self.setFont("Helvetica", 7.5)
                self.setFillColor(MUTED)
                self.drawRightString(
                    letter[0] - 0.5 * inch, 0.38 * inch,
                    "Page " + str(self._pageNumber) + " of " + str(n)
                )
                self.drawString(
                    0.5 * inch, 0.38 * inch,
                    customer_name + "  |  RingCentral ACE  |  Confidential"
                )
                self.setStrokeColor(RC_ORANGE)
                self.setLineWidth(1.5)
                self.line(0.5 * inch, 0.52 * inch, letter[0] - 0.5 * inch, 0.52 * inch)
                pdfcanvas.Canvas.showPage(self)
            pdfcanvas.Canvas.save(self)

    doc = SimpleDocTemplate(
        str(out_path), pagesize=letter,
        leftMargin=0.65 * inch, rightMargin=0.65 * inch,
        topMargin=0.75 * inch,  bottomMargin=0.85 * inch,
        title=customer_name + " -- ACE Transcripts",
        author="RingCentral ACE",
    )

    with_t   = [r for r in records if r.get("transcript")]
    tot_sec  = sum(r.get("duration_sec", 0) for r in records)
    hrs, rem = divmod(tot_sec, 3600)
    mins     = rem // 60
    story    = []

    # Cover header
    hdr = Table(
        [[Paragraph("RingCentral AI Conversation Expert", hdr_txt_s)]],
        colWidths=[doc.width]
    )
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), RC_ORANGE),
        ("TOPPADDING",    (0, 0), (-1, -1), 28),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 22),
        ("LEFTPADDING",   (0, 0), (-1, -1), 20),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 20),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.25 * inch))
    story.append(Paragraph("<b>" + customer_name + "</b>", cover_big_s))
    story.append(Paragraph("Call Transcript Report", cover_sub_s))

    # Stats
    stats_data = [[
        Paragraph(
            "<b>" + str(len(records)) + "</b><br/>"
            "<font size='8' color='#7A7A7A'>Total Calls</font>", stat_s),
        Paragraph(
            "<b>" + str(len(with_t)) + "</b><br/>"
            "<font size='8' color='#7A7A7A'>With Transcripts</font>", stat_or_s),
        Paragraph(
            "<b>" + str(hrs) + "h " + str(mins) + "m</b><br/>"
            "<font size='8' color='#7A7A7A'>Total Talk Time</font>", stat_s),
        Paragraph(
            "<b>" + date_from[:10] + "</b><br/>"
            "<font size='8' color='#7A7A7A'>Start Date</font>",
            ps("SD", fontSize=12, fontName="Helvetica-Bold",
               textColor=DARK, alignment=TA_CENTER, leading=18)),
        Paragraph(
            "<b>" + date_to[:10] + "</b><br/>"
            "<font size='8' color='#7A7A7A'>End Date</font>",
            ps("ED", fontSize=12, fontName="Helvetica-Bold",
               textColor=DARK, alignment=TA_CENTER, leading=18)),
    ]]
    st = Table(stats_data, colWidths=[doc.width / 5] * 5)
    st.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), BG_GREY),
        ("TOPPADDING",    (0, 0), (-1, -1), 16),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 16),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("LINEAFTER",     (0, 0), (3, 0),   0.5, RULE),
    ]))
    story.append(st)
    story.append(Spacer(1, 0.12 * inch))
    story.append(Paragraph(
        "Generated " + datetime.now().strftime("%B %d, %Y at %I:%M %p") +
        "  |  RingCentral AI Conversation Expert  |  Confidential", gen_s))
    story.append(Spacer(1, 0.35 * inch))
    story.append(HRFlowable(width="100%", thickness=2, color=RC_ORANGE))
    story.append(Spacer(1, 0.4 * inch))

    # Call blocks
    sp_styles     = [sp_a_s, sp_b_s]
    calls_to_show = [r for r in records if r.get("transcript") or r.get("summary")]

    for idx, rec in enumerate(calls_to_show, 1):
        direction  = rec.get("direction", "")
        from_name  = rec.get("from_name") or rec.get("from_number") or "Unknown"
        to_name    = rec.get("to_name")   or rec.get("to_number")   or "Unknown"
        start_time = rec.get("start_time", "")
        duration   = rec.get("duration_sec", 0)
        summary    = rec.get("summary", "")
        sentiment  = rec.get("sentiment", "")
        transcript = rec.get("transcript", "")

        try:
            dt       = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            date_str = dt.strftime("%B %d, %Y")
            time_str = dt.strftime("%I:%M %p")
        except Exception:
            date_str = start_time[:10]
            time_str = ""

        dur_m, dur_s = divmod(int(duration), 60)

        sl = sentiment.lower()
        if "positive"  in sl: sent_bg, sent_fg, sent_lbl = BG_GREEN, TX_GREEN, "Positive"
        elif "negative" in sl: sent_bg, sent_fg, sent_lbl = BG_RED,   TX_RED,   "Negative"
        else:                  sent_bg, sent_fg, sent_lbl = BG_GREY,  TX_GREY,  "Neutral"

        dir_bg = BG_BLUE   if direction == "Inbound" else BG_ORANGE
        dir_fg = TX_BLUE   if direction == "Inbound" else HexColor("#9B4A00")

        block = []
        block.append(Paragraph(
            "CALL " + str(idx) + " OF " + str(len(calls_to_show)), call_num_s))

        title = (
            "Inbound call from <b>" + from_name + "</b>"
            if direction == "Inbound"
            else "Outbound call to <b>" + to_name + "</b>"
        )
        block.append(Paragraph(title, call_title_s))

        meta_parts = [p for p in [
            date_str, time_str,
            "Duration: " + str(dur_m) + "m " + str(dur_s) + "s",
            "From: " + from_name, "To: " + to_name,
        ] if p]
        block.append(Paragraph(" &nbsp;|&nbsp; ".join(meta_parts), meta_s))

        badges = [[
            Paragraph("<b>" + direction + "</b>",
                      ps("DB", fontSize=8.5, fontName="Helvetica-Bold",
                         textColor=dir_fg, alignment=TA_CENTER)),
            Paragraph("<b>" + sent_lbl + "</b>",
                      ps("SB2", fontSize=8.5, fontName="Helvetica-Bold",
                         textColor=sent_fg, alignment=TA_CENTER)),
            Paragraph("", ps("SP", fontSize=8)),
        ]]
        bt = Table(badges, colWidths=[1.0 * inch, 1.0 * inch, doc.width - 2.0 * inch])
        bt.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (0, 0), dir_bg),
            ("BACKGROUND",    (1, 0), (1, 0), sent_bg),
            ("TOPPADDING",    (0, 0), (1, 0), 5),
            ("BOTTOMPADDING", (0, 0), (1, 0), 5),
            ("LEFTPADDING",   (0, 0), (1, 0), 10),
            ("RIGHTPADDING",  (0, 0), (1, 0), 10),
        ]))
        block.append(Spacer(1, 4))
        block.append(bt)

        if summary:
            block.append(Paragraph("AI SUMMARY", sum_lbl_s))
            st2 = Table([[Paragraph(summary, sum_txt_s)]], colWidths=[doc.width])
            st2.setStyle(TableStyle([
                ("BACKGROUND",    (0, 0), (-1, -1), BG_ORANGE),
                ("TOPPADDING",    (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ("LEFTPADDING",   (0, 0), (-1, -1), 14),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 14),
                ("LINEBEFORE",    (0, 0), (0, -1),   3, RC_ORANGE),
            ]))
            block.append(st2)

        if transcript:
            block.append(Paragraph("FULL TRANSCRIPT", tr_lbl_s))
            unique_speakers = list(dict.fromkeys(
                line.split(": ")[0].split("] ")[-1].strip()
                for line in transcript.split("\n") if ": " in line
            ))
            sp_map = {
                sp: sp_styles[i % len(sp_styles)]
                for i, sp in enumerate(unique_speakers)
            }
            for line in transcript.split("\n"):
                if not line.strip():
                    continue
                m = re.match(r"^(\[\d+:\d+\])\s+(.+?):\s+(.+)$", line)
                if m:
                    ts2, speaker, text = m.group(1), m.group(2), m.group(3)
                    safe = (text.replace("&", "&amp;")
                               .replace("<", "&lt;")
                               .replace(">", "&gt;"))
                    block.append(Paragraph(
                        "<b>" + speaker + "</b> "
                        "<font size='7' color='#BBBBBB'>" + ts2 + "</font>",
                        sp_map.get(speaker, sp_a_s)
                    ))
                    block.append(Paragraph(safe, utt_s))
                else:
                    safe = (line.replace("&", "&amp;")
                               .replace("<", "&lt;")
                               .replace(">", "&gt;"))
                    block.append(Paragraph(safe, utt_s))
        else:
            block.append(Spacer(1, 6))
            block.append(Paragraph(
                "No transcript available for this call -- "
                "extension may not have a RingSense license.", no_tr_s))

        block.append(Spacer(1, 0.15 * inch))
        block.append(HRFlowable(width="100%", thickness=0.5, color=RULE))
        block.append(Spacer(1, 0.2 * inch))

        story.append(KeepTogether(block[:7]))
        story.extend(block[7:])

    doc.build(story, canvasmaker=NumCanvas)


# ── Save outputs ──────────────────────────────────────────────────────────────
def save_outputs(records, creds, date_from, date_to, script_dir):
    header("STEP 6 -- Saving Files")
    rule()

    customer_name = creds["customer_name"]
    slug          = re.sub(r"[^a-z0-9]+", "_", customer_name.lower()).strip("_")
    date_stamp    = datetime.now().strftime("%Y%m%d")

    # Excel
    xlsx_name = "transcripts_" + slug + "_" + date_stamp + ".xlsx"
    xlsx_path = script_dir / xlsx_name
    info("Building Excel spreadsheet...")
    build_excel(records, customer_name, date_from, date_to, xlsx_path)
    sz = xlsx_path.stat().st_size / 1024
    ok(xlsx_name + "  (" + str(round(sz)) + " KB)")

    # PDF
    pdf_name  = "transcripts_" + slug + "_" + date_stamp + ".pdf"
    pdf_path  = script_dir / pdf_name
    info("Building PDF (may take 15-30 seconds)...")
    build_pdf(records, customer_name, date_from, date_to, pdf_path)
    sz2 = pdf_path.stat().st_size / 1024 / 1024
    ok(pdf_name + "  (" + str(round(sz2, 1)) + " MB)")

    return xlsx_path, pdf_path


# ── Summary ───────────────────────────────────────────────────────────────────
def print_summary(customer_name, total_calls, with_transcripts, xlsx_path, pdf_path):
    print("")
    print(BOLD + G + "+======================================================+" + W)
    print(BOLD + G + "|  COMPLETE                                            |" + W)
    print(BOLD + G + "+======================================================+" + W)
    print("")
    print("  Customer    : " + customer_name)
    print("  Total calls : " + str(total_calls))
    print("  Transcripts : " + str(with_transcripts))
    print("  No transcript: " + str(total_calls - with_transcripts) + " (unlicensed extensions)")
    print("")
    print("  Files saved in the same folder as this script:")
    print("")

    for path in [xlsx_path, pdf_path]:
        size = path.stat().st_size
        sz   = (str(round(size / 1024 / 1024, 1)) + " MB"
                if size > 1024 * 1024
                else str(size // 1024) + " KB")
        ext  = path.suffix
        tag  = "XLSX" if ext == ".xlsx" else "PDF "
        print("    [" + tag + "]  " + path.name.ljust(52) + " " + DIM + sz + W)

    print("")
    print("  To open the folder:")
    print('    open "' + str(xlsx_path.parent) + '"')
    print("")
    print(DIM + "  Run this script again for your next customer." + W)
    print(DIM + "  Credentials are never saved." + W)
    print("")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print_banner()

    # The output folder is the same folder as this script
    script_dir = Path(__file__).parent.resolve()

    creds      = collect_credentials()
    token      = authenticate(creds)
    account_id = get_account_id(token)
    date_from, date_to = choose_date_range()

    print("")
    print("  " + BOLD + "Ready to run:" + W)
    print("    Customer  : " + creds["customer_name"])
    print("    Date range: " + date_from[:10] + " to " + date_to[:10])
    print("    Output to : " + str(script_dir))
    print("")

    if not confirm("Start download?"):
        print("  Cancelled.")
        sys.exit(0)

    call_logs = download_call_logs(token, account_id, date_from, date_to)

    if not call_logs:
        warn("No recorded calls found in this date range.")
        warn("Try a wider date range or check that call recording is enabled.")
        sys.exit(0)

    records, with_transcripts = fetch_transcripts(token, call_logs)
    xlsx_path, pdf_path = save_outputs(records, creds, date_from, date_to, script_dir)

    print_summary(
        creds["customer_name"],
        len(call_logs),
        with_transcripts,
        xlsx_path,
        pdf_path,
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  " + Y + "Cancelled." + W + "\n")
        sys.exit(0)
