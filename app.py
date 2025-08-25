from flask import Flask, request, render_template_string, redirect, url_for, session
import os
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime, date, timedelta
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# -------------------------
# Google API Setup
# -------------------------
SCOPE = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Use your downloaded credential file
CREDS = Credentials.from_service_account_file("hr-work-log-458400127634", scopes=SCOPE)
gc = gspread.authorize(CREDS)
service = build("sheets", "v4", credentials=CREDS)

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "").strip() or None
SPREADSHEET_NAME = "Employee_worklog"
if SPREADSHEET_ID:
    SPREAD = gc.open_by_key(SPREADSHEET_ID)
else:
    SPREAD = gc.open(SPREADSHEET_NAME)

sheet = SPREAD.sheet1
SPREADSHEET_ID = SPREAD.id
SHEET_GID = sheet.id

FIRST_LOG_DMY = "07-08-2025"
FIRST_LOG_DATE = datetime.strptime(FIRST_LOG_DMY, "%d-%m-%Y").date()

# -------------------------
# Helpers
# -------------------------
def is_leave(text: str) -> bool:
    return (text or "").strip().lower() == "leave"

def is_sunday(d: date) -> bool:
    return d.weekday() == 6  # Sunday = 6

def get_headers_raw():
    return sheet.row_values(2)

def parse_header_date(raw: str):
    if not raw or not raw.strip():
        return None
    raw = raw.strip()
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None

def headers_info():
    raws = get_headers_raw()
    info = []
    for i, r in enumerate(raws):
        info.append({
            "col_index": i,           # 0-based col index within the sheet
            "raw": r,
            "date": parse_header_date(r),
            "lower": (r or "").strip().lower()
        })
    return info

def find_header_col_by_dmy(target_dmy: str):
    info = headers_info()
    for h in info:
        if h["date"] and h["date"].strftime("%d-%m-%Y") == target_dmy:
            return h["col_index"]
    return None

def find_header_index_by_name(name_lower: str):
    info = headers_info()
    for h in info:
        if h["lower"] == name_lower:
            return h["col_index"]
    return None

def find_user_row_by_email(email: str, email_col_idx: int):
    all_rows = sheet.get_all_values()
    for i, row in enumerate(all_rows[2:], start=3):  # data starts from row 3
        if len(row) > email_col_idx and row[email_col_idx].strip().lower() == email:
            return i
    return None

def ensure_date_columns_up_to_today():
    headers = get_headers_raw()
    existing_normalized = set()
    for r in headers:
        d = parse_header_date(r)
        if d:
            existing_normalized.add(d.strftime("%d-%m-%Y"))
    today = date.today()
    cur = FIRST_LOG_DATE
    while cur <= today:
        s = cur.strftime("%d-%m-%Y")
        if s not in existing_normalized:
            # Row 2 holds date headers
            sheet.update_cell(2, len(headers) + 1, s)
            headers.append(s)
            existing_normalized.add(s)
        cur += timedelta(days=1)

# ---------- Formatting helpers ----------
def _batch_update(requests: list):
    """Small helper to send batchUpdate safely."""
    if not requests:
        return
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests}
    ).execute()

def format_cell(row_1b: int, col_1b: int, text: str, *, red: bool = False):
    """
    Update a single cell's value and apply Times New Roman font.
    If red=True, also set the text color to red.
    Indices expected are 1-based (as used by gspread's update_cell).
    """
    text_format = {"fontFamily": "Times New Roman"}
    if red:
        text_format["foregroundColor"] = {"red": 1.0, "green": 0.0, "blue": 0.0}
    requests = [{
        "updateCells": {
            "rows": [{
                "values": [{
                    "userEnteredValue": {"stringValue": text},
                    "userEnteredFormat": {"textFormat": text_format}
                }]
            }],
            "fields": "userEnteredValue,userEnteredFormat.textFormat",
            "range": {
                "sheetId": SHEET_GID,
                "startRowIndex": row_1b - 1,
                "endRowIndex": row_1b,
                "startColumnIndex": col_1b - 1,
                "endColumnIndex": col_1b
            }
        }
    }]
    _batch_update(requests)

def auto_fill_leave_sunday():
    """
    Fill 'Leave' for leave cells and 'Sunday' for Sunday cells in the sheet automatically.
    - 'Sunday' is written in RED text.
    - All values written by this function are in Times New Roman.
    """
    headers = headers_info()
    email_col = find_header_index_by_name("email")
    if email_col is None:
        return
    all_rows = sheet.get_all_values()
    for row_idx, row in enumerate(all_rows[2:], start=3):  # data rows start at 3
        for h in headers:
            d = h["date"]
            if not d:
                continue
            col_idx0 = h["col_index"]  # 0-based
            current_val = row[col_idx0] if col_idx0 < len(row) else ""
            if is_leave(current_val):
                format_cell(row_idx, col_idx0 + 1, "Leave", red=False)
            elif is_sunday(d):
                if not current_val.strip():
                    format_cell(row_idx, col_idx0 + 1, "Sunday", red=True)

# ---------- Detect red locks ----------
def _fetch_grid_with_formatting():
    """Fetch full sheet grid with formatting."""
    result = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        includeGridData=True,
        ranges=sheet.title
    ).execute()
    sheets_meta = result.get("sheets", [])
    if not sheets_meta:
        return []
    data_blocks = sheets_meta[0].get("data", [])
    if not data_blocks:
        return []
    return data_blocks[0].get("rowData", [])

def _is_text_red(cell_value: dict) -> bool:
    """Return True if effective text color is red (with tolerance)."""
    fmt = (cell_value or {}).get("effectiveFormat", {}).get("textFormat", {})
    color = fmt.get("foregroundColor", {})
    r = color.get("red", 0.0) or 0.0
    g = color.get("green", 0.0) or 0.0
    b = color.get("blue", 0.0) or 0.0
    return (r > 0.8) and (g < 0.3) and (b < 0.3)

def get_red_marked_cells():
    """
    Return a set of (row_index, col_index) (1-based) where the TEXT COLOR is red.
    """
    red_cells = set()
    grid = _fetch_grid_with_formatting()
    for r_idx_0, row in enumerate(grid):
        for c_idx_0, cell in enumerate(row.get("values", [])):
            if _is_text_red(cell):
                red_cells.add((r_idx_0 + 1, c_idx_0 + 1))
    return red_cells

# -------------------------
# Templates
# -------------------------
BOOTSTRAP = '<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">'
LOGIN_HTML = f"""
<!doctype html>
<html>
<head>{BOOTSTRAP}</head>
<body class="bg-light">
<div class="container py-5" style="max-width:520px;">
  <div class="card p-4 shadow-sm">
    <h3 class="text-center mb-3">Employee Login</h3>
    <form method="post">
      <input type="email" name="email" class="form-control mb-3" placeholder="Email" required autofocus>
      <button class="btn btn-primary w-100">Login</button>
    </form>
    {{% if error %}}
      <div class="alert alert-danger mt-3">{{{{ error }}}}</div>
    {{% endif %}}
  </div>
</div>
</body>
</html>
"""
WORKLOG_HTML = f"""
<!doctype html>
<html>
<head>{BOOTSTRAP}</head>
<body class="bg-light">
<div class="container py-4">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <div>
      <h4 class="mb-0">Welcome, {{{{ name }}}}</h4>
      <small class="text-muted">{{{{ email }}}}</small>
    </div>
    <div>
      <a href="{{{{ url_for('logout') }}}}" class="btn btn-outline-secondary">Logout</a>
    </div>
  </div>
  {{% if flash_msg %}}
    <div class="alert alert-{{{{ flash_type }}}}">{{{{ flash_msg }}}}</div>
  {{% endif %}}
  <div class="card p-3">
    <h5 class="mb-3">Work log (last 5 days)</h5>
    <div class="table-responsive">
      <table class="table table-sm table-bordered align-middle">
        <thead class="table-light">
          <tr><th>Date</th><th>Log</th><th style="width:160px">Action</th></tr>
        </thead>
        <tbody>
          {{% for dmy, text, is_locked, lock_reason in logs %}}
            <tr>
              <td style="white-space:nowrap;">{{{{ dmy }}}}</td>
              <td style="white-space:pre-wrap; {{ 'opacity:0.6;' if is_locked else '' }}">
                {{{{ text }}}}
                {{% if is_locked and not text %}}
                  <span class="text-muted fst-italic">( {{ lock_reason }} )</span>
                {{% endif %}}
              </td>
              <td>
                {{% if is_locked %}}
                  <button class="btn btn-secondary btn-sm w-100" style="opacity:0.6;cursor:not-allowed;" title="{{{{ lock_reason }}}}">
                    {{{{ text if text else lock_reason }}}}
                  </button>
                {{% else %}}
                  <a href="{{{{ url_for('edit_log', dmy=dmy) }}}}" class="btn btn-warning btn-sm w-100">Edit</a>
                {{% endif %}}
              </td>
            </tr>
          {{% endfor %}}
        </tbody>
      </table>
    </div>
  </div>
</div>
</body>
</html>
"""
EDIT_HTML = f"""
<!doctype html>
<html>
<head>{BOOTSTRAP}</head>
<body class="bg-light">
<div class="container py-5" style="max-width:900px;">
  <a href="{{{{ url_for('worklog') }}}}" class="btn btn-link mb-3">&larr; Back to logs</a>
  <div class="card p-4">
    <h4 class="mb-3">Edit log for <strong>{{{{ dmy }}}}</strong></h4>
    {{% if is_locked %}}
      <div class="alert alert-warning">
        This date is <strong>locked</strong> ({{{{ lock_reason }}}}). Editing is disabled.
      </div>
      <pre class="mb-0" style="opacity:0.7;white-space:pre-wrap;">{{{{ log_text }}}}</pre>
    {{% else %}}
      <form method="post">
        <div class="mb-3">
          <textarea name="log" class="form-control" rows="8" placeholder="Enter work details...">{{{{ log_text }}}}</textarea>
        </div>
        <button class="btn btn-success">Save</button>
      </form>
    {{% endif %}}
  </div>
</div>
</body>
</html>
"""

# -------------------------
# Routes
# -------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        headers = get_headers_raw()
        headers_lower = [h.strip().lower() for h in headers]
        if "name" not in headers_lower or "email" not in headers_lower:
            return render_template_string(LOGIN_HTML, error="❌ Sheet missing NAME or EMAIL in row 2.")
        name_col = headers_lower.index("name")
        email_col = headers_lower.index("email")
        row = find_user_row_by_email(email, email_col)
        if row is None:
            return render_template_string(LOGIN_HTML, error="❌ Email not found. Please try again.")
        row_vals = sheet.row_values(row)
        user_name = row_vals[name_col] if len(row_vals) > name_col else ""
        session["email"] = email
        session["name"] = user_name
        return redirect(url_for("worklog"))
    return render_template_string(LOGIN_HTML)

@app.route("/worklog")
def worklog():
    if "email" not in session:
        return redirect(url_for("login"))
    ensure_date_columns_up_to_today()
    auto_fill_leave_sunday()
    info = headers_info()
    email_col = find_header_index_by_name("email")
    if email_col is None:
        return render_template_string(
            WORKLOG_HTML,
            name=session["name"],
            email=session["email"],
            logs=[],
            flash_msg="❌ Sheet missing EMAIL column",
            flash_type="danger"
        )
    row_index = find_user_row_by_email(session["email"], email_col)
    if row_index is None:
        return render_template_string(
            WORKLOG_HTML,
            name=session["name"],
            email=session["email"],
            logs=[],
            flash_msg="❌ User row not found",
            flash_type="danger"
        )
    row_vals = sheet.row_values(row_index)
    today = date.today()
    five_days_ago = today - timedelta(days=4)
    red_cells = get_red_marked_cells()
    logs = []
    seen = set()
    for h in info:
        d = h["date"]
        if not d or d < FIRST_LOG_DATE or d > today:
            continue
        if d < five_days_ago:
            continue
        dmy = d.strftime("%d-%m-%Y")
        if dmy in seen:
            continue
        seen.add(dmy)
        col_idx0 = h["col_index"]
        text = row_vals[col_idx0] if col_idx0 < len(row_vals) else ""
        col_1b = col_idx0 + 1
        row_1b = row_index
        is_red_locked = (row_1b, col_1b) in red_cells
        is_std_locked = is_leave(text) or is_sunday(d)
        is_locked = is_std_locked or is_red_locked
        reason = "Locked by owner (red text)" if is_red_locked else ("Sunday/Leave" if is_std_locked else "")
        logs.append((dmy, text, is_locked, reason))
    logs.sort(key=lambda x: datetime.strptime(x[0], "%d-%m-%Y").date())
    return render_template_string(
        WORKLOG_HTML,
        name=session["name"],
        email=session["email"],
        logs=logs,
        flash_msg=None,
        flash_type="success"
    )

@app.route("/edit/<dmy>", methods=["GET", "POST"])
def edit_log(dmy):
    if "email" not in session:
        return redirect(url_for("login"))
    try:
        selected_date = datetime.strptime(dmy, "%d-%m-%Y").date()
    except ValueError:
        return redirect(url_for("worklog"))
    if selected_date > date.today():
        return redirect(url_for("worklog"))
    col_idx0 = find_header_col_by_dmy(dmy)
    if col_idx0 is None:
        headers = get_headers_raw()
        sheet.update_cell(2, len(headers) + 1, dmy)
        col_idx0 = find_header_col_by_dmy(dmy)
    email_col = find_header_index_by_name("email")
    if email_col is None:
        return redirect(url_for("worklog"))
    row_idx = find_user_row_by_email(session["email"], email_col)
    if row_idx is None:
        return redirect(url_for("worklog"))
    row_vals = sheet.row_values(row_idx)
    current_text = row_vals[col_idx0] if col_idx0 < len(row_vals) else ""
    red_cells = get_red_marked_cells()
    is_red_locked = (row_idx, col_idx0 + 1) in red_cells
    is_std_locked = is_leave(current_text) or is_sunday(selected_date)
    is_locked = is_std_locked or is_red_locked
    lock_reason = "Locked by owner (red text)" if is_red_locked else ("Sunday/Leave" if is_std_locked else "")
    if request.method == "POST":
        if is_locked:
            return render_template_string(
                EDIT_HTML,
                dmy=dmy,
                log_text=current_text,
                is_locked=True,
                lock_reason=lock_reason
            )
        new_text = request.form.get("log", "").strip()
        format_cell(row_idx, col_idx0 + 1, new_text, red=False)
        return redirect(url_for("worklog"))
    return render_template_string(
        EDIT_HTML,
        dmy=dmy,
        log_text=current_text,
        is_locked=is_locked,
        lock_reason=lock_reason
    )

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/healthz")
def healthz():
    # quick ping to verify auth and sheet access
    try:
        _ = sheet.title
        return "ok", 200
    except Exception as e:
        return f"error: {e}", 500

# -------------------------
# Run app
# -------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)


