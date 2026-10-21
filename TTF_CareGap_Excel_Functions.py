"""Demo CRUD functions for the Excel workbook.
Power BI buttons insert, update or delete a care gap, then record an audit entry.
For this demo, follow-up dates are optional and may be past or future.
Closed outreach statuses still clear the date. Authentication, source-compatible
values and safe single-record deletion are retained. Not a production scheduling policy."""

import datetime
import logging
import re
import time
import requests
import fabric.functions as fn
from azure.keyvault.secrets import SecretClient

udf = fn.UserDataFunctions()

# --- Environment-specific identifiers --------------------------------------------------------
# Key Vault that holds the app registration's credentials (secret names are used in _graph_headers).
KEY_VAULT_URL = "https://kv-ttf-caregap-asyni.vault.azure.net/"
# Graph drive id of the Communication site's "Documents" library, and the item id of the workbook.
DRIVE_ID = "b!6wijVUEulUevn5fRx0maYayDsFA69AFDuRgliCpjLnKfBP_k6VWKRqW9gM5nWaJB"
ITEM_ID = "01AVGXI3B7HJ35DEYM2RBIZAE4DRUIXCDP"
# Name of the Excel table (ListObject) on the "Outreach" sheet.
TABLE = "CareGapOutreach"
WORKBOOK = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook"

# Zero-based column positions in the CareGapOutreach table. Full header order:
#   0 Care Gap Id | 1 Patient Id | 2 Patient Name | 3 Quality Measure | 4 Clinic | 5 Care Coordinator
#   6 Due Date | 7 Outreach Status | 8 Priority | 9 Next Follow-Up Date | 10 Outreach Notes
#   11 Last Updated By | 12 Last Updated (UTC)
C_ID, C_PATIENTID, C_NAME, C_MEASURE, C_CLINIC, C_COORD = 0, 1, 2, 3, 4, 5
C_DUE, C_STATUS, C_PRIORITY, C_FOLLOWUP, C_NOTES, C_BY, C_AT = 6, 7, 8, 9, 10, 11, 12

# Maps each clinic (as offered by the Clinic tile slicer on Add New Care Gap) to its care
# coordinator, used only by create_care_gap_xl to fill the Care Coordinator cell of a new row.
CLINIC_COORDINATORS = {
    "Northside Family Medicine": "Jasmine Reed",
    "Riverside Community Clinic": "Tom Becker",
    "Lakeview Primary Care": "Ana Morales",
    "Downtown Health Center": "Kevin Liu",
}
QUALITY_MEASURES = (
    "Diabetes: HbA1c Test", "Breast Cancer Screening", "Colorectal Cancer Screening",
    "Controlling High Blood Pressure", "Annual Wellness Visit", "Statin Therapy for CVD",
)

# --- Business vocabulary (must match the report slicers and the other two targets) ----------
OPEN_STATUSES = ["Not Started", "Attempted - No Answer", "Contacted", "Scheduled"]
CLOSED_STATUSES = ["Completed", "Patient Declined", "Unable to Reach"]
PRIORITIES = ["High", "Medium", "Low"]

# Module-level cache for the Graph access token; a warm UDF instance reuses it across calls.
_token_cache = {"value": None, "expires": 0}


def _parse_date(value: str) -> datetime.date | None:
    """Read an optional ISO follow-up date; blank or None means no planned follow-up."""
    text = (value or "").strip()
    if text in ("", "None"):
        return None
    try:
        return datetime.date.fromisoformat(text[:10])
    except ValueError:
        raise fn.UserThrownError(f"'{value}' is not a valid follow-up date. Use YYYY-MM-DD.")


def _user(ctx: fn.UserDataFunctionContext) -> str:
    """Read the caller identity for the audit log, falling back to the object ID or unknown."""
    try:
        return ctx.executing_user.get("PreferredUsername") or ctx.executing_user.get("Oid") or "unknown"
    except Exception:
        return "unknown"


def _graph_headers(key_vault: fn.FabricItem) -> dict:
    """Read app credentials from Key Vault and obtain a Microsoft Graph token.
    Reuse a cached token until close to expiry. Never embed or log credentials."""
    if not (_token_cache["value"] and _token_cache["expires"] > time.time() + 120):
        secrets = SecretClient(vault_url=KEY_VAULT_URL, credential=key_vault.get_access_token())
        tenant = secrets.get_secret("sp-writer-tenant-id").value
        resp = requests.post(
            f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
            data={"client_id": secrets.get_secret("sp-writer-client-id").value,
                  "client_secret": secrets.get_secret("sp-writer-client-secret").value,
                  "grant_type": "client_credentials", "scope": "https://graph.microsoft.com/.default"},
            timeout=30)
        if not resp.ok:
            raise fn.UserThrownError("Could not authenticate to Microsoft Graph for the Excel workbook.")
        body = resp.json()
        _token_cache.update(value=body["access_token"], expires=time.time() + int(body.get("expires_in", 3600)))
    return {"Authorization": f"Bearer {_token_cache['value']}", "Content-Type": "application/json"}


def _raise_for_excel(resp: requests.Response, action: str):
    """Report Excel API errors, including a readable message when the workbook is locked."""
    if resp.ok:
        return
    if resp.status_code == 423 or "locked" in resp.text.lower():
        raise fn.UserThrownError("The workbook is locked for editing (likely open in desktop Excel). Close it and try again.")
    raise fn.UserThrownError(f"Excel {action} failed ({resp.status_code}).")


def _find_row(headers: dict, care_gap_id: int):
    """Find an Excel table row by Care Gap ID and return its index and cell values.
    The header is excluded from the zero-based data-row index."""
    resp = requests.get(f"{WORKBOOK}/tables/{TABLE}/range", headers=headers, params={"$select": "values"}, timeout=30)
    _raise_for_excel(resp, "read")
    values = resp.json()["values"]
    for i, row in enumerate(values[1:]):  # values[0] is the header row
        try:
            if int(float(row[C_ID])) == int(care_gap_id):
                return i, row
        except (TypeError, ValueError):
            continue  # skip blank or non-numeric rows rather than failing
    raise fn.UserThrownError(f"Care gap {care_gap_id} was not found in the Excel workbook.")


def _next_care_gap_id(sql: fn.FabricSqlConnection) -> int:
    """Reserve a shared, monotonic ID; deletion must not cause historical IDs to be reused."""
    conn = sql.connect()
    try:
        cur = conn.cursor()
        cur.execute("SELECT NEXT VALUE FOR dbo.CareGapIdSequence")
        value = int(cur.fetchone()[0])
        cur.close()
        return value
    finally:
        conn.close()


def _write_row(headers: dict, index: int, row: list):
    """Save one Excel table row, preserving the unchanged cell values read from the source."""
    resp = requests.patch(f"{WORKBOOK}/tables/{TABLE}/rows/itemAt(index={index})", headers=headers,
                          json={"values": [row]}, timeout=30)
    _raise_for_excel(resp, "update")


def _log(sql: fn.FabricSqlConnection, care_gap_id: int, action: str, old, new: str, user: str):
    """Append a source-tagged audit entry to the Fabric SQL database.
    The external source write and this SQL commit are separate operations."""
    conn = sql.connect()
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO dbo.WriteBackLog (CareGapId, TargetSystem, ActionName, OldValue, NewValue, ChangedBy) "
            "VALUES (?, 'Excel', ?, ?, ?, ?)", care_gap_id, action, old, (new or "")[:400], user)
        conn.commit()
        cur.close()
    finally:
        conn.close()


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.generic_connection(argName="keyVault", audienceType="KeyVault")
@udf.context(argName="ctx")
@udf.function()
def update_outreach_xl(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                       careGapId: int, outreachStatus: str, priority: str, nextFollowUpDate: str) -> str:
    """Update the selected gap with the report status, priority and optional follow-up.
    Closed statuses clear the date; this demo has no scheduling restrictions.
    The operation also records the old and new values in the SQL audit log."""
    # --- 1. Validate the request (same rules as the other two targets) -------------------
    if outreachStatus not in OPEN_STATUSES + CLOSED_STATUSES:
        raise fn.UserThrownError(f"Outreach status must be one of: {', '.join(OPEN_STATUSES + CLOSED_STATUSES)}.")
    if priority not in PRIORITIES:
        raise fn.UserThrownError("Priority must be High, Medium or Low.")

    follow_up = None if outreachStatus in CLOSED_STATUSES else _parse_date(nextFollowUpDate)

    # --- 2. Authenticate and find the row ------------------------------------------------
    user = _user(ctx)
    headers = _graph_headers(keyVault)
    index, row = _find_row(headers, careGapId)
    old = f"{row[C_STATUS]} | {row[C_PRIORITY]} | follow-up {row[C_FOLLOWUP] or 'none'}"
    new_fu = follow_up.isoformat() if follow_up else ""

    # --- 3. Update the in-memory row, then write it back in one call ---------------------
    row[C_STATUS], row[C_PRIORITY], row[C_FOLLOWUP] = outreachStatus, priority, new_fu
    row[C_BY] = user
    row[C_AT] = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
    _write_row(headers, index, row)

    # --- 4. Audit (only after Excel accepted the change) ---------------------------------
    _log(sqlDB, careGapId, "Update Outreach", old,
         f"{outreachStatus} | {priority} | follow-up {new_fu or 'none'}", user)
    suffix = " (gap closed, follow-up cleared)" if outreachStatus in CLOSED_STATUSES else f", follow-up {new_fu or 'none'}"
    return f"Excel updated - {row[C_NAME]}, {row[C_MEASURE]}: now {outreachStatus}, priority {priority}{suffix}."




def _create_care_gap_xl(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                        patientName: str, clinic: str, qualityMeasure: str, dueDate: str, priority: str,
                        patient_id: str | None = None, outreach_status: str = "Not Started",
                        follow_up: datetime.date | None = None) -> str:
    """Write the validated insert fields to the source, then record the audit entry.
    This is a private helper, not a callable UDF. SQL-only fields use synthetic demo values."""
    # --- 1. Validate the request (same rules as the other two targets) -------------------
    name = (patientName or "").strip()
    if not name:
        raise fn.UserThrownError("Please type a patient name before creating a care gap.")
    clinic = (clinic or "").strip()
    quality_measure = (qualityMeasure or "").strip()
    if not clinic:
        raise fn.UserThrownError("Please select a clinic before creating a care gap.")
    if not quality_measure:
        raise fn.UserThrownError("Please select a quality measure before creating a care gap.")
    if priority not in PRIORITIES:
        raise fn.UserThrownError("Priority must be High, Medium or Low.")
    try:
        due_date = datetime.date.fromisoformat((dueDate or "").strip()[:10])
    except ValueError:
        raise fn.UserThrownError(f"'{dueDate}' is not a valid due date. Use YYYY-MM-DD.")

    # --- 2. Authenticate, allocate a new Care Gap Id and build the new row -----------------
    user = _user(ctx)
    headers = _graph_headers(keyVault)
    coordinator = CLINIC_COORDINATORS.get(clinic, "Unassigned")
    new_id = _next_care_gap_id(sqlDB)
    patient_id = patient_id or f"P-2{new_id}"
    now = datetime.datetime.utcnow()

    # Column order must match the C_* constants: Care Gap Id, Patient Id, Patient Name,
    # Quality Measure, Clinic, Care Coordinator, Due Date, Outreach Status, Priority,
    # Next Follow-Up Date, Outreach Notes, Last Updated By, Last Updated (UTC).
    row = [new_id, patient_id, name, quality_measure, clinic, coordinator, due_date.isoformat(),
           outreach_status, priority, follow_up.isoformat() if follow_up else "", "",
           user, now.strftime("%Y-%m-%d %H:%M:%S")]

    resp = requests.post(f"{WORKBOOK}/tables/{TABLE}/rows", headers=headers, json={"values": [row]}, timeout=30)
    _raise_for_excel(resp, "create")

    # --- 3. Audit (OldValue is NULL - there is no "before" state for a brand-new row) ------
    try:
        _log(sqlDB, new_id, "Create Care Gap", None,
             (f"{patient_id} | {name} | {clinic} | {quality_measure} | due {due_date.isoformat()} | "
              f"{outreach_status} | {priority} | follow-up {follow_up.isoformat() if follow_up else 'none'}")[:400], user)
    except Exception as exc:
        logging.exception("Excel gap %s created, but SQL audit failed", new_id)
        raise fn.UserThrownError(f"Excel gap #{new_id} was created, but its audit entry could not be saved. Do not repeat the insert; contact the demo owner.") from exc

    return f"Created care gap #{new_id} in Excel for {name} - {quality_measure}."






def _insert_values(patientName: str, clinic: str, qualityMeasure: str, dueDate: str,
                   priority: str, outreachStatus: str, nextFollowUpDate: str):
    """Check source-compatible fields and parse dates; no demo scheduling rules are enforced."""
    name, clinic, measure = (patientName or "").strip(), (clinic or "").strip(), (qualityMeasure or "").strip()
    if not name or len(name) > 100:
        raise fn.UserThrownError("Enter a patient name between 1 and 100 characters.")
    if clinic not in CLINIC_COORDINATORS:
        raise fn.UserThrownError("Select one of the demo clinics.")
    if measure not in QUALITY_MEASURES:
        raise fn.UserThrownError("Select one of the demo quality measures.")
    if priority not in PRIORITIES:
        raise fn.UserThrownError("Priority must be High, Medium or Low.")
    if outreachStatus not in OPEN_STATUSES + CLOSED_STATUSES:
        raise fn.UserThrownError("Select a valid outreach status.")
    try:
        due = datetime.date.fromisoformat((dueDate or "").strip())
    except ValueError:
        raise fn.UserThrownError("Enter a valid due date using YYYY-MM-DD.")
    follow_up = None if outreachStatus in CLOSED_STATUSES else _parse_date(nextFollowUpDate)
    return name, clinic, measure, due.isoformat(), follow_up


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.generic_connection(argName="keyVault", audienceType="KeyVault")
@udf.context(argName="ctx")
@udf.function()
def insert_care_gap_xl(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                       patientId: str, patientName: str, clinic: str, qualityMeasure: str, dueDate: str,
                       priority: str, outreachStatus: str, nextFollowUpDate: str) -> str:
    """Insert the eight report fields as a new care gap and record an audit entry.
    Patient IDs may repeat across gaps, but each gap receives a unique non-reused ID.
    Follow-up is optional; closed statuses clear it. No patient-master lookup is performed."""
    patient_id = (patientId or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,9}", patient_id):
        raise fn.UserThrownError("Patient ID must be 1-10 letters, digits, hyphens or underscores, starting with a letter or digit.")
    name, clinic, measure, due, follow_up = _insert_values(
        patientName, clinic, qualityMeasure, dueDate, priority, outreachStatus, nextFollowUpDate)
    return _create_care_gap_xl(sqlDB, keyVault, ctx, name, clinic, measure, due, priority,
                               patient_id, outreachStatus, follow_up)


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.generic_connection(argName="keyVault", audienceType="KeyVault")
@udf.context(argName="ctx")
@udf.function()
def delete_care_gap_xl(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                      careGapId: int) -> str:
    """Delete the single workbook row selected in the main report worklist.

    Match by Care Gap ID, never by Patient ID; duplicate matches fail closed. Recheck the
    resolved table row immediately before deleting since Excel row indexes can move.
    This demo assumes no concurrent row reordering between that check and DELETE: Excel
    does not provide a transaction spanning the lookup and row deletion. Do not sort/edit
    this workbook in another session while presenting a delete.

    The SQL audit write follows the workbook commit. If it fails, report that the source
    row was already removed, rather than implying that deletion rolled back.
    """
    if not isinstance(careGapId, int) or isinstance(careGapId, bool) or careGapId <= 0:
        raise fn.UserThrownError("Select exactly one care gap in the main worklist before deleting.")
    headers = _graph_headers(keyVault)
    resp = requests.get(f"{WORKBOOK}/tables/{TABLE}/range", headers=headers,
                        params={"$select": "values"}, timeout=30)
    _raise_for_excel(resp, "read")
    matches = []
    for index, row in enumerate(resp.json()["values"][1:]):
        try:
            value = float(row[C_ID])
        except (TypeError, ValueError):
            logging.warning("Non-numeric Care Gap ID in Excel table row %s", index)
            continue
        if value == careGapId:
            matches.append((index, row))
    if not matches:
        raise fn.UserThrownError(f"Care gap {careGapId} was not found in Excel. Refresh the worklist.")
    if len(matches) != 1:
        raise fn.UserThrownError(f"Care gap {careGapId} is duplicated in Excel; no row was deleted.")
    index, row = matches[0]
    row_url = f"{WORKBOOK}/tables/{TABLE}/rows/itemAt(index={index})"
    current = requests.get(row_url, headers=headers, timeout=30)
    _raise_for_excel(current, "read")
    if current.json()["values"] != [row]:
        raise fn.UserThrownError("The Excel row changed during deletion. Refresh and select it again; no row was deleted.")
    deleted = requests.delete(row_url, headers=headers, timeout=30)
    _raise_for_excel(deleted, "delete")
    try:
        _log(sqlDB, careGapId, "Delete Care Gap",
             f"{row[C_PATIENTID]} | {row[C_NAME]} | {row[C_MEASURE]}"[:400],
             "Permanently deleted (including embedded notes)", _user(ctx))
    except Exception as exc:
        logging.exception("Excel gap %s deleted, but SQL audit failed", careGapId)
        raise fn.UserThrownError(f"Excel gap #{careGapId} was deleted, but its audit entry could not be saved. Do not repeat the deletion; contact the demo owner.") from exc
    return f"Deleted Excel care gap #{careGapId}. Audit history retained."
