"""Demo CRUD functions for the SharePoint list.
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
# Graph site id of the Communication site: "<hostname>,<site collection id>,<web id>".
SITE_ID = "mngenvmcap249728.sharepoint.com,55a308eb-2e41-4795-af9f-97d1c7499a61,50b083ac-f43a-4301-b918-25882a632e72"
# Graph list id of the "Care Gap Outreach" list.
LIST_ID = "6c822deb-a59a-4b98-9617-ea46971c9e29"
GRAPH = "https://graph.microsoft.com/v1.0"
LIST_URL = f"{GRAPH}/sites/{SITE_ID}/lists/{LIST_ID}"

# --- Business vocabulary (must match the list's choice columns and the report slicers) ------
OPEN_STATUSES = ["Not Started", "Attempted - No Answer", "Contacted", "Scheduled"]
CLOSED_STATUSES = ["Completed", "Patient Declined", "Unable to Reach"]
PRIORITIES = ["High", "Medium", "Low"]

# The list has no demographic columns; new rows need only the clinic's coordinator.
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

# Module-level cache for the Graph access token. A warm UDF instance reuses it across calls,
# which skips three Key Vault reads and a token request (about 1-2 seconds per click).
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
    if _token_cache["value"] and _token_cache["expires"] > time.time() + 120:
        return {"Authorization": f"Bearer {_token_cache['value']}", "Content-Type": "application/json"}
    secrets = SecretClient(vault_url=KEY_VAULT_URL, credential=key_vault.get_access_token())
    tenant = secrets.get_secret("sp-writer-tenant-id").value
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={"client_id": secrets.get_secret("sp-writer-client-id").value,
              "client_secret": secrets.get_secret("sp-writer-client-secret").value,
              "grant_type": "client_credentials", "scope": "https://graph.microsoft.com/.default"},
        timeout=30)
    if not resp.ok:
        raise fn.UserThrownError("Could not authenticate to SharePoint (Microsoft Graph token request failed).")
    body = resp.json()
    _token_cache.update(value=body["access_token"], expires=time.time() + int(body.get("expires_in", 3600)))
    return {"Authorization": f"Bearer {body['access_token']}", "Content-Type": "application/json"}


def _find_item(headers: dict, care_gap_id: int) -> dict:
    """Find the SharePoint item by its unique Care Gap ID, not by Patient ID.
    Source errors and missing or duplicate matches are reported to the caller."""
    resp = requests.get(f"{LIST_URL}/items", headers=headers, timeout=30, params={
        "$expand": "fields($select=Title,PatientId,QualityMeasure,OutreachStatus,Priority,NextFollowUpDate,OutreachNotes)",
        "$filter": f"fields/CareGapId eq {int(care_gap_id)}"})
    if not resp.ok:
        raise fn.UserThrownError(f"SharePoint lookup failed ({resp.status_code}).")
    items = resp.json().get("value", [])
    if not items:
        raise fn.UserThrownError(f"Care gap {care_gap_id} was not found in the SharePoint list.")
    if len(items) != 1:
        raise fn.UserThrownError(f"Care gap {care_gap_id} is not unique in the SharePoint list; no change was made.")
    return items[0]


def _next_care_gap_id(sql: fn.FabricSqlConnection) -> int:
    """Reserve a shared, monotonic ID; deleted records cannot give their IDs to new rows."""
    conn = sql.connect()
    try:
        cur = conn.cursor()
        cur.execute("SELECT NEXT VALUE FOR dbo.CareGapIdSequence")
        value = int(cur.fetchone()[0])
        cur.close()
        return value
    finally:
        conn.close()


def _log(sql: fn.FabricSqlConnection, care_gap_id: int, action: str, old: str, new: str, user: str):
    """Append a source-tagged audit entry to the Fabric SQL database.
    The external source write and this SQL commit are separate operations."""
    conn = sql.connect()
    try:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO dbo.WriteBackLog (CareGapId, TargetSystem, ActionName, OldValue, NewValue, ChangedBy) "
            "VALUES (?, 'SharePoint List', ?, ?, ?, ?)", care_gap_id, action, old, (new or "")[:400], user)
        conn.commit()
        cur.close()
    finally:
        conn.close()


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.generic_connection(argName="keyVault", audienceType="KeyVault")
@udf.context(argName="ctx")
@udf.function()
def update_outreach_sp(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                       careGapId: int, outreachStatus: str, priority: str, nextFollowUpDate: str) -> str:
    """Update the selected gap with the report status, priority and optional follow-up.
    Closed statuses clear the date; this demo has no scheduling restrictions.
    The operation also records the old and new values in the SQL audit log."""
    # --- 1. Validate the request (same rules as the Fabric SQL DB function) ---------------
    if outreachStatus not in OPEN_STATUSES + CLOSED_STATUSES:
        raise fn.UserThrownError(f"Outreach status must be one of: {', '.join(OPEN_STATUSES + CLOSED_STATUSES)}.")
    if priority not in PRIORITIES:
        raise fn.UserThrownError("Priority must be High, Medium or Low.")

    follow_up = None if outreachStatus in CLOSED_STATUSES else _parse_date(nextFollowUpDate)

    # --- 2. Authenticate and find the list item -------------------------------------------
    user = _user(ctx)
    headers = _graph_headers(keyVault)
    item = _find_item(headers, careGapId)
    f = item["fields"]
    old_fu = (f.get("NextFollowUpDate") or "")[:10] or "none"
    new_fu = follow_up.isoformat() if follow_up else "none"

    # --- 3. Patch only the columns that change (None clears the date column) -------------
    patch = {"OutreachStatus": outreachStatus, "Priority": priority, "LastUpdatedBy": user,
             "NextFollowUpDate": f"{follow_up.isoformat()}T12:00:00Z" if follow_up else None}
    resp = requests.patch(f"{LIST_URL}/items/{item['id']}/fields", headers=headers, json=patch, timeout=30)
    if not resp.ok:
        raise fn.UserThrownError(f"SharePoint update failed ({resp.status_code}).")

    # --- 4. Audit (only after SharePoint accepted the change) -----------------------------
    _log(sqlDB, careGapId, "Update Outreach",
         f"{f.get('OutreachStatus')} | {f.get('Priority')} | follow-up {old_fu}",
         f"{outreachStatus} | {priority} | follow-up {new_fu}", user)

    suffix = " (gap closed, follow-up cleared)" if outreachStatus in CLOSED_STATUSES else f", follow-up {new_fu}"
    return f"SharePoint updated - {f.get('Title')}, {f.get('QualityMeasure')}: {f.get('OutreachStatus')} -> {outreachStatus}, priority {priority}{suffix}."




def _create_care_gap_sp(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                        patientName: str, clinic: str, qualityMeasure: str, dueDate: str, priority: str,
                        patient_id: str | None = None, outreach_status: str = "Not Started",
                        follow_up: datetime.date | None = None) -> str:
    """Write the validated insert fields to the source, then record the audit entry.
    This is a private helper, not a callable UDF. SQL-only fields use synthetic demo values."""
    # --- 1. Validate the common insert fields --------------------------------------------
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

    # --- 2. Authenticate, allocate a new CareGapId and create the item ---------------------
    user = _user(ctx)
    headers = _graph_headers(keyVault)
    coordinator = CLINIC_COORDINATORS.get(clinic, "Unassigned")
    new_id = _next_care_gap_id(sqlDB)
    patient_id = patient_id or f"P-2{new_id}"

    fields = {
        "Title": name, "CareGapId": new_id, "PatientId": patient_id, "QualityMeasure": quality_measure,
        "Clinic": clinic, "CareCoordinator": coordinator, "DueDate": f"{due_date.isoformat()}T12:00:00Z",
        "OutreachStatus": outreach_status, "Priority": priority, "LastUpdatedBy": user,
    }
    if follow_up is not None:
        fields["NextFollowUpDate"] = f"{follow_up.isoformat()}T12:00:00Z"
    resp = requests.post(f"{LIST_URL}/items", headers=headers, json={"fields": fields}, timeout=30)
    if not resp.ok:
        raise fn.UserThrownError(f"SharePoint create failed ({resp.status_code}).")

    # --- 3. Audit (OldValue is NULL - there is no "before" state for a brand-new item) -----
    try:
        _log(sqlDB, new_id, "Create Care Gap", None,
             (f"{patient_id} | {name} | {clinic} | {quality_measure} | due {due_date.isoformat()} | "
              f"{outreach_status} | {priority} | follow-up {follow_up.isoformat() if follow_up else 'none'}")[:400], user)
    except Exception as exc:
        logging.exception("SharePoint gap %s created, but SQL audit failed", new_id)
        raise fn.UserThrownError(f"SharePoint gap #{new_id} was created, but its audit entry could not be saved. Do not repeat the insert; contact the demo owner.") from exc

    return f"Created care gap #{new_id} in SharePoint for {name} - {quality_measure}."






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
def insert_care_gap_sp(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
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
    return _create_care_gap_sp(sqlDB, keyVault, ctx, name, clinic, measure, due, priority,
                               patient_id, outreachStatus, follow_up)


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.generic_connection(argName="keyVault", audienceType="KeyVault")
@udf.context(argName="ctx")
@udf.function()
def delete_care_gap_sp(sqlDB: fn.FabricSqlConnection, keyVault: fn.FabricItem, ctx: fn.UserDataFunctionContext,
                      careGapId: int) -> str:
    """Delete the single SharePoint item selected in the main report worklist.

    No typed-ID input is required. If-Match protects against deleting
    an item edited since this lookup. Existing SQL audit history is retained. SharePoint
    deletion and SQL logging are separate commits: a logging failure after deletion is
    explicitly reported, never disguised as a rollback or a successful audit.
    """
    if not isinstance(careGapId, int) or isinstance(careGapId, bool) or careGapId <= 0:
        raise fn.UserThrownError("Select exactly one care gap in the main worklist before deleting.")
    headers = _graph_headers(keyVault)
    item = _find_item(headers, careGapId)
    etag = item.get("eTag") or item.get("@odata.etag")
    if not etag:
        raise fn.UserThrownError("SharePoint did not return an item version; no record was deleted.")
    resp = requests.delete(f"{LIST_URL}/items/{item['id']}", headers={**headers, "If-Match": etag}, timeout=30)
    if resp.status_code != 204:
        raise fn.UserThrownError(f"SharePoint delete failed ({resp.status_code}). Refresh the worklist before retrying.")
    fields = item["fields"]
    try:
        _log(sqlDB, careGapId, "Delete Care Gap",
             f"{fields.get('PatientId')} | {fields.get('Title')} | {fields.get('QualityMeasure')}"[:400],
             "Permanently deleted (including embedded notes)", _user(ctx))
    except Exception as exc:
        logging.exception("SharePoint gap %s deleted, but SQL audit failed", careGapId)
        raise fn.UserThrownError(f"SharePoint gap #{careGapId} was deleted, but its audit entry could not be saved. Do not repeat the deletion; contact the demo owner.") from exc
    return f"Deleted SharePoint care gap #{careGapId}. Audit history retained."
