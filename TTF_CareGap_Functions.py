"""Demo CRUD functions for the Fabric SQL Database.
Power BI buttons insert, update or delete a care gap, then record an audit entry.
For this demo, follow-up dates are optional and may be past or future.
Closed outreach statuses still clear the date. Authentication, source-compatible
values and safe single-record deletion are retained. Not a production scheduling policy."""

import datetime
import random
import re
import fabric.functions as fn

udf = fn.UserDataFunctions()

# Outreach statuses that keep a care gap "open" (still being worked by a coordinator).
OPEN_STATUSES = ["Not Started", "Attempted - No Answer", "Contacted", "Scheduled"]
# Outreach statuses that end outreach; a closed gap never keeps a follow-up date.
CLOSED_STATUSES = ["Completed", "Patient Declined", "Unable to Reach"]
# Allowed care-team priorities (must match the CHECK constraint on dbo.CareGap.Priority).
PRIORITIES = ["High", "Medium", "Low"]

# Demographic/plan values used only to fill the NOT NULL columns that the "quick add" form
# on the Add New Care Gap page doesn't ask the user for. These are illustrative demo values,
# not meant to resemble real patient data.
PAYER_TYPES = ["Medicare Advantage", "Medicaid", "Commercial"]

# Maps each clinic (as offered by the Clinic tile slicer) to its care coordinator and primary
# care provider, mirroring the clinic roster used when the original synthetic data was
# generated. A clinic typed or selected outside this list falls back to "Unassigned".
CLINIC_STAFF = {
    "Northside Family Medicine": ("Jasmine Reed", "Dr. Helen Park"),
    "Riverside Community Clinic": ("Tom Becker", "Dr. Omar Haddad"),
    "Lakeview Primary Care": ("Ana Morales", "Dr. Susan Whitfield"),
    "Downtown Health Center": ("Kevin Liu", "Dr. Raj Menon"),
}

# Maps each quality measure (as offered by the Quality Measure tile slicer) to the short
# code stored in dbo.CareGap.MeasureCode. A measure outside this list falls back to "OTH".
MEASURE_CODES = {
    "Diabetes: HbA1c Test": "CDC",
    "Breast Cancer Screening": "BCS",
    "Colorectal Cancer Screening": "COL",
    "Controlling High Blood Pressure": "CBP",
    "Annual Wellness Visit": "AWV",
    "Statin Therapy for CVD": "SPC",
}



def _user(ctx: fn.UserDataFunctionContext) -> str:
    """Read the caller identity for the audit log, falling back to the object ID or unknown."""
    try:
        return ctx.executing_user.get("PreferredUsername") or ctx.executing_user.get("Oid") or "unknown"
    except Exception:
        return "unknown"


def _parse_date(value: str) -> datetime.date | None:
    """Read an optional ISO follow-up date; blank or None means no planned follow-up."""
    text = (value or "").strip()
    if text in ("", "None"):
        return None
    try:
        return datetime.date.fromisoformat(text[:10])
    except ValueError:
        raise fn.UserThrownError(f"'{value}' is not a valid follow-up date. Use YYYY-MM-DD.")


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.context(argName="ctx")
@udf.function()
def update_outreach(sqlDB: fn.FabricSqlConnection, ctx: fn.UserDataFunctionContext,
                    careGapId: int, outreachStatus: str, priority: str, nextFollowUpDate: str) -> str:
    """Update the selected gap with the report status, priority and optional follow-up.
    Closed statuses clear the date; this demo has no scheduling restrictions.
    The operation also records the old and new values in the SQL audit log."""
    # --- 1. Validate the request before touching the database -----------------------------
    if outreachStatus not in OPEN_STATUSES + CLOSED_STATUSES:
        raise fn.UserThrownError(f"Outreach status must be one of: {', '.join(OPEN_STATUSES + CLOSED_STATUSES)}.")
    if priority not in PRIORITIES:
        raise fn.UserThrownError("Priority must be High, Medium or Low.")

    follow_up = None if outreachStatus in CLOSED_STATUSES else _parse_date(nextFollowUpDate)

    user = _user(ctx)
    conn = sqlDB.connect()
    try:
        cur = conn.cursor()

        # --- 2. Read the current values (for the confirmation message and the audit log) ---
        cur.execute("SELECT PatientName, QualityMeasure, OutreachStatus, Priority, NextFollowUpDate "
                    "FROM dbo.CareGap WHERE CareGapId = ?", careGapId)
        row = cur.fetchone()
        if row is None:
            raise fn.UserThrownError(f"Care gap {careGapId} was not found.")
        patient, measure, old_status, old_priority, old_fu = row
        old_fu_text = old_fu.isoformat() if old_fu else "none"
        new_fu_text = follow_up.isoformat() if follow_up else "none"

        # --- 3. Write the change and its audit entry (parameterized SQL, one transaction) ---
        cur.execute(
            "UPDATE dbo.CareGap SET OutreachStatus = ?, Priority = ?, NextFollowUpDate = ?, "
            "LastUpdatedBy = ?, LastUpdatedAt = SYSUTCDATETIME() WHERE CareGapId = ?",
            outreachStatus, priority, follow_up, user, careGapId)
        cur.execute(
            "INSERT INTO dbo.WriteBackLog (CareGapId, TargetSystem, ActionName, OldValue, NewValue, ChangedBy) "
            "VALUES (?, 'Fabric SQL DB', 'Update Outreach', ?, ?, ?)",
            careGapId,
            f"{old_status} | {old_priority} | follow-up {old_fu_text}",
            f"{outreachStatus} | {priority} | follow-up {new_fu_text}",
            user)
        conn.commit()
        cur.close()
    finally:
        # Close the connection on every path; an exception before commit leaves nothing saved.
        conn.close()

    # --- 4. Confirmation shown to the report user ----------------------------------------
    suffix = " (gap closed, follow-up cleared)" if outreachStatus in CLOSED_STATUSES else f", follow-up {new_fu_text}"
    return f"{patient} - {measure}: {old_status} -> {outreachStatus}, priority {priority}{suffix}."


def _create_care_gap(sqlDB: fn.FabricSqlConnection, ctx: fn.UserDataFunctionContext,
                     patientName: str, clinic: str, qualityMeasure: str, dueDate: str,
                     priority: str, patient_id: str | None = None,
                     outreach_status: str = "Not Started",
                     follow_up: datetime.date | None = None) -> str:
    """Write the validated insert fields to the source, then record the audit entry.
    This is a private helper, not a callable UDF. SQL-only fields use synthetic demo values."""
    # --- 1. Validate the request ------------------------------------------------------------
    name = (patientName or "").strip()
    if not name:
        raise fn.UserThrownError("Please type a patient name before creating a care gap.")
    if len(name) > 100:  # matches NVARCHAR(100) on dbo.CareGap.PatientName
        raise fn.UserThrownError("Patient name is limited to 100 characters.")
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

    # --- 2. Auto-filled demographic/plan values (not collected by the "quick add" form) -----
    coordinator, provider = CLINIC_STAFF.get(clinic, ("Unassigned", "Unassigned"))
    measure_code = MEASURE_CODES.get(quality_measure, "OTH")
    age = random.randint(45, 82)
    sex = random.choice(["F", "M"])
    payer_type = random.choice(PAYER_TYPES)
    risk_score = round(random.uniform(0.6, 2.2), 2)
    gap_opened = datetime.date.today()
    user = _user(ctx)

    conn = sqlDB.connect()
    try:
        cur = conn.cursor()

        cur.execute("SELECT NEXT VALUE FOR dbo.CareGapIdSequence")
        new_id = cur.fetchone()[0]
        patient_id = patient_id or f"P-2{new_id}"
        cur.execute(
            "INSERT INTO dbo.CareGap (CareGapId, PatientId, PatientName, Age, Sex, PayerType, Clinic, "
            "PrimaryCareProvider, CareCoordinator, QualityMeasure, MeasureCode, RiskScore, GapOpenedDate, "
            "DueDate, OutreachStatus, Priority, NextFollowUpDate, LastUpdatedBy, LastUpdatedAt) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, SYSUTCDATETIME())",
            new_id, patient_id, name, age, sex, payer_type, clinic, provider, coordinator,
            quality_measure, measure_code, risk_score, gap_opened, due_date, outreach_status, priority, follow_up, user)

        # --- 4. Audit (OldValue is NULL - there is no "before" state for a brand-new row) ---
        cur.execute(
            "INSERT INTO dbo.WriteBackLog (CareGapId, TargetSystem, ActionName, OldValue, NewValue, ChangedBy) "
            "VALUES (?, 'Fabric SQL DB', 'Create Care Gap', NULL, ?, ?)",
            new_id, (f"{patient_id} | {name} | {clinic} | {quality_measure} | due {due_date.isoformat()} | "
                     f"{outreach_status} | {priority} | follow-up {follow_up.isoformat() if follow_up else 'none'}")[:400], user)
        conn.commit()
        cur.close()
    finally:
        conn.close()

    return f"Created care gap #{new_id} for {name} - {quality_measure} at {clinic}."


def _patient_id(value: str) -> str:
    """Validate the 10-character SQL field and exclude spreadsheet formula prefixes."""
    value = (value or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,9}", value):
        raise fn.UserThrownError("Patient ID must be 1-10 letters, digits, hyphens or underscores, starting with a letter or digit.")
    return value






def _insert_values(patientName: str, clinic: str, qualityMeasure: str, dueDate: str,
                   priority: str, outreachStatus: str, nextFollowUpDate: str):
    """Check source-compatible fields and parse dates; no demo scheduling rules are enforced."""
    name, clinic, measure = (patientName or "").strip(), (clinic or "").strip(), (qualityMeasure or "").strip()
    if not name or len(name) > 100:
        raise fn.UserThrownError("Enter a patient name between 1 and 100 characters.")
    if clinic not in CLINIC_STAFF:
        raise fn.UserThrownError("Select one of the demo clinics.")
    if measure not in MEASURE_CODES:
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
@udf.context(argName="ctx")
@udf.function()
def insert_care_gap(sqlDB: fn.FabricSqlConnection, ctx: fn.UserDataFunctionContext,
                    patientId: str, patientName: str, clinic: str, qualityMeasure: str, dueDate: str,
                    priority: str, outreachStatus: str, nextFollowUpDate: str) -> str:
    """Insert the eight report fields as a new care gap and record an audit entry.
    Patient IDs may repeat across gaps, but each gap receives a unique non-reused ID.
    Follow-up is optional; closed statuses clear it. No patient-master lookup is performed."""
    patient_id = _patient_id(patientId)
    name, clinic, measure, due, follow_up = _insert_values(
        patientName, clinic, qualityMeasure, dueDate, priority, outreachStatus, nextFollowUpDate)
    return _create_care_gap(sqlDB, ctx, name, clinic, measure, due, priority,
                            patient_id, outreachStatus, follow_up)


@udf.connection(argName="sqlDB", alias="TTFCareGapDB")
@udf.context(argName="ctx")
@udf.function()
def delete_care_gap(sqlDB: fn.FabricSqlConnection, ctx: fn.UserDataFunctionContext,
                    careGapId: int) -> str:
    """Permanently delete the single gap selected in the report's main worklist.

    The button passes SELECTEDVALUE(Care Gap Id); blank/multiple selections disable it.
    No separately typed confirmation is required. Lock the gap while removing its
    child notes, deleting the parent and inserting a Delete Care Gap audit entry. Existing
    audit rows are never deleted. All three writes commit atomically; errors roll back.
    """
    if not isinstance(careGapId, int) or isinstance(careGapId, bool) or careGapId <= 0:
        raise fn.UserThrownError("Select exactly one care gap in the main worklist before deleting.")
    user = _user(ctx)
    conn = sqlDB.connect()
    try:
        cur = conn.cursor()
        cur.execute("SELECT PatientId, PatientName, QualityMeasure FROM dbo.CareGap WITH (UPDLOCK, HOLDLOCK) "
                    "WHERE CareGapId = ?", careGapId)
        row = cur.fetchone()
        if row is None:
            raise fn.UserThrownError(f"Care gap {careGapId} was not found. Refresh the worklist.")
        cur.execute("DELETE FROM dbo.OutreachNote WHERE CareGapId = ?", careGapId)
        cur.execute("DELETE FROM dbo.CareGap WHERE CareGapId = ?", careGapId)
        cur.execute(
            "INSERT dbo.WriteBackLog (CareGapId, TargetSystem, ActionName, OldValue, NewValue, ChangedBy) "
            "VALUES (?, 'Fabric SQL DB', 'Delete Care Gap', ?, 'Permanently deleted (including outreach notes)', ?)",
            careGapId, f"{row[0]} | {row[1]} | {row[2]}"[:400], user)
        conn.commit()
        cur.close()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return f"Deleted SQL care gap #{careGapId} for Patient ID {row[0]}. Audit history retained."
