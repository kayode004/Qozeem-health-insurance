# ============================================================
#  Qozeem Automated Healthcare Insurance Claims Processing
#  and Patient Management Record (AHICPMR)
#  © 2024 Qozeem Health Technologies
# ============================================================

import streamlit as st
import pandas as pd
import hashlib
import os
import uuid
import smtplib
from datetime import datetime, date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Qozeem AHICPMR",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────
# FILE PATHS
# ─────────────────────────────────────────────────────────────
CLAIMS_FILE         = "claims.csv"
PATIENTS_FILE       = "patients.csv"
USERS_FILE          = "users.csv"
AUDIT_FILE          = "audit_log.csv"
NOTIFICATIONS_FILE  = "notifications.csv"

PLANS       = ["Basic", "Standard", "Premium", "Executive"]
MONTH_NAMES = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]
PLAN_LIMITS = {"Basic":30_000,"Standard":150_000,"Premium":350_000,"Executive":500_000}
PLAN_PREMIUMS = {"Basic":5_000,"Standard":12_000,"Premium":25_000,"Executive":50_000}

# ─────────────────────────────────────────────────────────────
# UTILITY
# ─────────────────────────────────────────────────────────────
def hp(pwd: str) -> str:
    return hashlib.sha256(pwd.encode()).hexdigest()

def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def today() -> str:
    return date.today().strftime("%Y-%m-%d")

def gen_id(prefix: str) -> str:
    return f"{prefix}{str(uuid.uuid4())[:8].upper()}"

def load(file: str) -> pd.DataFrame:
    return pd.read_csv(file) if os.path.exists(file) else pd.DataFrame()

def save(df: pd.DataFrame, file: str):
    df.to_csv(file, index=False)

def to_excel(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Report")
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────
# DATA INITIALISATION — runs once on startup
# ─────────────────────────────────────────────────────────────
def init_data():
    if not os.path.exists(USERS_FILE):
        pd.DataFrame([
            {"username":"hospital_admin",  "password_hash":hp("hospital123"),"role":"hospital_admin",
             "organization":"Lagos General Hospital","email":"hospital@qozeem.com","created_at":now()},
            {"username":"hospital_officer","password_hash":hp("officer123"),  "role":"hospital_officer",
             "organization":"Lagos General Hospital","email":"officer@qozeem.com","created_at":now()},
            {"username":"hmo_admin",       "password_hash":hp("hmo123"),      "role":"hmo_admin",
             "organization":"Qozeem Health HMO","email":"hmo@qozeem.com","created_at":now()},
            {"username":"hmo_officer",     "password_hash":hp("hmooff123"),   "role":"hmo_officer",
             "organization":"Qozeem Health HMO","email":"hmooff@qozeem.com","created_at":now()},
        ]).to_csv(USERS_FILE, index=False)

    if not os.path.exists(PATIENTS_FILE):
        pd.DataFrame([
            {"patient_id":"PAT001","name":"John Doe",          "dob":"1985-03-15","gender":"Male",
             "phone":"08012345678","email":"john@mail.com",   "address":"12 Lagos St",
             "policy_number":"POL001","plan":"Basic",    "hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-01-01","expiry_date":"2024-12-31","status":"Active"},
            {"patient_id":"PAT002","name":"Jane Smith",         "dob":"1990-07-22","gender":"Female",
             "phone":"08023456789","email":"jane@mail.com",   "address":"45 Abuja Ave",
             "policy_number":"POL002","plan":"Premium",  "hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-01-01","expiry_date":"2024-12-31","status":"Active"},
            {"patient_id":"PAT003","name":"Ali Bello",          "dob":"1978-11-08","gender":"Male",
             "phone":"08034567890","email":"ali@mail.com",    "address":"7 Kano Rd",
             "policy_number":"POL003","plan":"Basic",    "hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-02-01","expiry_date":"2025-01-31","status":"Active"},
            {"patient_id":"PAT004","name":"Ngozi Eze",          "dob":"1995-05-30","gender":"Female",
             "phone":"08045678901","email":"ngozi@mail.com",  "address":"3 Enugu Cl",
             "policy_number":"POL004","plan":"Standard", "hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-03-01","expiry_date":"2025-02-28","status":"Active"},
            {"patient_id":"PAT005","name":"Emeka Obi",          "dob":"1982-09-14","gender":"Male",
             "phone":"08056789012","email":"emeka@mail.com",  "address":"21 PH Blvd",
             "policy_number":"POL005","plan":"Executive","hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-01-15","expiry_date":"2025-01-14","status":"Active"},
            {"patient_id":"PAT006","name":"Fatima Yusuf",       "dob":"1993-12-01","gender":"Female",
             "phone":"08067891234","email":"fatima@mail.com", "address":"9 Kaduna Rd",
             "policy_number":"POL006","plan":"Standard", "hospital":"Lagos General Hospital",
             "hmo":"Qozeem Health HMO","enrollment_date":"2024-04-01","expiry_date":"2025-03-31","status":"Active"},
        ]).to_csv(PATIENTS_FILE, index=False)

    if not os.path.exists(CLAIMS_FILE):
        pd.DataFrame([
            {"claim_id":"CLM001","policy_number":"POL001","patient_name":"John Doe",
             "diagnosis":"Malaria","treatment":"Antimalarial + IV drip","claim_amount":25000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-01-15","date_processed":"2024-01-17",
             "status":"APPROVED","decision_reason":"Within Basic plan limit","processed_by":"hmo_admin",
             "month":"January","year":2024},
            {"claim_id":"CLM002","policy_number":"POL002","patient_name":"Jane Smith",
             "diagnosis":"Appendicitis","treatment":"Appendectomy surgery","claim_amount":600000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-02-10","date_processed":"2024-02-12",
             "status":"REJECTED","decision_reason":"Exceeds maximum limit of ₦500,000","processed_by":"hmo_admin",
             "month":"February","year":2024},
            {"claim_id":"CLM003","policy_number":"POL003","patient_name":"Ali Bello",
             "diagnosis":"Hypertension","treatment":"Medication + monitoring","claim_amount":35000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-03-05","date_processed":"2024-03-06",
             "status":"PARTIALLY APPROVED","decision_reason":"Basic plan cap ₦30,000 applied","processed_by":"hmo_admin",
             "month":"March","year":2024},
            {"claim_id":"CLM004","policy_number":"POL004","patient_name":"Ngozi Eze",
             "diagnosis":"Diabetes management","treatment":"Insulin + lab tests","claim_amount":80000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-04-20","date_processed":"",
             "status":"PENDING","decision_reason":"","processed_by":"",
             "month":"April","year":2024},
            {"claim_id":"CLM005","policy_number":"POL005","patient_name":"Emeka Obi",
             "diagnosis":"Chest pain evaluation","treatment":"ECG, Echo, Stress test","claim_amount":150000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-05-08","date_processed":"",
             "status":"PENDING","decision_reason":"","processed_by":"",
             "month":"May","year":2024},
            {"claim_id":"CLM006","policy_number":"POL006","patient_name":"Fatima Yusuf",
             "diagnosis":"Typhoid fever","treatment":"IV antibiotics + fluids","claim_amount":45000,
             "hospital":"Lagos General Hospital","hmo":"Qozeem Health HMO",
             "date_submitted":"2024-06-03","date_processed":"",
             "status":"PENDING","decision_reason":"","processed_by":"",
             "month":"June","year":2024},
        ]).to_csv(CLAIMS_FILE, index=False)

    for fpath, cols in [
        (AUDIT_FILE,         ["log_id","timestamp","user","role","action","details"]),
        (NOTIFICATIONS_FILE, ["notif_id","recipient_org","recipient_role","message","type","claim_id","timestamp","read"]),
    ]:
        if not os.path.exists(fpath):
            pd.DataFrame(columns=cols).to_csv(fpath, index=False)


# ─────────────────────────────────────────────────────────────
# AUDIT LOG
# ─────────────────────────────────────────────────────────────
def log_action(user: str, role: str, action: str, details: str = ""):
    df = load(AUDIT_FILE)
    row = pd.DataFrame([{"log_id":gen_id("LOG"),"timestamp":now(),
                         "user":user,"role":role,"action":action,"details":details}])
    df = pd.concat([df, row], ignore_index=True)
    save(df, AUDIT_FILE)

# ─────────────────────────────────────────────────────────────
# NOTIFICATIONS
# ─────────────────────────────────────────────────────────────
def add_notification(recipient_org: str, recipient_role: str,
                     message: str, notif_type: str, claim_id: str = ""):
    df = load(NOTIFICATIONS_FILE)
    row = pd.DataFrame([{"notif_id":gen_id("NTF"),"recipient_org":recipient_org,
                         "recipient_role":recipient_role,"message":message,
                         "type":notif_type,"claim_id":claim_id,"timestamp":now(),"read":False}])
    df = pd.concat([df, row], ignore_index=True)
    save(df, NOTIFICATIONS_FILE)

def get_notifications(org: str, role: str) -> pd.DataFrame:
    df = load(NOTIFICATIONS_FILE)
    if df.empty:
        return df
    mask = (df["recipient_org"] == org) & (
        (df["recipient_role"] == role) | (df["recipient_role"] == "all")
    )
    return df[mask].sort_values("timestamp", ascending=False)

def count_unread(org: str, role: str) -> int:
    df = get_notifications(org, role)
    if df.empty:
        return 0
    return int((df["read"].astype(str) == "False").sum())

def mark_all_read(org: str, role: str):
    df = load(NOTIFICATIONS_FILE)
    if df.empty:
        return
    mask = (df["recipient_org"] == org) & (
        (df["recipient_role"] == role) | (df["recipient_role"] == "all")
    )
    df.loc[mask, "read"] = True
    save(df, NOTIFICATIONS_FILE)

# ─────────────────────────────────────────────────────────────
# EMAIL
# ─────────────────────────────────────────────────────────────
def send_email(to_email: str, subject: str, html_body: str) -> bool:
    try:
        cfg = st.secrets.get("email", {})
        sender   = cfg.get("sender", "")
        password = cfg.get("password", "")
        if not sender or not password:
            return False
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"]    = sender
        msg["To"]      = to_email
        msg.attach(MIMEText(html_body, "html"))
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(sender, password)
            s.sendmail(sender, to_email, msg.as_string())
        return True
    except Exception:
        return False

def notify_hmo_by_email(hmo_org: str, claim_id: str, hospital: str,
                         patient: dict, diagnosis: str, treatment: str, amount: float):
    users = load(USERS_FILE)
    if users.empty:
        return
    admins = users[(users["organization"] == hmo_org) & (users["role"] == "hmo_admin")]
    body = f"""
    <h2 style="color:#0066cc;">🏥 New Insurance Claim Received</h2>
    <table border="1" cellpadding="8" cellspacing="0">
      <tr><td><b>Claim ID</b></td><td>{claim_id}</td></tr>
      <tr><td><b>Hospital</b></td><td>{hospital}</td></tr>
      <tr><td><b>Patient</b></td><td>{patient['name']}</td></tr>
      <tr><td><b>Policy</b></td><td>{patient['policy_number']} ({patient['plan']} Plan)</td></tr>
      <tr><td><b>Diagnosis</b></td><td>{diagnosis}</td></tr>
      <tr><td><b>Treatment</b></td><td>{treatment}</td></tr>
      <tr><td><b>Amount</b></td><td>₦{amount:,.0f}</td></tr>
      <tr><td><b>Submitted</b></td><td>{today()}</td></tr>
    </table>
    <br><p>Please log in to <b>Qozeem AHICPMR</b> to review and process this claim.</p>
    """
    for _, u in admins.iterrows():
        send_email(u["email"], f"[Qozeem AHICPMR] New Claim {claim_id} from {hospital}", body)

def notify_hospital_by_email(hospital_org: str, claim_id: str, patient_name: str,
                              decision: str, reason: str, processed_by: str):
    users = load(USERS_FILE)
    if users.empty:
        return
    admins = users[(users["organization"] == hospital_org) & (users["role"] == "hospital_admin")]
    icon_map = {"APPROVED":"✅","REJECTED":"❌","PARTIALLY APPROVED":"🟡","UNDER REVIEW":"🔍"}
    icon = icon_map.get(decision, "📋")
    color_map = {"APPROVED":"#28a745","REJECTED":"#dc3545","PARTIALLY APPROVED":"#ffc107","UNDER REVIEW":"#17a2b8"}
    color = color_map.get(decision, "#333")
    body = f"""
    <h2 style="color:{color};">{icon} Insurance Claim Decision — {decision}</h2>
    <table border="1" cellpadding="8" cellspacing="0">
      <tr><td><b>Claim ID</b></td><td>{claim_id}</td></tr>
      <tr><td><b>Patient</b></td><td>{patient_name}</td></tr>
      <tr><td><b>Decision</b></td><td style="color:{color};"><b>{decision}</b></td></tr>
      <tr><td><b>Reason</b></td><td>{reason}</td></tr>
      <tr><td><b>Processed By</b></td><td>{processed_by}</td></tr>
      <tr><td><b>Date</b></td><td>{today()}</td></tr>
    </table>
    <br><p>Log in to <b>Qozeem AHICPMR</b> to view full claim details.</p>
    """
    for _, u in admins.iterrows():
        send_email(u["email"], f"[Qozeem AHICPMR] Claim {claim_id} — {decision}", body)

# ─────────────────────────────────────────────────────────────
# CLAIM ENGINE
# ─────────────────────────────────────────────────────────────
def auto_process_claim(amount: float, plan: str, claims_count: int):
    if amount > 500_000:
        return "REJECTED", "Claim exceeds the maximum allowable limit of ₦500,000"
    if claims_count > 3:
        return "REJECTED", f"Patient has exceeded the maximum claim count (3) for this period"
    cap = PLAN_LIMITS.get(plan, 30_000)
    if amount > cap:
        return "PARTIALLY APPROVED", f"{plan} plan cap of ₦{cap:,} applied — excess not covered"
    return "APPROVED", "Claim meets all requirements and is within plan limits"

# ─────────────────────────────────────────────────────────────
# RISK SCORING
# ─────────────────────────────────────────────────────────────
def risk_score(claims_count: int, total_claimed: float) -> str:
    if claims_count > 3 or total_claimed > 400_000:
        return "🔴 High"
    if claims_count > 1 or total_claimed > 150_000:
        return "🟡 Medium"
    return "🟢 Low"

# ─────────────────────────────────────────────────────────────
# AUTHENTICATION
# ─────────────────────────────────────────────────────────────
def authenticate(username: str, password: str):
    users = load(USERS_FILE)
    if users.empty:
        return None
    match = users[users["username"] == username]
    if match.empty:
        return None
    u = match.iloc[0]
    return u if u["password_hash"] == hp(password) else None

def login_page():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='text-align:center;padding:30px 0 10px'>
          <h1>🏥 Qozeem AHICPMR</h1>
          <p style='color:#555;font-size:15px'>
            Automated Healthcare Insurance Claims Processing<br>& Patient Management Record
          </p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")
        with st.form("login_form"):
            username  = st.text_input("👤 Username")
            password  = st.text_input("🔑 Password", type="password")
            submitted = st.form_submit_button("Login", use_container_width=True)
        if submitted:
            user = authenticate(username, password)
            if user is not None:
                st.session_state["user"]  = user["username"]
                st.session_state["role"]  = user["role"]
                st.session_state["org"]   = user["organization"]
                st.session_state["email"] = user["email"]
                log_action(username, user["role"], "LOGIN",
                           f"Logged in from {user['organization']}")
                st.rerun()
            else:
                st.error("❌ Invalid credentials. Please try again.")
        st.markdown("---")
        st.markdown("""
**Demo Credentials**

| Role | Username | Password |
|---|---|---|
| Hospital Admin | `hospital_admin` | `hospital123` |
| Hospital Officer | `hospital_officer` | `officer123` |
| HMO Admin | `hmo_admin` | `hmo123` |
| HMO Officer | `hmo_officer` | `hmooff123` |
        """)

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
def sidebar() -> str:
    with st.sidebar:
        st.markdown("### 🏥 Qozeem AHICPMR")
        st.markdown("---")
        role = st.session_state["role"]
        org  = st.session_state["org"]
        user = st.session_state["user"]

        st.write(f"👤 **{user}**")
        st.write(f"🏢 {org}")
        st.write(f"🔖 `{role.replace('_', ' ').title()}`")

        unread = count_unread(org, role)
        if unread > 0:
            st.error(f"🔔 {unread} unread notification(s)")
        st.markdown("---")

        if "hospital" in role:
            pages = ["📊 Dashboard","🧾 Submit Claim","📋 My Claims",
                     "👥 Patients","🔔 Notifications","📥 Download Reports"]
        else:
            pages = ["📊 Dashboard","📨 Incoming Claims","✅ Process Claims",
                     "📈 Analytics","💰 Billing & Accounting","⚠️ Risk Scoring",
                     "📋 Audit Log","🔔 Notifications","📥 Download Reports"]
            if role == "hmo_admin":
                pages.append("👤 User Management")

        page = st.radio("Navigate", pages, label_visibility="collapsed")
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            log_action(user, role, "LOGOUT", "")
            st.session_state.clear()
            st.rerun()
    return page


# ═══════════════════════════════════════════════════════════════
#  ████  HOSPITAL PORTAL  ████
# ═══════════════════════════════════════════════════════════════

def hospital_dashboard():
    org = st.session_state["org"]
    st.title("📊 Hospital Dashboard")
    st.caption(f"🏢 {org}")
    st.markdown("---")

    claims   = load(CLAIMS_FILE)
    patients = load(PATIENTS_FILE)

    my_claims   = claims[claims["hospital"] == org] if not claims.empty else pd.DataFrame()
    my_patients = patients[patients["hospital"] == org] if not patients.empty else pd.DataFrame()

    total     = len(my_claims)
    approved  = len(my_claims[my_claims["status"] == "APPROVED"])        if not my_claims.empty else 0
    partial   = len(my_claims[my_claims["status"] == "PARTIALLY APPROVED"]) if not my_claims.empty else 0
    rejected  = len(my_claims[my_claims["status"] == "REJECTED"])        if not my_claims.empty else 0
    pending   = len(my_claims[my_claims["status"] == "PENDING"])         if not my_claims.empty else 0
    total_val = my_claims["claim_amount"].sum()                           if not my_claims.empty else 0
    n_patients = len(my_patients)

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Patients",       n_patients)
    c2.metric("Total Claims",   total)
    c3.metric("✅ Approved",    approved)
    c4.metric("🟡 Partial",    partial)
    c5.metric("❌ Rejected",   rejected)
    c6.metric("⏳ Pending",    pending)
    c7.metric("💰 Total (₦)", f"{total_val:,.0f}")

    st.markdown("---")

    if not my_claims.empty:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Claims Status Distribution")
            sc = my_claims["status"].value_counts().reset_index()
            sc.columns = ["Status","Count"]
            colors = {"APPROVED":"#28a745","REJECTED":"#dc3545",
                      "PARTIALLY APPROVED":"#ffc107","PENDING":"#6c757d","UNDER REVIEW":"#17a2b8"}
            fig = px.pie(sc, names="Status", values="Count",
                         color="Status", color_discrete_map=colors)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.subheader("Claim Amount by Month")
            if "month" in my_claims.columns:
                monthly = my_claims.groupby("month")["claim_amount"].sum().reset_index()
                monthly.columns = ["Month","Total Amount (₦)"]
                fig2 = px.bar(monthly, x="Month", y="Total Amount (₦)",
                              color_discrete_sequence=["#007bff"])
                st.plotly_chart(fig2, use_container_width=True)

        st.subheader("📋 Recent Claims")
        st.dataframe(my_claims.tail(10), use_container_width=True)
    else:
        st.info("No claims submitted yet. Use 'Submit Claim' to send your first claim.")


def submit_claim():
    st.title("🧾 Submit New Claim to HMO")
    org = st.session_state["org"]
    st.markdown("---")

    patients = load(PATIENTS_FILE)
    if patients.empty:
        st.warning("No patients registered. Please register patients first.")
        return

    my_patients = patients[patients["hospital"] == org]
    if my_patients.empty:
        st.warning("No patients found for your hospital.")
        return

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Patient Information")
        selected_name = st.selectbox("Select Patient", my_patients["name"].tolist())
        patient = my_patients[my_patients["name"] == selected_name].iloc[0]
        st.info(f"""
**Policy Number:** {patient['policy_number']}
**Insurance Plan:** {patient['plan']}
**HMO:** {patient['hmo']}
**Enrollment:** {patient['enrollment_date']}
**Expiry:** {patient['expiry_date']}
**Status:** {patient['status']}
        """)
        plan_cap = PLAN_LIMITS.get(patient["plan"], 30_000)
        st.warning(f"📌 Plan Cap: ₦{plan_cap:,}")

    with col2:
        st.subheader("Claim Details")
        diagnosis = st.text_input("Diagnosis / Condition *")
        treatment = st.text_area("Treatment / Services Rendered *", height=100)
        amount    = st.number_input("Claim Amount (₦) *", min_value=0, step=500, format="%d")
        month     = st.selectbox("Month of Service", MONTH_NAMES, index=datetime.now().month - 1)
        year      = st.number_input("Year", min_value=2020, max_value=2030,
                                    value=datetime.now().year, step=1)
        notes     = st.text_area("Additional Notes", height=70)

    st.markdown("---")
    if st.button("📤 Submit Claim to HMO", type="primary", use_container_width=True):
        if not diagnosis.strip() or not treatment.strip() or amount == 0:
            st.error("Please fill in all required fields (Diagnosis, Treatment, Amount).")
            return

        existing = load(CLAIMS_FILE)
        count = len(existing[existing["policy_number"] == patient["policy_number"]]) \
                if not existing.empty else 0

        claim_id = gen_id("CLM")
        new_claim = pd.DataFrame([{
            "claim_id":       claim_id,
            "policy_number":  patient["policy_number"],
            "patient_name":   patient["name"],
            "diagnosis":      diagnosis,
            "treatment":      treatment,
            "claim_amount":   amount,
            "hospital":       org,
            "hmo":            patient["hmo"],
            "date_submitted": today(),
            "date_processed": "",
            "status":         "PENDING",
            "decision_reason":"",
            "processed_by":   "",
            "month":          month,
            "year":           year,
            "notes":          notes,
        }])
        df = pd.concat([existing, new_claim], ignore_index=True) if not existing.empty else new_claim
        save(df, CLAIMS_FILE)

        # In-app notification → HMO
        add_notification(
            recipient_org=patient["hmo"],
            recipient_role="hmo_admin",
            message=(f"🏥 New claim {claim_id} from {org} for {patient['name']} "
                     f"— ₦{amount:,} | {patient['plan']} Plan | {diagnosis}"),
            notif_type="NEW_CLAIM",
            claim_id=claim_id
        )

        # Email notification → HMO admins
        notify_hmo_by_email(patient["hmo"], claim_id, org, patient.to_dict(),
                            diagnosis, treatment, amount)

        log_action(st.session_state["user"], st.session_state["role"],
                   "SUBMIT_CLAIM",
                   f"Claim {claim_id} for {patient['name']} — ₦{amount:,}")

        st.success(f"✅ Claim **{claim_id}** submitted to **{patient['hmo']}** successfully!")
        st.info("🔔 The HMO has been notified in-app and by email. Await their decision.")
        st.balloons()


def hospital_my_claims():
    st.title("📋 My Submitted Claims")
    org = st.session_state["org"]
    st.markdown("---")

    df = load(CLAIMS_FILE)
    if df.empty:
        st.info("No claims submitted yet.")
        return

    my = df[df["hospital"] == org].copy()

    col1, col2, col3 = st.columns(3)
    with col1:
        sf = st.selectbox("Status", ["All","PENDING","APPROVED","PARTIALLY APPROVED","REJECTED","UNDER REVIEW"])
    with col2:
        mf = st.selectbox("Month", ["All"] + MONTH_NAMES)
    with col3:
        search = st.text_input("🔍 Search Patient Name")

    if sf != "All":
        my = my[my["status"] == sf]
    if mf != "All":
        my = my[my["month"] == mf]
    if search:
        my = my[my["patient_name"].str.contains(search, case=False, na=False)]

    st.metric("Showing", len(my))
    st.dataframe(my, use_container_width=True)

    if not my.empty:
        st.download_button(
            "📥 Export to Excel", to_excel(my), "my_claims.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def hospital_patients():
    st.title("👥 Patient Management")
    org = st.session_state["org"]
    st.markdown("---")

    tab1, tab2 = st.tabs(["📋 All Patients", "➕ Register New Patient"])

    with tab1:
        df = load(PATIENTS_FILE)
        if not df.empty:
            my = df[df["hospital"] == org]
            st.metric("Enrolled Patients", len(my))
            expiring = my[my["expiry_date"] <= today()] if "expiry_date" in my.columns else pd.DataFrame()
            if not expiring.empty:
                st.warning(f"⚠️ {len(expiring)} policy/policies have expired or are expiring.")
            st.dataframe(my, use_container_width=True)
            if not my.empty:
                st.download_button("📥 Export Patient List", to_excel(my), "patients.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No patients registered yet.")

    with tab2:
        if st.session_state["role"] != "hospital_admin":
            st.warning("Only Hospital Admins can register new patients.")
            return
        col1, col2 = st.columns(2)
        with col1:
            name   = st.text_input("Full Name *")
            dob    = st.date_input("Date of Birth", min_value=date(1900,1,1))
            gender = st.selectbox("Gender", ["Male","Female","Other"])
            phone  = st.text_input("Phone Number *")
            email  = st.text_input("Email Address")
            address = st.text_area("Home Address")
        with col2:
            policy_num = st.text_input("Policy Number *")
            plan       = st.selectbox("Insurance Plan *", PLANS)
            hmo_name   = st.text_input("HMO Organization *", value="Qozeem Health HMO")
            enroll     = st.date_input("Enrollment Date")
            expiry     = st.date_input("Policy Expiry Date")

        if st.button("✅ Register Patient", type="primary", use_container_width=True):
            if not name.strip() or not phone.strip() or not policy_num.strip():
                st.error("Name, Phone, and Policy Number are required.")
                return
            existing = load(PATIENTS_FILE)
            if not existing.empty and policy_num in existing["policy_number"].values:
                st.error(f"Policy number **{policy_num}** already exists.")
                return
            pat_id = gen_id("PAT")
            new_pat = pd.DataFrame([{
                "patient_id":pat_id,"name":name,"dob":str(dob),"gender":gender,
                "phone":phone,"email":email,"address":address,
                "policy_number":policy_num,"plan":plan,"hospital":org,"hmo":hmo_name,
                "enrollment_date":str(enroll),"expiry_date":str(expiry),"status":"Active"
            }])
            df = pd.concat([existing, new_pat], ignore_index=True) if not existing.empty else new_pat
            save(df, PATIENTS_FILE)
            log_action(st.session_state["user"], st.session_state["role"],
                       "REGISTER_PATIENT", f"Registered {name} ({pat_id})")
            st.success(f"✅ Patient **{name}** registered with ID **{pat_id}**")
            st.rerun()


def hospital_notifications():
    st.title("🔔 Notifications from HMO")
    org  = st.session_state["org"]
    role = st.session_state["role"]
    st.markdown("---")

    notifs = get_notifications(org, role)
    if notifs.empty:
        st.info("No notifications yet. Notifications from the HMO will appear here.")
        return

    col1, col2 = st.columns([3,1])
    with col1:
        st.metric("Total Notifications", len(notifs))
    with col2:
        if st.button("✅ Mark All as Read", use_container_width=True):
            mark_all_read(org, role)
            st.rerun()

    for _, n in notifs.iterrows():
        is_read = str(n["read"]) == "True"
        icon = {"NEW_CLAIM":"🧾","APPROVED":"✅","REJECTED":"❌",
                "PARTIALLY_APPROVED":"🟡","UNDER_REVIEW":"🔍"}.get(str(n["type"]),"🔔")
        msg = f"{icon} {n['message']}  \n🕐 *{n['timestamp']}*"
        if not is_read:
            st.success(msg)
        else:
            st.info(msg)


def hospital_download():
    st.title("📥 Download Reports")
    org = st.session_state["org"]
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Claims Report")
        df = load(CLAIMS_FILE)
        if not df.empty:
            my = df[df["hospital"] == org]
            st.download_button("📥 All My Claims (Excel)", to_excel(my), "claims_report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col2:
        st.subheader("Patients Report")
        patients = load(PATIENTS_FILE)
        if not patients.empty:
            my_p = patients[patients["hospital"] == org]
            st.download_button("📥 Patient List (Excel)", to_excel(my_p), "patients_report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═══════════════════════════════════════════════════════════════
#  ████  HMO PORTAL  ████
# ═══════════════════════════════════════════════════════════════

def hmo_dashboard():
    org = st.session_state["org"]
    st.title("📊 HMO Dashboard")
    st.caption(f"🏢 {org}")
    st.markdown("---")

    df       = load(CLAIMS_FILE)
    patients = load(PATIENTS_FILE)

    if df.empty:
        st.info("No claims data yet.")
        return

    total     = len(df)
    approved  = len(df[df["status"] == "APPROVED"])
    partial   = len(df[df["status"] == "PARTIALLY APPROVED"])
    rejected  = len(df[df["status"] == "REJECTED"])
    pending   = len(df[df["status"] == "PENDING"])
    total_val = df["claim_amount"].sum()
    approved_val = df[df["status"] == "APPROVED"]["claim_amount"].sum()

    c1,c2,c3,c4,c5,c6,c7 = st.columns(7)
    c1.metric("Total Claims",   total)
    c2.metric("✅ Approved",   approved)
    c3.metric("🟡 Partial",   partial)
    c4.metric("❌ Rejected",  rejected)
    c5.metric("⏳ Pending",   pending)
    c6.metric("💰 Total (₦)", f"{total_val:,.0f}")
    c7.metric("💵 Paid (₦)",  f"{approved_val:,.0f}")

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Claims Status Distribution")
        sc = df["status"].value_counts().reset_index()
        sc.columns = ["Status","Count"]
        colors = {"APPROVED":"#28a745","REJECTED":"#dc3545",
                  "PARTIALLY APPROVED":"#ffc107","PENDING":"#6c757d","UNDER REVIEW":"#17a2b8"}
        fig = px.pie(sc, names="Status", values="Count",
                     color="Status", color_discrete_map=colors)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Claims by Insurance Plan")
        if not patients.empty:
            merged = df.merge(patients[["policy_number","plan"]], on="policy_number", how="left")
            if "plan" in merged.columns:
                pc = merged["plan"].value_counts().reset_index()
                pc.columns = ["Plan","Count"]
                fig2 = px.bar(pc, x="Plan", y="Count", color="Plan",
                              color_discrete_sequence=px.colors.qualitative.Set2)
                st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.subheader("Monthly Claim Volume")
        if "month" in df.columns:
            mv = df.groupby("month").size().reset_index(name="Count")
            fig3 = px.line(mv, x="month", y="Count", markers=True,
                           color_discrete_sequence=["#007bff"],
                           labels={"month":"Month"})
            st.plotly_chart(fig3, use_container_width=True)

    with col4:
        st.subheader("Claim Amount Histogram")
        fig4 = px.histogram(df, x="claim_amount", nbins=20,
                            color_discrete_sequence=["#6f42c1"],
                            labels={"claim_amount":"Claim Amount (₦)"})
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("📋 Latest Claims")
    st.dataframe(df.tail(10), use_container_width=True)


def hmo_incoming_claims():
    st.title("📨 Incoming Claims")
    st.markdown("---")

    df = load(CLAIMS_FILE)
    if df.empty:
        st.info("No claims received yet.")
        return

    pending = df[df["status"] == "PENDING"]

    c1,c2 = st.columns(2)
    c1.metric("⏳ Awaiting Review", len(pending))
    c2.metric("📨 Total Received",  len(df))

    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        hf = st.selectbox("Filter Hospital", ["All"] + sorted(df["hospital"].unique().tolist()))
    with col2:
        mf = st.selectbox("Filter Month", ["All"] + MONTH_NAMES)
    with col3:
        sf = st.selectbox("Filter Status", ["Pending Only","All"])

    filtered = pending if sf == "Pending Only" else df.copy()
    if hf != "All":
        filtered = filtered[filtered["hospital"] == hf]
    if mf != "All":
        filtered = filtered[filtered["month"] == mf]

    st.dataframe(filtered, use_container_width=True)


def hmo_process_claims():
    st.title("✅ Process Claims")
    st.markdown("---")

    if st.session_state["role"] == "hmo_officer":
        st.warning("🔒 Officers have read-only access. Contact HMO Admin to process claims.")
        df = load(CLAIMS_FILE)
        st.dataframe(df, use_container_width=True)
        return

    df = load(CLAIMS_FILE)
    if df.empty:
        st.warning("No claims available.")
        return

    actionable = df[df["status"].isin(["PENDING","UNDER REVIEW"])]
    if actionable.empty:
        st.success("🎉 All claims have been processed!")
        return

    st.metric("Claims Awaiting Decision", len(actionable))
    claim_ids = actionable["claim_id"].tolist()
    selected_id = st.selectbox("Select Claim ID", claim_ids)
    claim = actionable[actionable["claim_id"] == selected_id].iloc[0]

    patients = load(PATIENTS_FILE)
    plan = "Basic"
    if not patients.empty:
        pm = patients[patients["policy_number"] == claim["policy_number"]]
        if not pm.empty:
            plan = pm.iloc[0]["plan"]

    all_claims = load(CLAIMS_FILE)
    claims_count = len(all_claims[all_claims["policy_number"] == claim["policy_number"]])

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📋 Claim Details")
        detail_data = {
            "Field": ["Claim ID","Patient","Policy","Plan","Hospital",
                      "Diagnosis","Treatment","Amount Requested",
                      "Date Submitted","Previous Claims Count"],
            "Value": [claim["claim_id"], claim["patient_name"],
                      claim["policy_number"], plan, claim["hospital"],
                      claim["diagnosis"], claim["treatment"],
                      f"₦{int(claim['claim_amount']):,}",
                      claim["date_submitted"], claims_count]
        }
        st.table(pd.DataFrame(detail_data))

    with col2:
        st.subheader("🤖 Auto-Assessment")
        auto_dec, auto_reason = auto_process_claim(
            float(claim["claim_amount"]), plan, claims_count
        )
        risk = risk_score(claims_count, float(claim["claim_amount"]))

        color_map = {"APPROVED":"success","REJECTED":"error","PARTIALLY APPROVED":"warning"}
        getattr(st, color_map.get(auto_dec, "info"))(f"**Recommended Decision:** {auto_dec}")
        st.info(f"**Reason:** {auto_reason}")
        st.warning(f"**Risk Level:** {risk}")

        st.markdown("---")
        st.subheader("✍️ Final Decision")
        decision_opts = ["APPROVED","PARTIALLY APPROVED","REJECTED","UNDER REVIEW"]
        dec_idx = decision_opts.index(auto_dec) if auto_dec in decision_opts else 0
        decision = st.selectbox("Decision", decision_opts, index=dec_idx)
        reason   = st.text_area("Reason / Remarks", value=auto_reason)

        if st.button("📤 Submit Decision & Notify Hospital", type="primary", use_container_width=True):
            df.loc[df["claim_id"] == selected_id, "status"]          = decision
            df.loc[df["claim_id"] == selected_id, "decision_reason"] = reason
            df.loc[df["claim_id"] == selected_id, "processed_by"]    = st.session_state["user"]
            df.loc[df["claim_id"] == selected_id, "date_processed"]  = today()
            save(df, CLAIMS_FILE)

            # In-app notification → Hospital
            icon_map = {"APPROVED":"✅","REJECTED":"❌","PARTIALLY APPROVED":"🟡","UNDER REVIEW":"🔍"}
            icon = icon_map.get(decision, "📋")
            add_notification(
                recipient_org=claim["hospital"],
                recipient_role="hospital_admin",
                message=(f"{icon} Claim {selected_id} for {claim['patient_name']} "
                         f"→ **{decision}**. Reason: {reason}"),
                notif_type=decision.replace(" ","_"),
                claim_id=selected_id
            )

            # Email notification → Hospital admins
            notify_hospital_by_email(
                claim["hospital"], selected_id, claim["patient_name"],
                decision, reason, st.session_state["user"]
            )

            log_action(st.session_state["user"], st.session_state["role"],
                       "PROCESS_CLAIM",
                       f"Claim {selected_id} → {decision} | {reason[:60]}")

            st.success(f"✅ Decision **{decision}** recorded. Hospital notified in-app and by email.")
            st.rerun()


def hmo_analytics():
    st.title("📈 Analytics & Insights")
    st.markdown("---")

    df       = load(CLAIMS_FILE)
    patients = load(PATIENTS_FILE)

    if df.empty:
        st.info("No data available.")
        return

    tab1, tab2, tab3 = st.tabs(["📊 Claims Analytics","🏥 Hospital Performance","📅 Trends"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Overall Approval Rate")
            approved = len(df[df["status"] == "APPROVED"])
            rate = round(approved / len(df) * 100, 1) if len(df) > 0 else 0
            fig = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=rate,
                title={"text":"Approval Rate (%)"},
                delta={"reference": 70},
                gauge={"axis":{"range":[0,100]},
                       "bar":{"color":"#28a745"},
                       "steps":[{"range":[0,50],"color":"#ffe0e0"},
                                 {"range":[50,75],"color":"#fff8cc"},
                                 {"range":[75,100],"color":"#d4edda"}]}
            ))
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            st.subheader("Average Claim Amount by Status")
            avg = df.groupby("status")["claim_amount"].mean().reset_index()
            avg.columns = ["Status","Avg Amount (₦)"]
            fig2 = px.bar(avg, x="Status", y="Avg Amount (₦)", color="Status",
                          color_discrete_map={"APPROVED":"#28a745","REJECTED":"#dc3545",
                                              "PARTIALLY APPROVED":"#ffc107","PENDING":"#6c757d"})
            st.plotly_chart(fig2, use_container_width=True)

        st.subheader("Top 5 Diagnoses by Claim Volume")
        if "diagnosis" in df.columns:
            diag = df["diagnosis"].value_counts().head(5).reset_index()
            diag.columns = ["Diagnosis","Count"]
            fig3 = px.bar(diag, x="Count", y="Diagnosis", orientation="h",
                          color_discrete_sequence=["#6f42c1"])
            st.plotly_chart(fig3, use_container_width=True)

    with tab2:
        st.subheader("Claims Per Hospital")
        hosp = df.groupby("hospital").agg(
            Claims=("claim_id","count"),
            Total_Claimed=("claim_amount","sum"),
            Avg_Claim=("claim_amount","mean")
        ).reset_index()
        hosp.columns = ["Hospital","Claims","Total Claimed (₦)","Avg Claim (₦)"]
        st.dataframe(hosp, use_container_width=True)
        fig4 = px.bar(hosp, x="Hospital", y="Claims",
                      color_discrete_sequence=["#007bff"])
        st.plotly_chart(fig4, use_container_width=True)

    with tab3:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Monthly Total Claim Value")
            if "month" in df.columns:
                trend = df.groupby("month")["claim_amount"].sum().reset_index()
                trend.columns = ["Month","Total (₦)"]
                fig5 = px.line(trend, x="Month", y="Total (₦)", markers=True,
                               color_discrete_sequence=["#6f42c1"])
                st.plotly_chart(fig5, use_container_width=True)
        with col2:
            st.subheader("Monthly Approval Rate")
            if "month" in df.columns:
                def apr(x):
                    return round((x["status"]=="APPROVED").sum()/len(x)*100,1)
                mrate = df.groupby("month").apply(apr).reset_index()
                mrate.columns = ["Month","Approval Rate (%)"]
                fig6 = px.bar(mrate, x="Month", y="Approval Rate (%)",
                              color_discrete_sequence=["#28a745"])
                st.plotly_chart(fig6, use_container_width=True)


def hmo_billing():
    st.title("💰 Billing & Monthly Accounting")
    st.markdown("---")

    df       = load(CLAIMS_FILE)
    patients = load(PATIENTS_FILE)

    if df.empty:
        st.info("No billing data.")
        return

    month_sel = st.selectbox("Select Month", MONTH_NAMES)
    month_df  = df[df["month"] == month_sel]

    st.subheader(f"📅 {month_sel} — Claims Summary")
    c1,c2,c3,c4,c5 = st.columns(5)
    c1.metric("Total Claims",         len(month_df))
    c2.metric("Total Submitted (₦)",  f"{month_df['claim_amount'].sum():,.0f}")
    c3.metric("✅ Approved Payout",   f"₦{month_df[month_df['status']=='APPROVED']['claim_amount'].sum():,.0f}")
    c4.metric("🟡 Partial Payout",    f"₦{month_df[month_df['status']=='PARTIALLY APPROVED']['claim_amount'].sum():,.0f}")
    c5.metric("❌ Rejected Value",    f"₦{month_df[month_df['status']=='REJECTED']['claim_amount'].sum():,.0f}")

    st.markdown("---")

    if not patients.empty:
        st.subheader("💵 Premium Revenue Estimate")
        pc = patients["plan"].value_counts().reset_index()
        pc.columns = ["Plan","Members"]
        pc["Monthly Premium (₦)"] = pc["Plan"].map(PLAN_PREMIUMS)
        pc["Total Revenue (₦)"]   = pc["Members"] * pc["Monthly Premium (₦)"]
        st.dataframe(pc, use_container_width=True)

        total_rev   = pc["Total Revenue (₦)"].sum()
        total_payout = (month_df[month_df["status"]=="APPROVED"]["claim_amount"].sum() +
                        month_df[month_df["status"]=="PARTIALLY APPROVED"]["claim_amount"].sum() * 0.5)
        net = total_rev - total_payout

        col1, col2, col3 = st.columns(3)
        col1.metric("💵 Premium Revenue",    f"₦{total_rev:,.0f}")
        col2.metric("💸 Estimated Payout",   f"₦{total_payout:,.0f}")
        col3.metric("📊 Net Position",        f"₦{net:,.0f}",
                    delta="Surplus" if net >= 0 else "Deficit")

        fig = go.Figure(data=[
            go.Bar(name="Premium Revenue",  x=[month_sel], y=[total_rev],   marker_color="#28a745"),
            go.Bar(name="Claims Payout",    x=[month_sel], y=[total_payout],marker_color="#dc3545"),
            go.Bar(name="Net Position",     x=[month_sel], y=[net],          marker_color="#007bff"),
        ])
        fig.update_layout(barmode="group", title=f"{month_sel} — Revenue vs Payout")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader(f"📋 {month_sel} Claims Detail")
    if not month_df.empty:
        st.dataframe(month_df, use_container_width=True)
        st.download_button(
            f"📥 Export {month_sel} Report (Excel)",
            to_excel(month_df),
            f"billing_{month_sel}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info(f"No claims found for {month_sel}.")


def hmo_risk():
    st.title("⚠️ Risk Scoring & Fraud Detection")
    st.markdown("---")

    df       = load(CLAIMS_FILE)
    patients = load(PATIENTS_FILE)

    if df.empty or patients.empty:
        st.info("Insufficient data for risk analysis.")
        return

    agg = df.groupby("policy_number").agg(
        claims_count   = ("claim_id",      "count"),
        total_claimed  = ("claim_amount",  "sum"),
        avg_claimed    = ("claim_amount",  "mean")
    ).reset_index()

    merged = patients.merge(agg, on="policy_number", how="left").fillna(0)
    merged["risk_level"] = merged.apply(
        lambda r: risk_score(int(r["claims_count"]), float(r["total_claimed"])), axis=1
    )

    high   = len(merged[merged["risk_level"] == "🔴 High"])
    medium = len(merged[merged["risk_level"] == "🟡 Medium"])
    low    = len(merged[merged["risk_level"] == "🟢 Low"])

    c1,c2,c3 = st.columns(3)
    c1.metric("🔴 High Risk",   high)
    c2.metric("🟡 Medium Risk", medium)
    c3.metric("🟢 Low Risk",    low)

    st.markdown("---")

    col1, col2 = st.columns(2)
    with col1:
        rc = merged["risk_level"].value_counts().reset_index()
        rc.columns = ["Risk Level","Count"]
        fig = px.pie(rc, names="Risk Level", values="Count",
                     color="Risk Level",
                     color_discrete_map={"🔴 High":"#dc3545","🟡 Medium":"#ffc107","🟢 Low":"#28a745"})
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        fig2 = px.scatter(merged, x="claims_count", y="total_claimed",
                          color="risk_level", hover_data=["name","policy_number","plan"],
                          color_discrete_map={"🔴 High":"#dc3545","🟡 Medium":"#ffc107","🟢 Low":"#28a745"},
                          labels={"claims_count":"Claims Count","total_claimed":"Total Claimed (₦)"},
                          title="Risk Scatter Plot")
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("📋 Patient Risk Table")
    cols = ["name","policy_number","plan","claims_count","total_claimed","avg_claimed","risk_level"]
    avail = [c for c in cols if c in merged.columns]
    st.dataframe(merged[avail].sort_values("risk_level"), use_container_width=True)

    st.markdown("---")
    st.subheader("🚨 High-Risk Alerts")
    hr = merged[merged["risk_level"] == "🔴 High"]
    if hr.empty:
        st.success("✅ No high-risk patients detected at this time.")
    else:
        for _, r in hr.iterrows():
            st.error(
                f"⚠️ **{r['name']}** | Policy: {r['policy_number']} | Plan: {r['plan']} | "
                f"Claims: {int(r['claims_count'])} | Total Claimed: ₦{r['total_claimed']:,.0f}"
            )


def hmo_audit():
    st.title("📋 Audit Log")
    st.markdown("---")

    df = load(AUDIT_FILE)
    if df.empty:
        st.info("No audit records yet.")
        return

    col1, col2 = st.columns(2)
    with col1:
        uf = st.selectbox("Filter by User",   ["All"] + sorted(df["user"].unique().tolist()))
    with col2:
        af = st.selectbox("Filter by Action", ["All"] + sorted(df["action"].unique().tolist()))

    filtered = df.copy()
    if uf != "All":
        filtered = filtered[filtered["user"] == uf]
    if af != "All":
        filtered = filtered[filtered["action"] == af]

    st.metric("Records", len(filtered))
    st.dataframe(filtered.sort_values("timestamp", ascending=False), use_container_width=True)

    if not filtered.empty:
        st.download_button("📥 Export Audit Log (Excel)", to_excel(filtered), "audit_log.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def hmo_notifications():
    st.title("🔔 HMO Notifications")
    org  = st.session_state["org"]
    role = st.session_state["role"]
    st.markdown("---")

    notifs = get_notifications(org, role)
    if notifs.empty:
        st.info("No notifications yet.")
        return

    col1, col2 = st.columns([3,1])
    col1.metric("Total", len(notifs))
    with col2:
        if st.button("✅ Mark All Read", use_container_width=True):
            mark_all_read(org, role)
            st.rerun()

    for _, n in notifs.iterrows():
        is_read = str(n["read"]) == "True"
        icon = {"NEW_CLAIM":"🧾","APPROVED":"✅","REJECTED":"❌",
                "PARTIALLY_APPROVED":"🟡","UNDER_REVIEW":"🔍"}.get(str(n["type"]),"🔔")
        msg = f"{icon} {n['message']}  \n🕐 *{n['timestamp']}*"
        if not is_read:
            st.success(msg)
        else:
            st.info(msg)


def hmo_user_management():
    st.title("👤 User Management")
    st.markdown("---")

    tab1, tab2 = st.tabs(["👥 All Users","➕ Add New User"])

    with tab1:
        users = load(USERS_FILE)
        if not users.empty:
            display = users.drop(columns=["password_hash"], errors="ignore")
            st.dataframe(display, use_container_width=True)

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            nu_user  = st.text_input("Username *")
            nu_pass  = st.text_input("Password *", type="password")
            nu_email = st.text_input("Email Address *")
        with col2:
            nu_role  = st.selectbox("Role", ["hospital_admin","hospital_officer","hmo_admin","hmo_officer"])
            nu_org   = st.text_input("Organization *")

        if st.button("➕ Create User", type="primary", use_container_width=True):
            if not nu_user or not nu_pass or not nu_email or not nu_org:
                st.error("All fields are required.")
                return
            users = load(USERS_FILE)
            if not users.empty and nu_user in users["username"].values:
                st.error(f"Username **{nu_user}** already exists.")
                return
            new_row = pd.DataFrame([{
                "username":nu_user,"password_hash":hp(nu_pass),
                "role":nu_role,"organization":nu_org,
                "email":nu_email,"created_at":now()
            }])
            users = pd.concat([users, new_row], ignore_index=True) if not users.empty else new_row
            save(users, USERS_FILE)
            log_action(st.session_state["user"], st.session_state["role"],
                       "ADD_USER", f"Created user {nu_user} as {nu_role} in {nu_org}")
            st.success(f"✅ User **{nu_user}** created successfully.")
            st.rerun()


def hmo_download():
    st.title("📥 Download Reports")
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Claims Report")
        df = load(CLAIMS_FILE)
        if not df.empty:
            st.download_button("📥 All Claims (Excel)", to_excel(df), "all_claims.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col2:
        st.subheader("Patients Report")
        patients = load(PATIENTS_FILE)
        if not patients.empty:
            st.download_button("📥 All Patients (Excel)", to_excel(patients), "all_patients.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col3:
        st.subheader("Audit Report")
        audit = load(AUDIT_FILE)
        if not audit.empty:
            st.download_button("📥 Audit Log (Excel)", to_excel(audit), "audit_log.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────────────────────
# MAIN ROUTER
# ─────────────────────────────────────────────────────────────
def main():
    init_data()

    if "user" not in st.session_state:
        login_page()
        return

    page = sidebar()
    role = st.session_state["role"]

    if "hospital" in role:
        if   page == "📊 Dashboard":           hospital_dashboard()
        elif page == "🧾 Submit Claim":        submit_claim()
        elif page == "📋 My Claims":           hospital_my_claims()
        elif page == "👥 Patients":            hospital_patients()
        elif page == "🔔 Notifications":       hospital_notifications()
        elif page == "📥 Download Reports":    hospital_download()
    else:  # hmo
        if   page == "📊 Dashboard":           hmo_dashboard()
        elif page == "📨 Incoming Claims":     hmo_incoming_claims()
        elif page == "✅ Process Claims":      hmo_process_claims()
        elif page == "📈 Analytics":           hmo_analytics()
        elif page == "💰 Billing & Accounting":hmo_billing()
        elif page == "⚠️ Risk Scoring":       hmo_risk()
        elif page == "📋 Audit Log":           hmo_audit()
        elif page == "🔔 Notifications":       hmo_notifications()
        elif page == "📥 Download Reports":    hmo_download()
        elif page == "👤 User Management":     hmo_user_management()


if __name__ == "__main__":
    main()
