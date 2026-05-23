# 🏥 Qozeem AHICPMR
**Qozeem Automated Healthcare Insurance Claims Processing and Patient Management Record**

A comprehensive two-sided insurance claims management platform connecting Hospitals and HMO companies for seamless claim submission, processing, approval/rejection, and automated notifications.

---

## 📋 System Overview

Qozeem AHICPMR is a full-stack healthcare insurance platform that digitizes the entire claims lifecycle:

### 🏥 **Hospital Portal**
Hospitals can:
- Submit insurance claims to HMOs with patient diagnosis, treatment, and claim amount
- Track claim status in real-time
- Register and manage patient records with policy details
- View claim history and decisions
- Receive notifications when HMOs approve/reject claims
- Download reports in Excel format

### 🏢 **HMO Portal**
HMO companies can:
- Receive claim submissions from hospitals
- Review claims with auto-assessment (intelligent recommendation system)
- Approve, partially approve, reject, or flag claims for review
- Process claims with detailed decision reasons
- Send notifications back to hospitals
- Manage patient risk scoring and fraud detection
- View comprehensive analytics and monthly billing reports
- Manage users and audit logs
- Export detailed financial reports

---

## 🎯 What the App Does

### **Core Features**

#### 1. **📨 Claims Workflow**
- **Hospital** → Submits claim with patient diagnosis, treatment, amount
- **HMO** → Receives notification, reviews claim
- **HMO** → Auto-assessment provides recommendation based on:
  - Plan limits (Basic: ₦30K, Standard: ₦150K, Premium: ₦350K, Executive: ₦500K)
  - Patient claims history
  - Total claimed amount threshold (₦500K cap)
- **HMO** → Makes final decision: APPROVED / PARTIALLY APPROVED / REJECTED / UNDER REVIEW
- **Hospital** → Receives in-app & email notification with decision & reason

#### 2. **📊 Dashboards**
**Hospital Dashboard:**
- Total claims submitted, pending, approved, rejected count
- Total claim values
- Claims distribution pie chart
- Monthly claim trends
- Recent claims list

**HMO Dashboard:**
- Total claims received with status breakdown
- Approved vs total payouts
- Claims by plan type
- Hospital performance metrics
- Monthly claim volume
- Claim amount distribution

#### 3. **💰 Billing & Accounting**
- **Monthly summaries** broken down by patient plan
- **Premium revenue calculation** based on enrolled members and plan premiums
- **Payout tracking** for approved claims
- **Net position** (revenue - payout) to show profit/loss
- **Plan analysis** — which plans generate most claims
- **Monthly reports** — exportable Excel files

#### 4. **⚠️ Risk Scoring & Fraud Detection**
Automatically flags patients as:
- 🔴 **High Risk**: >3 claims OR total claimed >₦400K
- 🟡 **Medium Risk**: 2-3 claims OR total claimed ₦150K-₦400K
- 🟢 **Low Risk**: Otherwise

**Features:**
- Risk scatter plot (claims count vs total claimed)
- High-risk patient alerts
- Patient risk table with all metrics

#### 5. **📈 Analytics**
- Overall approval rate gauge
- Claims distribution by status
- Average claim amount by status
- Top diagnoses by frequency
- Hospital performance comparison
- Monthly trends (claim volume & approval rates)
- Plan-based claim analysis

#### 6. **👥 Patient Management**
- Register new patients with policy details
- Track enrollment and expiry dates
- View expiring policies (automatic alerts)
- Patient contact information and demographics
- Export patient lists
- Policy status tracking (Active/Inactive)

#### 7. **🔔 Real-Time Notifications**
**In-App Notifications:**
- Hospital → sees when HMO processes claim (decision + reason)
- HMO → sees when hospital submits new claim

**Email Notifications** (if configured):
- Hospital submits claim → HMO admins receive email alert
- HMO processes claim → Hospital admins receive email with decision

**Notification Bell:**
- Unread count displayed in sidebar
- Mark all as read button
- Timestamp tracking

#### 8. **📋 Audit Log**
- Every action logged: login, logout, claim submission, claim processing, patient registration
- Filterable by user and action type
- Timestamp tracking
- Exportable to Excel

#### 9. **📥 Report Downloads**
- Export claims as Excel
- Export patient lists as Excel
- Export audit logs as Excel
- Monthly billing reports with revenue/payout analysis

#### 10. **👤 User Management** (HMO Admin only)
- Add new users
- Assign roles: hospital_admin, hospital_officer, hmo_admin, hmo_officer
- Track user creation dates
- View all users

---

## 🔐 Role-Based Access

| Role | Can Do | Cannot Do |
|---|---|---|
| **Hospital Admin** | Submit claims, register patients, view all records, download reports | Process claims, manage HMO users |
| **Hospital Officer** | Submit claims, view claims and patients | Register patients, process claims |
| **HMO Admin** | Process claims, manage users, view all analytics, manage risk | Submit claims |
| **HMO Officer** | View claims and analytics (read-only), cannot process | Process claims, manage users |

---

## 💻 Demo Credentials

| Role | Username | Password | Organization |
|---|---|---|---|
| Hospital Admin | `hospital_admin` | `hospital123` | Lagos General Hospital |
| Hospital Officer | `hospital_officer` | `officer123` | Lagos General Hospital |
| HMO Admin | `hmo_admin` | `hmo123` | Qozeem Health HMO |
| HMO Officer | `hmo_officer` | `hmooff123` | Qozeem Health HMO |

---

## 📂 File Structure

```
qozeem_ahicpmr/
├── app.py                    ← Main application (all features)
├── requirements.txt          ← Python dependencies
├── .gitignore               ← Protects secrets
├── .streamlit/
│   └── secrets.toml        ← Email configuration (local only)
└── README.md               ← This file
```

**Data Files** (auto-created on first run):
- `claims.csv` — All submitted claims and decisions
- `patients.csv` — Patient & policy information
- `users.csv` — User accounts and roles
- `audit_log.csv` — Activity log
- `notifications.csv` — In-app notifications

---

## 🚀 Deployment on Streamlit Cloud

### **Step 1: Create GitHub Repository**
1. Go to [github.com/new](https://github.com/new)
2. Create new repo: `qozeem-ahicpmr`
3. Upload these files:
   - `app.py`
   - `requirements.txt`
   - `.gitignore`
   - `.streamlit/secrets.toml` (optional — only for email)
   - `README.md`

### **Step 2: Deploy to Streamlit Cloud**
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **New app** → select your repo
4. Set main file to `app.py`
5. Click **Deploy**

### **Step 3: Configure Secrets** (Optional — for email notifications)
1. In Streamlit Cloud app page → **Advanced settings** → **Secrets**
2. Add (if using Gmail for email alerts):
```toml
[email]
sender = "your-email@gmail.com"
password = "your-16-char-app-password"
```
3. Get Google App Password: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)

### **Step 4: Access Your App**
- Your app URL will be: `https://share.streamlit.io/[your-username]/qozeem-ahicpmr`
- Share with hospitals and HMOs
- They log in with credentials

---

## 📊 Sample Data Included

The system comes pre-loaded with:
- **6 sample patients** with active policies
- **6 sample claims** in various statuses
- **Demo HMO and Hospital** organizations

Use these to test the full workflow!

---

## 🔧 Key Technical Features

**Backend:**
- Python 3.8+
- Streamlit (web framework)
- Pandas (data management)
- Plotly (interactive charts)
- SHA-256 password hashing
- CSV-based data persistence

**Frontend:**
- Responsive multi-tab layouts
- Interactive dashboards with Plotly
- Real-time notifications
- Form validation
- Session management

**Security:**
- Password hashing (SHA-256)
- Role-based access control
- Audit logging for all actions
- Session-based authentication
- Secrets management for sensitive data

---

## 💡 Business Value

**For Hospitals:**
- ✅ Faster claim submissions (5 minutes vs days of paperwork)
- ✅ Real-time claim tracking
- ✅ Instant notifications on approvals/rejections
- ✅ Patient data centralized and searchable

**For HMOs:**
- ✅ Automated claim assessment (reduces manual review time)
- ✅ Risk detection (fraud prevention)
- ✅ Financial dashboard (revenue vs payout tracking)
- ✅ Audit trail (regulatory compliance)
- ✅ Analytics (business insights)

---

## 📈 Future Enhancements

- Database backend (PostgreSQL) for production-grade scalability
- SMS notifications
- Mobile app (React Native)
- Payment gateway integration
- Advanced fraud detection (ML models)
- Digital signatures for claims
- OCR for medical documents
- Multi-language support (Yoruba, Igbo, Hausa)
- FHIR compliance for healthcare data standards

---

## 📞 Support

For issues or feature requests, contact: **support@qozeem.com**

---

## 📄 License

© 2024 Qozeem Health Technologies. All rights reserved.

---

**Version:** 1.0  
**Last Updated:** January 2024  
**Built with:** Streamlit | Python | Plotly | Pandas
