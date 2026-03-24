# 📊 Weekly Update Report Tool

A web app for generating weekly AfterSchool21 data update reports — no coding needed.

🔗 **App Link:** [https://weekly-update-report-tool-cbkr4ex88oea46afxbc6tf.streamlit.app/](https://weekly-update-report-tool-cbkr4ex88oea46afxbc6tf.streamlit.app/)

---

## 📋 How to Use the App (Every Week)

### Part 1 — Download Excel Files from Transact

#### Step 1.1 — Set Up the Dashboard
1. Copy Guo's Dashboard
2. Create the following filters:
   - **Participant Demographics:** Adult
   - **Daily Activity Attendance Summary:** Date, and Adult
   - **Participant Attendance Hours:** Adult

#### Step 1.2 — Download Three Required Excel Files
1. Click the **Date** button in the filter and set the date range from **07/01/2025** to the current date
2. Download three files by adjusting the Adult filter:

   | File | Adult Filter Setting | Used As |
   |---|---|---|
   | All Sites file | **Omit** | `all_sheets` |
   | Adults/Family file | **Yes** | `adults_sheets` |
   | Students file | **No** | `students_sheets` |

> ⚠️ When downloading, make sure the dashboard has **fully loaded** before saving the Excel file.

---

### Part 2 — Open the Weekly Update Report Tool

Go to: [https://weekly-update-report-tool-cbkr4ex88oea46afxbc6tf.streamlit.app/](https://weekly-update-report-tool-cbkr4ex88oea46afxbc6tf.streamlit.app/)

#### Step 2.1 — Upload Excel Files
Upload your three downloaded Excel files into the correct slots:
- 🎒 **Students File** → the file downloaded with Adult = "No"
- 👪 **Adults / Family File** → the file downloaded with Adult = "Yes"
- 📋 **All Sites File** → the file downloaded with Adult = "Omit"

#### Step 2.2 — Fill in Program Info & Targets
- Enter the **Program / District Name** and **Report Date**
- Set the **Target # of Students, Target # of Parents, and Target Literacy Workshops** for each site, in the same order as they appear in your Participants by Hour Band report

#### Step 2.3 — Process Files
Click **"Process Files"** to run the pipeline.

#### Step 2.4 — Tag Family Activities
After processing, a list of activities will appear for each site. Tag each activity as:
- **Literacy Workshop** — counts toward completed workshops
- **Family Engagement Event** — counts toward engagement events
- **Neither** — everything else

#### Step 2.5 — Generate & Download
Click **"Generate Word Doc + Excel"** to download both output files.

---

## 📁 Excel Output Sheet Guide

| Sheet | Description |
|---|---|
| **Student Statistics** | Corresponds to the Student Summary Statistics table in the weekly report |
| **Parents Served Summary** | Corresponds to the last column of the Family Component Summary Statistics table. The number of workshops and family events still need to be manually filled if not tagged in the app. |
| **Missing Summary** | Corresponds to the Missing Student & Staff Information table. Similar to the table in the weekly updates. |
| **Missing Site Info** | Pulls out all rows that do not have a "Site" input. If empty, there is no missing site info. |
| **Total Missing Info** | Pulls out all rows with missing info from the Participant Demographics table, highlighted in red. |
| **Young DOB** | Flags students with birthdays after 2025-01-01. If blank, you can ignore this sheet. |
| **Site sheets** | One sheet per site — Site Summary Reports sorted by Activity then Session. Future events are filtered out. Missing data is flagged with "-". |

### Notes on OSIS Flags (Missing Summary sheet)

There are 3 ways missing OSIS numbers are flagged — use the one that fits your grant:

- **ParticipantID_Missing** — flags every ID in the ParticipantID column that does not start with 1 or 2, or is not 9 digits. *(Use this for NYC grants — this is the District ID column)*
- **State_ParticipantID_Missing** — flags every ID in the State_ParticipantID column that does not start with 1 or 2, or is not 9 digits
- **OSIS_Missing** — flags both columns if they don't start with 1 or 2, aren't 9 digits, or the two columns don't match each other

---

## 🚀 Developer Setup (One-Time — for whoever manages the app)

### Step 1 — Create a GitHub account
Go to [https://github.com](https://github.com) and sign up for a free account.

### Step 2 — Create a new repository
1. Click **+** → **New repository**
2. Name it `weekly-update-report-tool`
3. Set it to **Public**
4. Click **Create repository**

### Step 3 — Upload the two code files
In the repository, click **Add file → Upload files** and upload:
- `app.py`
- `requirements.txt`

Click **Commit changes**.

### Step 4 — Deploy on Streamlit Cloud
1. Go to [https://share.streamlit.io](https://share.streamlit.io)
2. Sign in with GitHub
3. Click **Create app** → **Deploy a public app from GitHub**
4. Select your repository and set Main file path to `app.py`
5. Click **Deploy!**

### Step 5 — Share the link
Once deployed, share the app URL with your team. They just open it in any browser — no installs needed.

### Updating the App
To update the app, edit `app.py` in GitHub and commit. The app redeploys automatically within seconds.

---

## 🛟 Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError` on Streamlit | Make sure `requirements.txt` is uploaded to GitHub alongside `app.py` |
| App shows error after clicking Generate | Double-check the correct files are uploaded to each slot |
| Sheet name error | Confirm your Excel files have the expected sheet names from Transact |
| File named `app (1).py` instead of `app.py` | Rename the file in GitHub using the pencil/edit icon |
| App not updating after changes | Go to Streamlit Cloud → three dots → **Reboot app** |
