# Sales CRM System Manual
## Professional & Scalable Sales-Focused CRM Blueprint

This manual outlines the configuration and rules for a sales-focused CRM designed to help Sales Representatives follow up with leads in under 30 seconds, ensuring no lead is ever lost.

---

## 🛠️ Relational Database Layout

Import the updated [leads_template.csv](file:///e:/CRM/Clients%20CRM/templates/leads_template.csv) layout into Airtable or Google Sheets.

---

## 📐 Formulas & Metric Calculations

Copy and paste these formulas directly into your Airtable fields:

### 1. WhatsApp Link Generator
Generates a clickable chat link directly from the phone number field:
* **Airtable**:
```javascript
IF({Phone Number}, "https://wa.me/" & SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({Phone Number}, "+", ""), "-", ""), " ", ""), "(", ""), ")", ""))
```

### 2. Follow-Up Alert System (Reminder)
Flags whether a follow-up is overdue, due today, or scheduled:
* **Airtable**:
```javascript
IF(
  AND({Next Follow-up Date}, {Next Follow-up Date} < TODAY(), NOT(OR({Lead Status} = "Won", {Lead Status} = "Lost", {Lead Status} = "Not Interested"))), 
  "🚨 OVERDUE", 
  IF(
    AND({Next Follow-up Date}, {Next Follow-up Date} = TODAY(), NOT(OR({Lead Status} = "Won", {Lead Status} = "Lost", {Lead Status} = "Not Interested"))), 
    "🟡 DUE TODAY", 
    IF({Next Follow-up Date}, "📅 Scheduled", "⚪ No Action Set")
  )
)
```

### 3. "Needs Attention" Inactivity Flag
Triggers a warning if a lead hasn't been contacted or logged for more than 3 days:
* **Airtable**:
```javascript
IF(
  AND(
    NOT(OR({Lead Status} = "Won", {Lead Status} = "Lost", {Lead Status} = "Not Interested")),
    DATETIME_DIFF(TODAY(), IF({Last Contact Date}, {Last Contact Date}, {Date Added}), 'days') > 3
  ),
  "⚠️ Needs Attention",
  "✅ OK"
)
```

### 4. Lead Temperature
* **Airtable**:
```javascript
IF({Lead Status} = "Won", "🔥 Closed Won", 
  IF({Priority} = "Urgent", "⚡ Urgent",
    IF(DATETIME_DIFF(TODAY(), IF({Last Contact Date}, {Last Contact Date}, {Date Added}), 'days') > 14, "❄️ Cold", 
      IF(OR({Lead Status} = "Negotiating", {Lead Status} = "Proposal Sent", {Lead Status} = "Appointment Booked"), "🔥 Hot", "⚡ Warm")
    )
  )
)
```

---

## 🎨 Color-Coded Lead Statuses

Configure the `Lead Status` field as a single-select dropdown with the following color choices:
* 🟢 **Won** (Green) — *Client Purchased*
* 🟡 **Negotiating** (Yellow) — *Discussing contract details*
* 🟠 **Follow Up Required** (Orange) — *Needs outreach*
* 🔵 **Contacted** (Blue) — *Initial outreach complete*
* ⚪ **New Lead** (Light Grey) — *Fresh lead*
* 🔴 **Lost** (Red) — *Opportunity lost*
* ⚫ **Ghosted** (Dark Grey) — *No response after interest*
* ❌ **Not Interested** (Red/Dark) — *Opted out or disqualified*
* 📅 **Appointment Booked** (Purple) — *Meeting scheduled*
* 📄 **Proposal Sent** (Pink) — *Price proposal sent*

---

## 🚥 Conditional Formatting & Highlighting Rules

Configure these rules in your Airtable grid views:
1. **Highlight Overdue Follow-ups (Red)**:
   - Rule: Where `Reminder` is `🚨 OVERDUE`.
   - Action: Highlight record background in **Red**.
2. **Highlight Today's Follow-ups (Yellow)**:
   - Rule: Where `Reminder` is `🟡 DUE TODAY`.
   - Action: Highlight record background in **Yellow**.
3. **Needs Attention Flag**:
   - Rule: Where `Needs Attention` is `⚠️ Needs Attention`.
   - Action: Highlight column text or add an emoji indicator.

---

## 🤖 Sales Automations

Set up these automations to save your sales representatives hours of manual work:

### 1. Move Won Deals to Clients List
* **Trigger**: When `Lead Status` changes to `Won`.
* **Action**: Update the lead record or copy/move the record to the `Clients` table for onboarding.

### 2. Archive Lost Leads after 30 days
* **Trigger**: When `Lead Status` changes to `Lost` or `Not Interested`.
* **Action**: Wait 30 days, then update record field `Archived = TRUE` or move to the `Archive` table.

### 3. Generate Weekly & Daily Follow-Up Queue
* **Trigger**: Every morning at 7:30 AM.
* **Action**: Filter leads where `Assigned Salesperson` is the current user AND (`Next Follow-up Date` is today or earlier). Present this list in a dedicated view called "My Tasks Today".
