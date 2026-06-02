# CRM Customer Search — Setup Instructions

## Step 1: Open Your Google Sheet
Open your CRM Google Sheet (the one with Customers, Contacts, Purchase Orders, etc.)

## Step 2: Open Apps Script
- Click **Extensions** → **Apps Script**
- Delete any existing code in the editor

## Step 3: Paste the Script
- Open the file **CRM_Search_Script.gs**
- Select all the text (Ctrl+A / Cmd+A)
- Paste it into the Apps Script editor
- Click **Save** (💾 icon or Ctrl+S)

## Step 4: Run Setup
- In the Apps Script editor, select the function **setupCRM** from the dropdown
- Click **▶ Run**
- Accept any permissions Google asks for (this is normal — the script only accesses your own spreadsheet)

## Step 5: Return to Your Sheet
- Close the Apps Script tab and go back to your Google Sheet
- You'll see new tabs: 🔍 Search, 📋 Customer Details, 👤 Contacts, 🛒 Purchase Orders, 🎫 Service Tickets
- A new menu **🔍 CRM Search** appears in the top menu bar

---

## How to Search

### Option A — Via the Menu (Recommended)
1. Click **🔍 CRM Search** → **Search Customer**
2. A search dialog pops up
3. Type a **customer name** (e.g., "UPS") or a **contact name** (e.g., "Omri" or "Shamir")
4. Press Enter or click **Search**

### Option B — Type directly in the Search tab
1. Click the **🔍 Search** tab
2. Type your query in cell **C4**
3. Use the menu **🔍 CRM Search** → **Search Customer**

---

## What Each Tab Shows

| Tab | Contents |
|-----|----------|
| 🔍 Search | Search box + quick results list |
| 📋 Customer Details | Full customer profile + recurring activity |
| 👤 Contacts | All contacts linked to the customer |
| 🛒 Purchase Orders | All POs for the customer |
| 🎫 Service Tickets | All support tickets for the customer |

---

## Important: Source Sheet Names Must Match
The script reads from these exact sheet names in your file:
- `Customers`
- `Contacts ` *(note the trailing space)*
- `Purchase Orders ` *(note the trailing space)*
- `Service Ticket ` *(note the trailing space)*
- `Recurring Activity`

If you rename those sheets, update the `SHEETS` config object at the top of the script accordingly.
