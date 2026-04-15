<img width="376" height="944" alt="image" src="https://github.com/user-attachments/assets/8c62e8e9-749f-4fb9-8783-3d28e086d9d3" />

<img width="377" height="945" alt="image" src="https://github.com/user-attachments/assets/d679b641-8aa3-4700-a095-33a55e23ea82" />


# Tour Admin Autofill — Chrome Extension

Automatically fills your admin panel tour form from Google Sheets + Google Docs data.

---

## Installation

1. Open Chrome → go to `chrome://extensions`
2. Enable **Developer mode** (top right toggle)
3. Click **Load unpacked**
4. Select the `tour-extension` folder

The extension icon will appear in your toolbar.

---

## Google API Setup (one-time, free)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use existing)
3. Go to **APIs & Services → Library**
4. Enable **Google Sheets API**
5. Enable **Google Docs API**
6. Go to **APIs & Services → Credentials**
7. Click **Create Credentials → API Key**
8. Copy the key — paste it into the extension settings

> **Note:** Restrict the key to Sheets + Docs APIs for security.

---

## Google Sheet Setup

Your sheet should have these columns (headers in row 1):

| Column | Header | Example |
|--------|--------|---------|
| A | Tour Title | Rome Colosseum Tour |
| B | Doc URL | https://docs.google.com/document/d/... |
| C | Service Type | Guide |
| D | Tour Type | Private |
| E | Activity Type | Walking |
| F | Sub Type | Historical |
| G | Description | (or leave blank, use Doc URL) |
| H | You Will See | Colosseum,Forum,Palatine Hill |
| I | You Will Learn | Roman history,Architecture |
| J | Mandatory Information | Comfortable shoes required |
| K | Recommended Information | Water bottle,Sunscreen |
| L | Included | Guide,Entry tickets |
| M | Not Included | Meals,Transport |
| N | Activity For | All ages |
| O | Voucher Type | Digital |
| P | No Of Pax | 10 |
| Q | Guide Language Instant | English |
| R | Guide Language Request | French,Spanish |
| S | Country | Italy |
| T | City | Rome |
| U | Longitude | 12.4922 |
| V | Latitude | 41.8902 |
| W | Meeting Point | Colosseum main entrance |
| X | Pickup Instructions | Look for blue flag,Wait at gate A |
| Y | End Point | Roman Forum |
| Z | Tags | History,Culture |
| AA | Price Model | Fixed Rate |
| AB | Currency | EUR |
| AC | Rate | 150 |
| AD | Rate B2C | 120 |
| AE | Rate Request | 140 |
| AF | Rate Request B2C | 110 |
| AG | Extra Hour | 30 |
| AH | Extra Hour B2C | 25 |
| AI | Extra Hour Request | 28 |
| AJ | Extra Hour Request B2C | 22 |
| AK | Holiday Supplement | 15 |
| AL | Weekend Supplement | 10 |
| AM | Start Date | 01/06/2025 |
| AN | End Date | 31/12/2025 |
| AO | Start Time | 09:00 |
| AP | End Time | 17:00 |
| AQ | Duration | 3h |
| AR | Cancellation | Free cancellation up to 24h |
| AS | Release | 2 work days |
| AT | B2C Enabled | true |
| AU | B2B Enabled | true |

**The extension auto-detects columns by header name** — column order doesn't matter as long as headers match.

---

## How to Use

1. Click the extension icon in your toolbar
2. Enter your **Sheet ID** (from the URL) and **API Key**
3. Click **Save & Load Tours** — your tours appear as a list
4. Go to the **Add Tour** page in your admin panel
5. Search for and click the tour you want to fill
6. Click **Open All Sections** first (opens all accordion panels)
7. Wait ~2 seconds for all panels to render
8. Click **▶ Start Autofill**
9. Review the filled fields and click Save in the admin panel

---

## Notes

- The extension fills fields but **does not submit** — you review before saving
- For ng-select dropdowns, values must match exactly what's in the dropdown list
- Date fields (Start/End Date) may need manual selection due to the date picker component
- If a section is closed when autofill runs, those fields will fail — always use "Open All Sections" first
- Description can come from the Google Doc linked in column B, or directly from column G
