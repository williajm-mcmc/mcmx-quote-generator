# MCMX Tools — User Guide

> **Who this guide is for:** Anyone who needs to use the MCMX Tools suite. No technical experience required. Read Section 1 to get the app installed, then jump straight to the tool you need.

---

## Table of Contents

1. [Getting the App](#1-getting-the-app)
2. [Tool Overview](#2-tool-overview)
3. [Quote Generator — Step-by-Step](#3-quote-generator)
4. [OTTO Tool — Step-by-Step](#4-otto-tool)
5. [Thermal Imaging Estimator — Step-by-Step](#5-thermal-imaging-estimator)
6. [IBE Estimator — Step-by-Step](#6-ibe-estimator)
7. [Cost Estimator — Step-by-Step](#7-cost-estimator)
8. [Saving, Loading & Generating Documents](#8-saving-loading--generating-documents)
9. [Tips & Troubleshooting](#9-tips--troubleshooting)

---

## 1. Getting the App

### Windows

#### Step 1 — Download
1. Open a web browser (Chrome, Edge, Firefox — any will work).
2. Go to the Releases page: **https://github.com/williajm-mcmc/mcmx-quote-generator/releases/latest**
3. Under **Assets**, click **`QuoteGenerator-Windows.zip`** to download it.
4. Your browser will save it to your **Downloads** folder.

#### Step 2 — Extract (Unzip)
1. Open your **Downloads** folder (press `Windows key + E`, then click Downloads on the left).
2. Find **`QuoteGenerator-Windows.zip`** — it will look like a folder with a zipper on it.
3. **Right-click** it and choose **"Extract All…"**
4. In the window that pops up, click **Extract**. A new folder called `QuoteGenerator` will appear.
5. Open that folder — inside you will see **`QuoteGenerator.exe`** and several other files.

> **Tip:** Move this `QuoteGenerator` folder somewhere permanent before you start using it — a good place is `C:\Program Files\QuoteGenerator` or a folder on your Desktop labeled `MCMX Tools`. If you run it directly from the zip or Downloads, it may behave unexpectedly.

#### Step 3 — Run the App
1. Double-click **`QuoteGenerator.exe`** inside the extracted folder.
2. Windows may show a blue warning box saying **"Windows protected your PC."** This is normal for apps that aren't downloaded from the Microsoft Store.
   - Click **"More info"**
   - Click **"Run anyway"**
3. The app will open.

#### Step 4 — Pin to the Taskbar (so you can open it with one click)
1. Make sure the app is **running** (you should see it open on your screen).
2. Look at the **taskbar** at the bottom of your screen — you'll see the MCMX Tools icon.
3. **Right-click** the icon in the taskbar.
4. Click **"Pin to taskbar"**.
5. The icon will stay on your taskbar permanently. You can now click it anytime to open the app — even if you close it first.

---

### macOS

#### Step 1 — Download
1. Open **Safari** or any browser.
2. Go to: **https://github.com/williajm-mcmc/mcmx-quote-generator/releases/latest**
3. Under **Assets**, click **`QuoteGenerator-Mac.zip`** to download it.
4. It will save to your **Downloads** folder.

#### Step 2 — Extract and Move to Applications
1. Open your **Downloads** folder (click the Finder icon in your Dock, then click Downloads on the left).
2. Double-click **`QuoteGenerator-Mac.zip`** — it will automatically unzip and show a **`QuoteGenerator.app`** file (it looks like an icon with the app logo).
3. **Drag `QuoteGenerator.app`** into your **Applications** folder. (In Finder, click "Applications" on the left sidebar, then drag the app into that window.)

#### Step 3 — Run the App
1. Open your **Applications** folder and double-click **`QuoteGenerator`**.
2. macOS may show a warning: **"QuoteGenerator cannot be opened because it is from an unidentified developer."**
   - Do NOT click "Move to Trash."
   - Instead, **right-click** (or Control+click) the app icon.
   - Choose **"Open"** from the menu.
   - In the new warning box, click **"Open"**.
   - You only need to do this the first time — after that, it opens normally.

#### Step 4 — Keep in Dock (so you can open it with one click)
1. Make sure the app is **running** — you'll see its icon in the Dock at the bottom of your screen.
2. **Right-click** (or Control+click) the app icon in the Dock.
3. Hover over **"Options"** in the menu.
4. Click **"Keep in Dock"**.
5. The icon stays in your Dock permanently. Click it anytime to open the app.

---

## 2. Tool Overview

The app has five tabs across the top. Here is a quick summary of what each one does:

### Quote Generator
- Build professional project quotes from scratch.
- Fill in project details, attach a customer site photo, and add as many pricing sections as you need.
- Mix and match **BOM (Bill of Material) tables** and **text paragraph sections**.
- Automatically adds pre-formatted MCMX service line items (engineering hours, travel, expenses).
- Generates a polished, formatted **Word document (.docx)** ready to send to customers.
- Save your work as a **.mcmxq** file and pick up where you left off anytime.

### OTTO Tool
- Purpose-built for quoting **OTTO 100 autonomous mobile robot (AMR)** deployments.
- Has a pre-loaded catalog of OTTO 100 components (chargers, carts, manuals, staging hardware).
- Check boxes to select what the customer needs — the tool builds the BOM automatically.
- Save as **.mcmxo** and generate an OTTO-branded Word document.

### Thermal Imaging Estimator
- Estimates cost and builds a schedule for **thermal imaging inspection** projects.
- Enter how many technicians are going, how they're traveling (driving or flying), and how many days of work.
- Automatically determines whether hotel nights are needed and calculates travel days.
- Shows a full work schedule (day-by-day) and a cost breakdown.
- Generates a thermal imaging report Word document.

### IBE Estimator
- Estimates cost and builds a schedule for **IBE (Infrared Building Envelope) inspection** projects.
- Enter the IBE level, total panels, technician count, work days, and travel details.
- Calculates labor, travel, meals, hotel, and mileage costs automatically.
- Shows a day-by-day schedule and a cost summary with margin built in.
- Can be exported to Excel for internal use.

### Cost Estimator
- An internal cost analysis tool that mirrors the Excel cost report format.
- Enter on-site hours, travel hours, and project details.
- Calculates cost before risk, risk/insurance, estimated cost, margin, and recommended resale price.
- Export to a **.mcmxc** file that opens directly in Excel for further review or sharing.

---

## 3. Quote Generator

### What you'll need before you start
- The **project name** (e.g. "Boiler FTSC Upgrade")
- **Customer name** and **location** (city, state)
- Your **contact info** (the tool has fields for Presented By and Account Manager names and emails)
- A rough idea of what parts, services, and sections the quote will include
- Optionally: a **photo of the customer's site**

---

### Step 1 — Fill in the Project Header

At the top of the Quote Generator tab you will see several fields:

1. **Project Name** — Type the name of the project. This appears on the cover page of the quote.
2. **Customer Name** — Type the customer's company name.
3. **Customer Location** — Type the city and state (e.g. "Detroit, MI").
4. **Contact Info** — Fill in your name, email, your account manager's name and email. These appear at the bottom of the cover page.
5. **Proposal Number** — Type the proposal number if you have one (e.g. "P-2024-001"). If you leave this blank, the app will use the filename you choose when saving.

---

### Step 2 — Attach a Customer Photo (Optional)

- Click **"Browse"** or drag and drop an image file into the photo box.
- The photo will appear on the cover page of the generated document.
- To remove it, click **"Clear"**.

---

### Step 3 — Add Sections to Your Quote

Your quote is built from one or more **sections**. There are two types:

#### Option A — Add a BOM Table (for pricing)
1. Click **"＋ Add BOM Table"** (or the equivalent Add Table button).
2. A pricing table section appears. Give it a name in the **Section Heading** field (e.g. "Pricing" or "Bill of Material").
3. Add rows using the **"＋ Add Row"** button, then type in:
   - **Part Number / Description** — what the item is
   - **Qty** — how many
   - **Cost** — your internal cost per unit
   - **Margin %** — your markup percentage (the sale price and line total calculate automatically)
4. To add pre-formatted MCMX service rows, use the three buttons at the bottom of the section:
   - **⚙ Add Engineering Service** — prompts for on-site and off-site hours, builds an AFSE line item at $225/hr
   - **🚗 Add Travel** — prompts for total drive hours, builds a travel line item at $175/hr
   - **📋 Add Expenses** — opens a dialog to enter mileage, meals, and hotel; builds an expenses line item with totals

> **Note:** The mileage, meals, and hotel breakdown you enter into the Expenses dialog is for your reference in the tool only. The generated Word document will show the service name and total price — not the individual breakdown.

#### Option B — Add a Paragraph Section (for text)
1. Click **"＋ Add Paragraph Section"**.
2. Type a **section heading** (e.g. "Scope of Work" or "Project Background").
3. Type your text in the large text area below. You can:
   - Use **Bold** (`Ctrl+B`) and *Italic* (`Ctrl+I`)
   - Create bullet lists
   - Add sub-headings
   - Insert tables using the table icon in the toolbar

#### Reordering Sections
- Each section has **▲ Up** and **▼ Down** arrow buttons at the top right.
- Click them to move a section up or down in the quote.

#### Removing Sections
- Click the **✕ Remove** button at the top right of any section to delete it.

---

### Step 4 — Set the Version

- In the top-right area, you'll see **Version** spinners (major and minor, e.g. V1.0).
- Increase the version number each time you revise and re-send a quote.
- You can also type a note in the **Change History** box to track what changed.

---

### Step 5 — Save Your Project

- Press **Ctrl+S** or go to **File → Save MCMXQ**.
- Choose where to save and give the file a name.
- The app will suggest a name like `MCMX-PROJNAME-20260416 V1.0` automatically.
- Save regularly — your work is not auto-saved.

---

### Step 6 — Generate the Word Document

1. Click the red **"Generate Quote"** button.
2. A "Save As" dialog will appear — choose where to save the `.docx` file and confirm the filename.
3. The app will create the Word document and automatically save a `.mcmxq` project file alongside it.
4. Open the `.docx` in Microsoft Word to review it.

> **First-time Word open tip:** When you open the generated document, Word may ask if you want to update fields. Click **Yes**. This ensures the Table of Contents shows the correct page numbers.

---

## 4. OTTO Tool

### What it's for
Quoting an **OTTO 100 AMR** deployment — selecting the right hardware bundle for a customer.

### What you'll need
- Customer name, project name, location, and contact details
- Knowledge of which OTTO 100 configuration the customer needs

---

### Step 1 — Fill in Project Information
1. Click the **OTTO** tab at the top of the app.
2. Fill in **Project Name**, **Customer Name**, and **Customer Location**.
3. Fill in the contact fields (Presented By and Account Manager).

---

### Step 2 — Select Items from the OTTO Catalog
1. The main area shows a **catalog of OTTO 100 components** organized into categories.
2. Check the box next to each item the customer needs. Quantities adjust automatically.
3. You can use the tab filter buttons across the top to filter by category (e.g. Robots, Chargers, Accessories).
4. Items you select are automatically added to the BOM with correct part numbers and pricing.

---

### Step 3 — Add Any Custom Items
- If a customer needs something not in the preset catalog, click **"Add Custom Item"** and fill in the details manually.

---

### Step 4 — Save and Generate
1. Click **"Save OTTO Project"** to save a `.mcmxo` file.
2. Click **"Generate Document"** to produce the OTTO-branded Word document.

---

## 5. Thermal Imaging Estimator

### What it's for
Scheduling and pricing a **thermal imaging inspection** — how many techs, how many days, how they're getting there, and what it costs.

### What you'll need
- Number of technicians going on-site
- Each tech's travel time from their home location to the site (in hours)
- How each tech is getting there (driving or flying)
- Number of working days on-site
- Hours worked per day

---

### Step 1 — Open the Thermal Imaging Tab
Click the **"Thermal Imaging"** tab at the top of the app.

### Step 2 — Enter Project Information
Fill in the **Project Name**, **Customer Name**, and **Customer Location** fields at the top.

### Step 3 — Set the Number of Technicians
- Use the **Technicians** spinner to set how many people are going (1, 2, 3, etc.).
- A row appears for each technician.

### Step 4 — Configure Each Technician
For each tech row, fill in:
- **Travel Time** — how many hours it takes them to drive (or their proximity to the airport if flying). Enter the one-way travel time in hours (e.g. `3.5` for 3 hours 30 minutes).
- **Mode** — choose **Driving** or **Flying** from the dropdown.
- **Flight Cost** — if flying, enter the estimated round-trip airfare cost.
- **Mileage** — if driving, enter the round-trip mileage.

> **How hotel nights are determined:** If any tech has a drive time of 2 hours or more, or is flying, the tool automatically adds Travel In and Travel Out days and calculates hotel nights. Local crews (everyone driving under 2 hours) will only see the on-site work days.

### Step 5 — Set Work Schedule
- **Work Days** — how many days will the techs be on-site working.
- **Hours / Day** — how many hours per day they will work.

### Step 6 — Review the Schedule
- Click **"Generate Schedule"** (or the equivalent button).
- The schedule panel will show a day-by-day view: Travel In day, work days, and Travel Out day (if applicable).
- Review it to make sure it looks right.

### Step 7 — Set Pricing
- **Margin %** — enter your markup percentage.
- **Pricing Tier** — select the pricing tier that applies, or enter a **Custom Price** if needed.
- The **Cost Breakdown** panel shows the total cost and recommended resale price.

### Step 8 — Generate the Report
- Click **"Generate Thermal Report"** to produce the Word document.
- Choose where to save it.

---

## 6. IBE Estimator

### What it's for
Scheduling and pricing an **IBE (Infrared Building Envelope) inspection** — similar to Thermal Imaging but tailored to the IBE inspection workflow with panel counts and IBE-specific levels.

### What you'll need
- IBE level (1, 2, or 3 — determines the inspection intensity)
- Total number of panels to inspect
- Number of technicians
- Work schedule details
- Travel details for each tech

---

### Step 1 — Open the IBE Estimator Tab
Click the **"IBE Estimator"** tab at the top of the app.

### Step 2 — Enter Project Information
Fill in the **Project Name**, **Customer Name**, and **Customer Location** fields.

### Step 3 — Set IBE Parameters
- **IBE Level** — select the inspection level (1, 2, or 3). Higher levels are more thorough and take longer per panel.
- **Total Panels** — enter the number of panels to inspect.
- **Technicians** — set how many techs are going.
- **Hours / Day** — how many hours per day the techs will work.
- The tool will calculate the required **Work Days** based on the panel count, level, and team size.

### Step 4 — Configure Technician Travel
For each technician, enter:
- **Travel Time** (one-way, in hours)
- **Mode** — Driving or Flying
- **Flight Cost** (if flying)
- **Mileage** (if driving)

> Same hotel logic as Thermal Imaging: Travel In / Travel Out rows only appear if a tech drives 2+ hours or is flying.

### Step 5 — Review the Schedule
- Click **"Generate Schedule"**.
- Check the day-by-day schedule in the Outlook panel.
- If anything looks off (wrong number of days, wrong travel days), adjust the inputs and regenerate.

### Step 6 — Confirm the Schedule
- Once the schedule looks correct, click **"Confirm Schedule"**.
- This locks the schedule and enables the cost and export buttons.

### Step 7 — Review Costs
- In the **Cost Breakdown** section, review:
  - **Hotel Nights** — total nights across all techs
  - **Total Meals** — calculated from travel and work days
  - **Mileage** — total round-trip drive miles
  - **Margin %** — adjust if needed
  - **Recommended Resale** — your customer-facing price

### Step 8 — Export
- Click **"Export to Excel"** to save a `.mcmxc` cost sheet alongside an `.xlsx` Excel file.
- Or save the project as a `.mcmxq` file using the Save button.

---

## 7. Cost Estimator

### What it's for
An internal cost analysis tool. Use it when you need a detailed cost breakdown to check your numbers before finalizing a quote — especially useful for complex projects.

### What you'll need
- Estimated on-site labor hours
- Estimated travel hours
- Your internal hourly rate
- Risk / insurance percentage your company applies
- Your target margin percentage

---

### Step 1 — Open the Cost Estimator Tab
Click the **"Cost Estimator"** tab.

### Step 2 — Enter Project Info
Fill in the **Project** name and any relevant notes at the top.

### Step 3 — Enter Hours
- **On-Site Time (hrs)** — total hours your team will spend on-site.
- **Travel Time (hrs)** — total travel hours (both ways, all techs combined).

### Step 4 — Set Rates and Percentages
- **Hourly Rate ($)** — your internal cost rate per hour.
- **Risk / Insurance (%)** — the percentage your company adds for risk and insurance (often 5–15%).
- **Margin (%)** — your target profit margin percentage.

### Step 5 — Review the Summary
The **Cost Summary** panel at the bottom updates automatically as you type:
- **Cost Before Risk** — raw labor + travel cost
- **Risk Cost** — the dollar amount added for risk/insurance
- **Estimated Cost** — total internal cost
- **Recommended Resale** — the price you should charge the customer
- **Profit** — dollar profit at your set margin

### Step 6 — Export
- Click **"Generate Cost Sheet"** to save both a `.mcmxc` file (for re-importing later) and an `.xlsx` Excel file you can open and share.
- The Excel file mirrors the company cost report format.

---

## 8. Saving, Loading & Generating Documents

### Saving Your Work
| Tool | File Type | How to Save |
|------|-----------|-------------|
| Quote Generator | `.mcmxq` | File → Save MCMXQ, or Ctrl+S |
| OTTO Tool | `.mcmxo` | "Save OTTO Project" button |
| Thermal / IBE / Cost | `.mcmxq` or `.mcmxc` | Save button in the tab |

> **Save early, save often.** The app does not auto-save your work.

### Opening a Saved Project
1. Go to **File → Import MCMXQ** (for Quote Generator projects).
2. Browse to your saved `.mcmxq` file and open it.
3. Everything — all sections, pricing, notes, photos, version history — will be restored exactly as you left it.

### Generating a Word Document
1. Make sure all your information is filled in.
2. Click the **Generate** button for that tool (red button, usually labeled "Generate Quote" or "Generate Document").
3. A **Save As** dialog will appear. Choose a folder and confirm the filename.
4. The app saves the `.docx` and often a `.mcmxq` project file alongside it automatically.
5. Open the `.docx` in Word. If prompted to **update fields**, click **Yes** — this refreshes the Table of Contents page numbers.

---

## 9. Tips & Troubleshooting

### The app won't open on Windows ("Windows protected your PC")
This is normal for apps not downloaded from the Microsoft Store.
- Click **"More info"**, then **"Run anyway"**.
- You only need to do this the first time. After that, it opens normally.

### The app won't open on macOS ("unidentified developer")
- Don't double-click — instead, **right-click** the app → **Open** → **Open**.
- You only have to do this once.

### The Word document Table of Contents shows all page 1
- Open the document in Word, then press **Ctrl+A** (select all) followed by **F9** (update fields).
- Click **"Update entire table"** if prompted.
- Save the document.

### I accidentally closed the app without saving
Unfortunately there is no auto-recovery. Re-open the app and re-enter your work. To avoid this, save with **Ctrl+S** frequently.

### My project file won't open
Make sure you're using the correct file type for the tab:
- Quote Generator → `.mcmxq`
- OTTO Tool → `.mcmxo`
- Cost Estimator → `.mcmxc`

### The section arrows (▲▼) don't seem to be doing anything
Click the arrow once and give the screen a moment to redraw. If a section doesn't move, make sure it isn't already at the top (▲) or bottom (▼) of the list.

### I need to update from v1.1 to v2.0
See the [Upgrading from v1.1](README.md#upgrading-from-v11--v20) section in the README. In short: delete the old app folder (Windows) or app icon (macOS) and install the new one. Your `.mcmxq` files will work without any changes.

### Where are my saved project files?
Wherever you chose to save them. The app suggests a filename but you pick the folder. A good practice is to keep a dedicated folder like `Documents\MCMX Projects` and save everything there.

---

*MCMX Quote Generator v2.0 — McNaughton-McKay Electrical Company*
