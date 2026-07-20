# Mutual Fund Portfolio Tracker (Excel)

> An advanced, fully offline Excel tracker for Indian Mutual Fund family portfolios. 

Manage your household wealth without compromising your data. This tracker imports parsed CAS data to automate performance tracking and tax calculations, keeping your financial life entirely local.

**Key Features:**
* **PAN-Level Performance:** Automates PAN-level Current and All-Time XIRR, alongside All-Time and Current Gains.
* **Scheme-Level Insights:** Tracks PAN-level MF schemes, detailing XIRR, invested amount, unrealized gains, and current value.
![Mutual Funds Portfolio Summary](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Mutual%20Funds%20Portfolio%20Summary.png)
* **Tax Harvesting Engine:** Features a strict FIFO tax harvesting engine for the ₹1.25L exemption.
![Tax Harvesting](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Tax%20Harvesting.png)
* **Redemption Calculator:** Simulates redemptions with detailed Equity/Debt and LTCG/STCG splits.
![Redemption Calculator](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Redemption%20Calculator.png)
* **Family Portfolio Summary:** Provides a consolidated, high-level overview of your entire household's wealth.
![Family Portfolio Summary](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Family%20Portfolio%20Summary.png)
* **Portfolio Allocation Summary:** Provides a consolidated and per PAN overview of Equity & Debt allocation as per defined rules.
* **Capital Gains Statements:** Generates FY-wise and PAN-wise capital gains statements broken down by Equity/Debt and LTCG/STCG.
![Capital Gains](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Capital%20Gains.png)
* **Folio Management:** Lists and organizes all your folios PAN-wise.
![PAN Mapper & Folios List](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/PAN%20Mapper%20%26%20Folios%20List.png)
* **Family Goal Tracking:** Track financial goals CAS-wide. Seamlessly map multiple mutual fund schemes and folios—even across different family members' PANs—to a single unified goal.
![Goals](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Goals.png)
* **Mutual Fund Calculators:** Lumpsum/SIP Calculator, SWP Calculator and IDCW Calculator.
![Lumpsum SIP SWP IDCW Calculators](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/Lumpsum%20SIP%20SWP%20IDCW%20Calculators.png)

---

## Why This?

For years, Kuvera was my go-to platform for managing and tracking Mutual Funds. It was seamless, particularly for managing a consolidated family portfolio. However, following its acquisition by CRED, the platform removed the ability to manage multiple family accounts under a single mobile number. 

As someone who strongly prioritizes digital privacy, linking multiple personal numbers or handing over granular financial data to third-party aggregators was not an option. I needed an alternative that was completely offline, private, and easy to use. 

My search for a replacement revealed a frustrating gap in the market:
* **The Cloud Trackers:** Too invasive, demanding excessive phone number linkages and cloud syncing.
* **The Self-Hosted Route:** Overly complex, requiring dedicated servers, Docker containers, and unnecessary technical maintenance just to track numbers.
* **The Third-Party Apps:** Closed-source applications with questionable data handling practices and vague terms of service.

To bypass these privacy issues entirely, I strongly recommend doing transactions directly via MFCentral, CAMS, KFinTech, other official RTAs, or the AMCs' official websites. But while transacting directly is secure, tracking everything across multiple family members without a central aggregator becomes a challenge. 

I needed something reliable, comprehensive, and fully under my control. 

**The Solution:** Good old Microsoft Excel. 

By building this tracker in a simple, everyday spreadsheet, the data stays exactly where it belongs—on your local machine. There are no servers, no hidden tracking, and no forced updates. It is a transparent, robust, and privacy-first way to manage investments without compromising on usability. 

---

## Privacy by Design: How Your Data Stays Yours

When I set out to build this, total data isolation was my absolute, non-negotiable priority. Here is exactly how this tracker protects your financial data:

* **Auditable CAS Parsing:** To import your transactions, this project relies on a script that generates a CSV from your Consolidated Account Statement (CAS) PDF. I use a fork of the popular open-source tool [codereverser/casparser](https://github.com/codereverser/casparser). My fork is hosted at [mehulboricha/casparser](https://github.com/mehulboricha/casparser). While both repositories are currently identical, I deliberately use a fork so I can personally vet the code for privacy and security updates before running any script that touches my financial documents. 
* **Controlled, On-Demand NAV Fetching:** Live data trackers often leak your IP and portfolio metadata through constant background syncing. Instead, NAV fetching here is strictly manual. You run a single, transparent command that fetches the latest NAVs for your specific schemes and drops them directly into the `Current NAVs` sheet. You control exactly when your machine connects to the internet.
* **100% Offline Excel Sandbox:** The Excel file itself is a complete walled garden. There are **zero external network calls** made from within the workbook. No background web queries, no hidden APIs, and no telemetry. Once your data is in the sheet, absolutely everything is calculated locally using standard Excel formulas, internal functions, and standalone VBScript.
* **Transparent Macros (VBA):** When you open the file, Excel will show a standard security warning: *"This workbook contains macros... Macros may contain viruses..."* **You must click Enable Macros** for the FIFO tax harvesting engine and portfolio calculations to function. Because privacy and security are paramount, you are highly encouraged to press `Alt + F11` to open the VBA editor and inspect the code yourself. You will see it only performs local data processing and makes zero external connections.

---

## How to Use

Follow these steps to extract your CAS data, import it into the tracker, and fetch the latest NAVs.

### Phase 1: Preparing Your Computer

#### -> For Mac (Continue Directly to Phase 2)

#### -> For Windows

* You must have Python installed and added to your system's PATH (so the python command works in PowerShell).
* You must run the next command inside PowerShell, not the legacy Command Prompt.

#### -> For Linux (X11)

* Install XClip via your package manager, e.g., ```sudo apt install xclip```

#### -> For Linux (Wayland)

* Install wl-clipboard via your package manager, e.g., ```sudo apt install wl-clipboard```

### Phase 2: Parse Your CAS PDF
First, convert your encrypted CAS PDF into a clean CSV format using the local parser.

1. Open your terminal and install the parser:

```
bash
git clone [https://github.com/mehulboricha/casparser](https://github.com/mehulboricha/casparser)
cd casparser
pipx install .
```

2. Run the parser on your downloaded CAS PDF (replace 123456 with your actual PDF password and cas.pdf with your filename):

```
casparser -o cas.csv -p "123456" cas.pdf
```

### Phase 3: Import Data into Excel

1. Now, bring the parsed data into the offline tracker.
2. Open the CAS sheet in the Excel workbook.
3. Press Ctrl + A to select all, then press Delete to clear any existing data.
4. Open the newly generated cas.csv file, press Ctrl + A (select all), and then Ctrl + C (copy).
5. Return to the CAS sheet in your Excel workbook and paste (Ctrl + V) the copied data starting from cell A1.

### Phase 4: Update Current NAVs

#### Method 1 - Using Script in Automate (Recommended)

1. Go to Automate Tab.
2. Click New Script -> Create in Code Editor.
3. Paste the following code:

```
// Define the shape of the JSON so TypeScript doesn't complain
interface MFResponse {
  meta?: {
    scheme_code?: string | number;
  };
  data?: {
    date?: string;
  }[];
}

async function main(workbook: ExcelScript.Workbook) {
  // 1. Connect to the specific sheet
  let sheet = workbook.getWorksheet("Current NAVs");
  if (!sheet) {
    console.log("Could not find a sheet named 'Current NAVs'");
    return;
  }

  // 2. Get today's date formatted as dd-mm-yyyy (matching the API format)
  let today = new Date();
  let dd = String(today.getDate()).padStart(2, '0');
  let mm = String(today.getMonth() + 1).padStart(2, '0');
  let yyyy = today.getFullYear();
  let todayStr = `${dd}-${mm}-${yyyy}`;

  // 3. Find the last used row in the sheet
  let usedRange = sheet.getUsedRange();
  let rowCount = usedRange.getRowCount();

  console.log(`Checking up to row ${rowCount}...`);

  // 4. Loop through each row starting at Row 2 (Index 1)
  for (let i = 1; i < rowCount; i++) {
    let bVal: string = sheet.getCell(i, 1).getValue().toString().trim(); // Column B: Scheme Code

    // If Column B has a value, process it
    if (bVal !== "") {

      // Explicitly declare types
      let dVal: string | number | boolean = sheet.getCell(i, 3).getValue(); // Column D: Total Units
      let eVal: string = sheet.getCell(i, 4).getValue().toString().trim(); // Column E: Existing JSON
      let gText: string = sheet.getCell(i, 6).getText().trim(); // Column G: Date (Index 6)

      // --- OPTIMIZATION CHECK START --- //

      // Condition A: Is Total Units exactly 0? 
      let isUnitsZero: boolean = (dVal === 0 || dVal === "0" || dVal === 0.0);

      let hasMatchingScheme: boolean = false;
      let isDateToday: boolean = false;

      // Condition B: Does Column G text match today's date?
      if (gText === todayStr || gText === `${mm}/${dd}/${yyyy}` || gText === `${today.getMonth() + 1}/${today.getDate()}/${yyyy}`) {
        isDateToday = true;
      }

      // Condition C: JSON Validation & Fallback Date Check
      if (eVal !== "") {
        try {
          // Cast the parsed JSON to our custom interface
          let parsedJson = JSON.parse(eVal) as MFResponse;

          if (parsedJson.meta && parsedJson.meta.scheme_code == bVal) {
            hasMatchingScheme = true;
          }

          // Fallback: If Column G is formatted weirdly by Excel, reliably check the JSON date
          if (parsedJson.data && parsedJson.data[0] && parsedJson.data[0].date === todayStr) {
            isDateToday = true;
          }
        } catch {
          // Fails safely if JSON is invalid or blank
        }
      }

      // The Skip Action: Skip if (Units are 0 and matched) OR (Date is already today)
      if ((isUnitsZero && hasMatchingScheme) || isDateToday) {
        console.log(`Row ${i + 1} Skipped: Units are 0 or Date is already updated today.`);
        continue;
      }

      // --- OPTIMIZATION CHECK END --- //


      // If it didn't skip, construct the URL and fetch the API
      let url: string = `https://api.mfapi.in/mf/${bVal}/latest`;

      try {
        let response = await fetch(url);

        if (response.ok) {
          let jsonText: string = await response.text();
          sheet.getCell(i, 4).setValue(jsonText);
        } else {
          sheet.getCell(i, 4).setValue("API Error: " + response.status);
        }
      } catch {
        sheet.getCell(i, 4).setValue("Fetch Failed");
      }

    } else {
      // --- STALE DATA CLEANUP --- //
      // If Column B is empty, ensure Column E is also empty
      let eVal: string = sheet.getCell(i, 4).getValue().toString().trim();
      if (eVal !== "") {
        sheet.getCell(i, 4).setValue("");
        console.log(`Row ${i + 1} Cleared: Scheme code in Column B was removed.`);
      }
    }
  }

  console.log("Done fetching API data!");
}
```

#### Method 2 - Using API Call in Python

1. Since the Excel file makes no network calls, we use a clever terminal command to fetch live NAVs securely using your clipboard.
2. Go to the Current NAVs sheet in your Excel workbook.
3. Copy the exact command below, paste it into your Terminal, but DO NOT press Enter yet:<br/><br/>

<b>For Mac</b>
```
pbpaste | tail -n +2 | python3 -c 'import sys,urllib.request; print("\n".join(urllib.request.urlopen(f"[https://api.mfapi.in/mf/](https://api.mfapi.in/mf/){c.strip()}/latest",timeout=30).read().decode().replace("\n","") if c.strip().isdigit() else "" for c in sys.stdin))' | pbcopy
```

<b>For Windows (PowerShell Only)</b>
```
(Get-Clipboard | Select-Object -Skip 1) | python -c "import sys,urllib.request; print('\n'.join(urllib.request.urlopen(f'https://api.mfapi.in/mf/{c.strip()}/latest',timeout=30).read().decode().replace('\n','') if c.strip().isdigit() else '' for c in sys.stdin))" | Set-Clipboard
```

<b>For Linux (X11)</b>
```
xclip -selection clipboard -o | tail -n +2 | python3 -c 'import sys,urllib.request; print("\n".join(urllib.request.urlopen(f"https://api.mfapi.in/mf/{c.strip()}/latest",timeout=30).read().decode().replace("\n","") if c.strip().isdigit() else "" for c in sys.stdin))' | xclip -selection clipboard
```

4. Go back to Excel and select Column B in the Current NAVs sheet, then copy it (Cmd + C). (This loads your scheme codes into the clipboard).
5. Return to your Terminal and press Enter to run the script.
6. Wait a few seconds for the script to finish processing (it will automatically copy the fetched NAV data back to your clipboard).
7. Return to the Current NAVs sheet, click on cell E2, and paste (Cmd + V).

### Phase 5: Update FIFO_Split_Ledger

Everytime you update CAS, if you want accurate Capital Gains calculation then go to "Capital Gains" sheet and Click on "Click to Update Ledger"

### Phase 6: Customize Your Excel Sheet (One-Time Setup)

To make future updates as frictionless as possible, I have included the essential terminal commands directly inside the Excel workbook on the **Current NAVs** sheet. This saves you from having to revisit this README every time you want to update your portfolio!

Take a minute to customize these cells for your specific setup:

**1. Update the Casparser Command**
On the right side of the sheet, you will find the default CAS parsing command:
```
casparser -o cas.csv -p "123456" cas.pdf
```
Double-click this cell and replace `"123456"` with your actual CAS PDF password, and `cas.pdf` with the standard filename you use. Now, whenever you get a new statement, your exact command is ready to copy.

**2. Update the Dynamic NAV Command**
By default, the spreadsheet contains the Mac terminal command for fetching live NAVs. If you are using Windows or Linux:
* Copy your OS-specific command from **Phase 4** above.
* Double-click the cell containing the Mac command in the Excel sheet.
* Delete the existing text and paste your new command.

Save the workbook (`Ctrl + S` or `Cmd + S`). Now, everything you need to track and update your wealth is fully self-contained within your local Excel file!

---

## How to Get CAS via CAMS or KFin Tech

1. Go to [CAMS CAS Download Page](https://www.camsonline.com/Investors/Statements/Consolidated-Account-Statement) or [KFin Tech CAS Download Page](https://mfs.kfintech.com/investor/General/ConsolidatedAccountStatement)
2. On CAMS you will see something like this.
   ![CAMS CAS](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/CAMS%20CAS.png)
3. On KFin Tech you will see something like this.
   ![KFin Tech CAS](https://github.com/mehulboricha/Mutual-Fund-Portfolio-Tracker-Excel/blob/main/Screenshots/KFin%20Tech%20CAS.png)
4. Enter the details as per following:
   Statement Type: Detailed
   Period: Select Date before you started investing upto TODAY
   Folio Listing: Select With Zero Balance Folios
   Email: Entered Email Registered with AMCs
   PAN: Keep Empty if you have multiple PANs associated with same details.
   Password: Password for the generated PDF.
5. You will receive CAS via email shortly.

---

## Demo CSV Generation Steps

If you want to test the tracker before importing your actual financial data, you can generate a dummy CAS dataset using Google Colab.

1. Go to [Google Colab](https://colab.research.google.com/) and click **New Notebook**.
2. Paste the Python code below into the first cell.
3. Click the **Play** button (or press `Shift + Enter`) to run the code.
4. Once the script outputs *"Successfully generated [X] rows!"*, click the **Folder icon** on the far-left sidebar.
5. Locate `demo_cas_data.csv`, click the three dots next to it, and select **Download**.

```
python
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Configurations based on your inputs
pans = ["AAAAA1111A", "BBBBB2222B", "CCCCC3333C", "DDDDD4444D"]
start_date = datetime(2010, 1, 1)
end_date = datetime(2026, 7, 1)

schemes = [
    {"amc": "HDFC Mutual Fund", "scheme": "HDFC Flexi Cap Fund - Direct Plan - Growth", "amfi": "119062", "cat": "Equity"},
    {"amc": "Parag Parikh Mutual Fund", "scheme": "Parag Parikh Flexi Cap Fund - Direct Growth", "amfi": "122639", "cat": "Equity"},
    {"amc": "SBI Mutual Fund", "scheme": "SBI Liquid Fund Direct Growth", "amfi": "119598", "cat": "Debt"},
    {"amc": "ICICI Prudential Mutual Fund", "scheme": "ICICI Pru Equity & Debt Fund - Direct Growth", "amfi": "120664", "cat": "Equity"},
    {"amc": "Axis Mutual Fund", "scheme": "Axis ELSS Tax Saver Fund - Direct Growth", "amfi": "112323", "cat": "Equity"} # Sold off scenario
]

rows = []

# Generate data for each PAN
for pan in pans:
    for idx, fund in enumerate(schemes):
        folio = f"{pan[:4]}999{idx}"
        current_date = start_date + timedelta(days=np.random.randint(0, 365))
        nav = 20.0
        balance = 0.0
        
        # Determine if this fund should be completely sold off to test All-Time XIRR
        sold_off = True if "Axis" in fund["scheme"] else False
        
        while current_date < end_date:
            # Random NAV growth
            nav *= (1 + np.random.uniform(-0.02, 0.03)) 
            
            # Monthly SIP Simulation
            amount = 5000.0
            units = amount / nav
            balance += units
            
            rows.append([fund["amc"], folio, pan, fund["scheme"], "DIRECT", "CAMS", fund["amfi"], 
                         current_date.strftime("%d-%b-%Y"), "Purchase - SIP", amount, round(units, 3), 
                         round(nav, 4), round(balance, 3), "PURCHASE_SIP", ""])
            
            # Random STAMP_DUTY / STT_TAX (Zero Unit Rows)
            if np.random.rand() > 0.8:
                tax_type = np.random.choice(["STAMP_DUTY", "STT_TAX"])
                tax_amt = round(amount * 0.0005, 2)
                rows.append([fund["amc"], folio, pan, fund["scheme"], "DIRECT", "CAMS", fund["amfi"], 
                             current_date.strftime("%d-%b-%Y"), tax_type.replace("_", " "), tax_amt, 0.0, 
                             "", round(balance, 3), tax_type, ""])
            
            # Random DIVIDEND_PAYOUT
            if np.random.rand() > 0.95 and fund["cat"] == "Equity":
                div_amt = round(balance * 0.5, 2)
                rows.append([fund["amc"], folio, pan, fund["scheme"], "DIRECT", "CAMS", fund["amfi"], 
                             current_date.strftime("%d-%b-%Y"), "Dividend Paid", div_amt, 0.0, 
                             "", round(balance, 3), "DIVIDEND_PAYOUT", div_amt])
                
            # Random Partial REDEMPTION
            if np.random.rand() > 0.97 and balance > 50:
                red_units = balance * 0.2
                red_amt = red_units * nav
                balance -= red_units
                rows.append([fund["amc"], folio, pan, fund["scheme"], "DIRECT", "CAMS", fund["amfi"], 
                             current_date.strftime("%d-%b-%Y"), "Redemption", -round(red_amt, 2), -round(red_units, 3), 
                             round(nav, 4), round(balance, 3), "REDEMPTION", ""])

            # Advance 1 month
            current_date += timedelta(days=30)
            
            # Trigger complete sell-off if applicable
            if sold_off and current_date.year == 2022 and balance > 0:
                red_units = balance
                red_amt = red_units * nav
                balance = 0.0
                rows.append([fund["amc"], folio, pan, fund["scheme"], "DIRECT", "CAMS", fund["amfi"], 
                             current_date.strftime("%d-%b-%Y"), "Full Redemption", -round(red_amt, 2), -round(red_units, 3), 
                             round(nav, 4), round(balance, 3), "REDEMPTION", ""])
                break # Stop generating records for this sold-off fund

df = pd.DataFrame(rows, columns=["amc", "folio", "pan", "scheme", "advisor", "rta", "amfi", "date", 
                                 "description", "amount", "units", "nav", "balance", "type", "dividend"])

# Output to CSV
df.to_csv("demo_cas_data.csv", index=False)
print(f"Successfully generated {len(df)} rows!")
```

---

## Community Contributions & Acknowledgements

This tracker was built to solve a personal pain point. I originally developed and tested all formulas, FIFO tax calculations, and VBA scripts using my own family's CAS data. To bring this to life, I relied on my personal knowledge, various online financial resources, and a bit of heavy lifting from AI assistants (including Google Gemini, ChatGPT, and GitHub Copilot).

While I have done my absolute best to ensure the math, XIRR calculations, and tax harvesting engines are fully accurate, mutual fund portfolios can have unique edge cases. **There may be mistakes or unhandled scenarios.** I would be thrilled if you helped make this better! Here is how you can contribute to the project:

* **Bug Reports:** If you spot a calculation error, a broken formula, or a macro crash, please open an **Issue** on GitHub. *(Important: Please ensure you completely anonymize your PAN, Folio numbers, and amounts before sharing any screenshots or data!)*
* **Code & Formula Optimization:** If you are an Excel wizard and know a more efficient formula or a cleaner way to write the VBA scripts (while maintaining the zero-external-call rule), Pull Requests are incredibly welcome.
* **Feature Requests:** Want a new metric, a different chart, or a new summary view? Let's discuss it in the Issues section.
* **Documentation:** Found a typo in this README or know a better way to explain a setup step? Feel free to submit a fix.

Whether it is pointing out a tiny typo or optimizing the VBA engine, all feedback and contributions are highly appreciated!
