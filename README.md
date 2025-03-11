# Extract Hyperlinks from Google Sheets

## Overview
This Google Apps Script automatically extracts hyperlinks from a column containing company names with embedded links and places them in an adjacent column. It works with both:
- **Hyperlinked rich text values**
- **HYPERLINK formulas**

## Features
- Extracts hyperlinks from column A and places them in column B.
- Works for both **HYPERLINK formulas** and **rich text hyperlinks**.
- Can be manually executed or set up to run automatically via a trigger.

## How to Use
### 1. Add the Script to Your Google Sheet
1. Open your **Google Sheet**.
2. Click on **Extensions** > **Apps Script**.
3. Create a new Script file.
4. Paste the code into it.

### 2. How to Run the Script
#### Run Manually
1. Open **Google Sheets**.
2. Click on **Extensions** > **Apps Script**.
3. In the script editor, select the function `extractHyperlinks`.
4. Click **Run** ▶️.
5. Grant necessary permissions when prompted.
6. The extracted hyperlinks will appear in column **B**.

#### Automate the Process
To automatically extract hyperlinks whenever the sheet is edited:
1. Go to **Apps Script Editor**.
2. Click on the **Triggers (⏰ clock icon)** in the left panel.
3. Click **+ Add Trigger**.
4. Configure the trigger:
   - **Choose function** → `extractHyperlinks`
   - **Choose event source** → `From spreadsheet`
   - **Choose event type** → `On edit`
5. Click **Save**.

## Deployment Options
If you want to deploy it for broader use:
- **Installable Trigger**: Run automatically on edits.
- **Add-on**: Publish for others in Google Workspace Marketplace.

## License
This project is licensed under the **MIT License**.
