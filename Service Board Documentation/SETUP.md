# Service Board Imports - Setup Guide

## Overview
This system automatically processes Excel exports from F2 Service Board software. Files are exported to your Desktop, then manually moved to the Google Drive folder, where they are automatically processed.

## Setup Instructions

### 1. Google Drive Folder
- **Folder ID**: `1nUy7lWNr1BVCAxyLsnFASCTszQpjgEnd`
- **URL**: https://drive.google.com/drive/u/0/folders/1nUy7lWNr1BVCAxyLsnFASCTszQpjgEnd
- **Local Path**: `G:\My Drive\Service Board Imports`

### 2. Apps Script Setup

#### Step 1: Create a New Apps Script Project
1. Go to https://script.google.com
2. Click "New Project"
3. Name it "Service Board Imports"

#### Step 2: Add the Script Files
1. Copy the contents of `Process Service Board Import.js` into the script editor
2. Copy the contents of `Custom Menu.js` into the script editor (or add the `onOpen()` function to your existing script)

#### Step 3: Enable Required APIs
1. In the Apps Script editor, go to **Services** (left sidebar)
2. Click **+** to add a service
3. Add **Google Drive API** (if not already enabled)
4. The script uses `Drive.Files.copy()` which requires the Drive API v2

#### Step 4: Authorize the Script
1. Run the `onOpen()` function once (or just open a Google Sheet - it will run automatically)
2. Authorize the required permissions when prompted

### 3. Usage

#### Manual Processing
1. Export file from F2 Service Board to your Desktop
2. Copy/drag the file into the Google Drive folder: `G:\My Drive\Service Board Imports`
3. Open any Google Sheet (or the one where you installed the menu)
4. Go to **Service Board** → **Process Imports**
5. Check the execution log for results

#### Automated Processing (Time-Driven Trigger)
1. In Apps Script editor, go to **Triggers** (clock icon, left sidebar)
2. Click **+ Add Trigger**
3. Configure:
   - **Function**: `processServiceBoardImports`
   - **Event source**: Time-driven
   - **Type**: Day timer (or Hour timer for more frequent checks)
   - **Time**: Choose your preferred time (e.g., daily at 9:00 AM)
4. Click **Save**

### 4. File Naming
Files must match this pattern:
- Format: `Service [date] at [time].xlsx`
- Example: `Service 2025-12-03 at 3.41.39 PM.xlsx`
- Pattern: `Service YYYY-MM-DD at H.MM.SS AM/PM.xlsx`

### 5. Processing Behavior
- Files are automatically detected if they match the naming pattern
- Each file is converted to Google Sheets for processing
- Files are tracked to prevent duplicate processing
- Temporary converted sheets are automatically deleted after processing
- Original Excel files remain in the folder

### 6. Customization
The `analyzeServiceData()` function in `Process Service Board Import.js` is where you can add custom logic:
- Compare with previous imports
- Update master databases
- Generate reports
- Send notifications
- Calculate metrics

### 7. Troubleshooting

#### Files Not Being Processed
- Check that files match the exact naming pattern
- Verify the file is in the correct Google Drive folder
- Check the execution log for errors
- Use **Service Board** → **View Processed Files** to see what's been processed

#### Reset Processed Files List
If you need to reprocess files:
- Go to **Service Board** → **Reset Processed Files List**
- This clears the tracking list (files will be processed again)

#### Check Execution Logs
1. In Apps Script editor, go to **Executions** (left sidebar)
2. Click on any execution to see detailed logs
3. Look for error messages or warnings

## File Structure
```
Service Board Imports/
├── Process Service Board Import.js    # Main processing script
├── Custom Menu.js                      # Menu creation script
└── [Excel files from F2]              # Files to be processed
```

## Next Steps
1. Customize the `analyzeServiceData()` function for your specific needs
2. Set up time-driven triggers for automated processing
3. Configure where processed data should be stored/analyzed
4. Add any additional analysis or reporting features

