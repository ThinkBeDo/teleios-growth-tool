# Teleios Growth Tool Automation

🏥 **Automated county demographic data processing for Alex Ledford at Teleios**

## Overview

This web application automates the manual copy-paste process Alex performs every 1-2 weeks when updating county demographic data in the Growth Tool Excel workbook. 

**Time Savings:** Reduces 15min-1hr manual process to 2-3 minutes!

## Features

- ✅ Upload multiple county Excel files + main workbook
- ✅ Automatically extract data from County Trend sheets  
- ✅ Map data to Raw sheet with proper field mapping
- ✅ Generate updated workbook for download
- ✅ Handle errors gracefully with user feedback

## Data Mapping

From County Trend sheets (rows 9+) to Raw sheet:
- Medicare Enrollment (Col C) → Medicare Beneficiaries (Col E)
- Resident Deaths (Col E) → Medicare Deaths (Col G)  
- Patients Served (Col I) → Hospice Unduplicated Beneficiaries (Col K)
- Hospice Deaths (Col G) → Hospice Deaths (Col I)

## Local Development

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open http://localhost:5000

## Deployment

This application is configured for Railway deployment with:
- `Procfile` for web process
- `runtime.txt` for Python version
- `requirements.txt` for dependencies

## Usage

1. Navigate to the web interface
2. Upload your main Growth Tool workbook
3. Select all county Excel files to process
4. Click "Process Files"
5. Download the updated workbook

## Project Details

- **Client:** Teleios (Chris Comeaux's company)
- **End User:** Alex Ledford
- **Project Key:** TELEIOS-GROWTH-AUTOMATION-KEY
- **Built:** August 2025
