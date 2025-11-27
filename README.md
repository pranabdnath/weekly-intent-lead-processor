# weekly-intent-lead-processor
This project automates the extraction, filtering and transformation of lead records received through email over the past week. The script reads daily CSV attachments from emails matching a specific subject line, identifies valid leads based on predefined rules, and aggregates all unique leads into a single structured output sheet.


# Weekly Intent Lead Processor

This project automates the extraction, filtering, de-duplication, and aggregation of weekly lead data using Google Apps Script.

## Purpose

Many organisations receive daily CSV reports through email. This script processes last week's data, applies status and city filters, removes duplicates across multiple days, and generates a clean unified output.

## What the script does

1. Calculates last week's Monday.
2. Collects seven days of data (Monday to Sunday).
3. Searches for emails matching the specified subject line.
4. Extracts the latest CSV attachment for each day.
5. Applies filters based on status, city, and the 6 PM to 6 PM processing window.
6. Removes duplicate buyers for both a single day and the whole week.
7. Computes derived MIS and CRM fields.
8. Writes cleaned rows into `LastWeekProcessedLeads`.

## Script location

The complete code is stored in:

`/apps-script/processLastWeekLeads.gs`

Copy and paste the file into your Google Apps Script environment.

## Requirements

- Gmail access for reading incoming files
- Google Drive container for downloaded reports
- Google Sheets for output storage

