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

Manual:
================
Project Description — Last Week Lead Processor (Apps Script Automation)


This project automates the extraction, filtering and transformation of lead records received through email over the past week. The script reads daily CSV attachments from emails matching a specific subject line, identifies valid leads based on predefined rules, and aggregates all unique leads into a single structured output sheet.

The script processes one entire week at once. It determines the date range by calculating the last Monday and iterating through seven days. For each day, the script finds the email corresponding to that calendar date, extracts the latest CSV attachment, and applies a defined time window (previous day 6:00 PM to current day 6:00 PM). This ensures that leads created after 6 PM are counted as part of the next day's dataset.

After extraction, each CSV file undergoes filtering logic based on city, status, and creation timestamp. Leads are then deduplicated at two levels:

- Duplicate buyer IDs within the same day

- Duplicate buyer IDs already processed earlier in the week

- Once filtered and deduplicated, the script enriches the data by computing multiple fields, including:

- Sorted date values

- Lead creation and assignment timestamps

- Integer-based Excel date equivalents

- Combined key fields (date × buyer ID)
- These values help in downstream sorting, joining, and MIS-style comparisons.

## The final consolidated output is written into a dedicated sheet along with a fixed header row. Every row represents a unique lead for that week with all derived values precomputed.

The script is robust against missing emails, missing CSV files, or inconsistent date formats. It includes defensive parsing utilities, a date-windowing system, and supports multiple CSV date formats.

This automation reduces manual weekly processing and ensures consistent, error-free aggregation of time-sensitive lead data. It is suitable for any workflow that needs:
• Time-window based extraction
• Multi-day data stitching
• Duplicate-free record consolidation
• Enriched output for reporting or MIS pipelines
