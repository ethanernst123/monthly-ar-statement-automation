An Excel-based automation tool that streamlines the preparation, tracking, and distribution of monthly Accounts Receivable statements exported from Oracle ERP. Built using Power Query, VBA, and conditional formatting to eliminate repetitive manual work and reduce statement preparation time from up to a week down to minutes per account.

## What It Does

- Automatically cleans and formats raw Oracle ERP statement data on refresh
- Removes unnecessary columns, keeping only customer-relevant information
- Applies accounting/currency formatting to all dollar amounts
- Color codes rows by status — red for credits, yellow for past due, bold for due this month
- Separates Claim Investigations into a dedicated Claims tab automatically
- Calculates a dynamic Monthly Summary and Payment Schedule that updates based on the current date using TODAY()
- Exports a clean formatted .xlsx file automatically to a designated OneDrive inbox folder
- Automatically files exported statements into the correct account subfolders
- Opens Outlook email drafts addressed to the correct customer with the statement attached
- Tracks which accounts have been sent, which have no invoices due, and which still need processing — all color coded on one tab
- Displays a live filtered unsent accounts list so you always know who still needs a statement

## Tools & Techniques Used

- **Oracle ERP (AR Module)** — source system for raw statement exports processed by this template
- **Power Query** — automated data transformation, column filtering, and sorting on refresh
- **VBA (Macros)** — five automated buttons handling export, folder sorting, email drafting, account tracking, and monthly reset
- **Conditional Formatting** — dynamic row highlighting based on payment status with header protection using ROW()>1
- **SUMPRODUCT & SUMIF Formulas** — dynamic summary calculations that update automatically based on the current month and year using TODAY()
- **Excel Table References** — formula references that never shift regardless of data size

## File Structure

| Tab | Purpose |
| --- | --- |
| Instructions | Full user guide, button reference, formula explanations, and visual contents list |
| Raw Data | Paste raw Oracle statement export here before each account |
| Clean Statement | Formatted output, summary boxes, tracking buttons, and unsent accounts list |
| Claims | Claim Investigation rows separated automatically — only exports if claims exist |
| Statements Email List | Master account list with customer names, account numbers, contacts, and emails |

## How It Works

1. Open the template — a fresh copy opens automatically from the .xltm file
2. Paste the raw Oracle statement export into the Raw Data tab starting at A1
3. Click Export & Clear Data — enter the account number and the template automatically exports the file to OneDrive, marks the account green, refreshes the unsent list, clears Raw Data, and refreshes Power Query ready for the next account
4. For accounts with nothing due, click No Invoices Due — marks the row grey, no file created
5. Repeat for all accounts
6. Click Email Statements — opens Outlook drafts for each statement addressed to the correct customer
7. Click Sort Statements — automatically files all exported statements into the correct account folders
8. At the start of next month, click Reset Statements to clear all tracking colors and start fresh

## Button Guide

- **Export & Clear Data** — primary workflow button. Exports, tracks, clears, and refreshes in one click
- **Sort Statements** — files exported statements from the inbox into account subfolders automatically
- **Email Statements** — opens Outlook email drafts with statements attached, addressed to the correct customer
- **No Invoices Due** — marks accounts grey when nothing is due that month, no file created
- **Reset Statements** — clears all green and grey tracking colors at the start of each month

## Summary Box Calculations

The template includes two dynamic summary boxes that recalculate based on the current calendar month.

Monthly Summary includes Due thru EOM for positive invoices due this month that are not yet past due, Past Due for invoices with a positive Days Late value, Total Credits for the sum of all credit memos, and Net Balance as the combined total.

Payment Schedule breaks down what is due by the 10th, 20th, and end of month to help customers plan payment runs.

## Skills Demonstrated

- Excel automation without third-party tools
- Power Query data transformation and dynamic sorting
- VBA macro development for multi-step workflow automation
- Oracle ERP data processing and formatting
- Financial data presentation and conditional formatting
- End-to-end process automation including file management and email distribution
- Account tracking and status management across a team workflow
- Process improvement and workflow documentation with full in-file SOP
