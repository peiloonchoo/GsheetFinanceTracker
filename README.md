# GsheetFinanceTracker

A personal finance tracker built as a Google Apps Script for Google Sheets. Track bank accounts, investments, and loans all in one place.

## Features

- **Accounts** — Create multi-currency accounts (MYR, USD, SGD, etc.) and log income/expense transactions with running balance
- **Stocks** — Portfolio tracking with live price refresh (auto every 5 min) and sell/buy history
- **Mutual Funds** — Track unit trust holdings with NAV updates via FSMOne and FIMM (daily auto-refresh at 7 AM/7 PM)
- **Crypto** — Cryptocurrency portfolio with buy/sell tracking
- **Gold** — Physical gold holdings tracker
- **Loans & Debts** — Track money lent and borrowed, with repayment recording
- **Interest** — Apply interest rates to accounts, with optional daily auto-posting
- **Retirement** — Configurable retirement portfolio projection sheet
- **Dashboard** — Consolidated net worth view across all accounts and portfolios

## Setup

1. Open a new Google Sheet
2. Go to **Extensions → Apps Script**
3. Paste the contents of `code.gs` and save
4. Reload the spreadsheet — a **💰 Finance** menu will appear
5. Use the menu to create accounts, portfolios, and sheets as needed

## Usage

All features are accessible via the **💰 Finance** custom menu after setup. Start by creating an account, then add your investment sheets as needed.

## Requirements

- Google account with access to Google Sheets and Apps Script
- Internet access (for live stock prices and mutual fund NAV fetching)