# Ca$h Flow & CAPEX Project Tracking Platform

An open-source web application and analytical toolkit designed for real-time project management, CAPEX budget monitoring, and automated cash flow forecasting.

## 📌 Overview

Managing multi-project budgets across different domain areas often leads to fragmented data, delayed status reporting, and manual spreadsheet overhead. 

This repository provides a centralized, web-based environment to track project execution, monitor budget deviation, manage Purchase Orders (POs) / Service Acts, and project future cash flows with accuracy.

## 🚀 Key Features

- **Cash Flow Forecasting:** Automated cash flow projection engine based on payment schedules, PO milestones, and historical execution trends.
- **CAPEX & OPEX Budget Tracking:** Real-time variance monitoring (*Budgeted vs. Executed*) with alerts for funding limits and budget sweep events.
- **Interactive Dashboards:** Dynamic KPIs and status visualizers for tracking project progress, payment stages, and vendor contracts.
- **Automation & Integrations:** Connectors for syncing data between cloud spreadsheets (Google Sheets / Excel), web interfaces, and backend databases.

## 🛠 Tech Stack

- **Frontend / Web UI:** HTML5, CSS3, JavaScript / Google Apps Script Web Apps
- **Backend & Logic:** Python / Google Apps Script (GAS)
- **Data Analytics:** Pandas, OpenPyXL, JSON REST APIs
- **Integrations:** Cloud Storage, Webhook APIs, Automated Spreadsheet Workflows

## 📂 Repository Structure

```text
├── src/
│   ├── web/           # Web app UI templates and client scripts
│   ├── analytics/     # Cash flow projection and KPI calculation modules
│   └── integrations/  # Spreadsheet, API, and ETL automation scripts
├── docs/              # User guides and API documentation
└── README.md          # Project overview
