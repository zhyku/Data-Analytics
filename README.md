📊 Revenue Analytics Dashboard with Transactional Data Engineering

## Overview
Built a formula-driven Excel/Google Sheets model to clean, standardize, and analyze a high-variance transactional dataset. The project focuses on resolving data inconsistencies, engineering reliable financial metrics, and generating accurate revenue and customer lifecycle insights.

---

## Key Features

### Data Engineering & Validation
- Standardized inconsistent date formats (mixed `YYYY-MM` and `DD-MM-YYYY`) into a unified monthly timeline.
- Resolved duplicate transaction IDs and handled negative or invalid quantities.
- Normalized case-sensitive fields (Customer, Account Manager, Billing Cycle).

### Financial Modeling & Revenue Analytics
- Built dynamic calculations for Net Amount, GST (18%), and Final Amount across tax-inclusive and tax-exclusive cases.
- Developed a fully automated Month-over-Month (MoM) growth model.
- Implemented logic to correctly aggregate fragmented month-end transactions, preventing double-counting and broken trend lines.

### Customer Lifecycle Insights
- Classified transactions into Active, Churned, Paused, and Refunded states.
- Identified “Hybrid” customers (mixed lifecycle states) to surface retention risks and billing inconsistencies.

### Regional Performance Intelligence
- Built pivot-based summaries for:
  - Total revenue by region
  - Active transaction volume
  - Average revenue per active transaction

---

## Technical Toolkit
- **Functions:** ARRAYFORMULA, SUMIFS, INDEX/MATCH, FILTER, SORT, UNIQUE  
- **Logic Handling:** Nested IFs, array matching, and error handling (IFERROR)  
- **Data Cleaning:** TRIM, UPPER, text normalization  
- **Automation:** Dynamic arrays enabling fully automated, manual-free reporting  

---

## Key Formula Logic

**1. Monthly Revenue (Filtered Aggregation)**  
Isolates domestic active revenue while excluding churned or non-INR transactions.

```excel
=SUMIFS(FinalAmount, Status, "Active", Currency, "INR", Month, A2)
```

**2. MoM Growth (Month-over-Month Change)**  
Calculates growth between consecutive months with safe handling for missing or zero previous values.

```excel
=IFERROR((Current_Month_Total - Previous_Month_Total) / Previous_Month_Total, 0)
```

---

## Outcome
Converted a noisy transactional dataset into a structured analytical model that:
- Ensures consistent and accurate financial calculations  
- Reveals true MoM growth trends  
- Highlights churn risks and inconsistent customer behavior  
- Enables scalable, formula-driven reporting  

---

## Project Files
- `Transactions_Analysis_Dashboard.xlsx` – Core analytical model  
- `/screenshots` – Visual breakdown of sub-sheets, formula architecture, and dashboards
