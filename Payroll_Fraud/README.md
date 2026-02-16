# 🕵️‍♂️ Payroll Fraud Detection & Audit Tool

## **Project Overview**
This tool was engineered to automate the forensic auditing of payroll datasets to identify shift overlaps, "Midnight Crossover" discrepancies, and compliance violations. By migrating from a manual review process to this automated C# engine, I addressed systemic issues that previously led to financial penalties and optimized data integrity across high-volume datasets.

## **Key Forensic Logic**
The system processes raw Excel data through a custom **Temporal Attribution Engine** to map shifts into a 24-hour matrix:

* **Midnight Crossover Adjustment**: Automatically detects and links shifts spanning across two calendar days (e.g., 22:00 to 06:00) to ensure chronological continuity.
* **16.5-Hour Safety Guardrail**: Flags any single shift or combination of shifts exceeding the 16.5-hour safety limit in **RED** on the final dashboard.
* **Conflict Detection**: Utilizes logic-based auditing to identify "Primary Duty" and "Secondary Duty" overlaps occurring within the same hour, flagging potential fraud for manual review.

## **Technical Stack**
* **Language**: C# (.NET Framework).
* **Data Ingestion**: OLEDB (JET 4.0) and SpreadsheetLight for high-fidelity Excel parsing.
* **Regex Engine**: Custom regular expressions for sanitizing "dirty" legacy system exports and identifying time-series patterns.
* **Reporting**: Automated generation of color-coded Excel dashboards with diagnostic audit logs.

## **Repository Structure**
* **`/src`**: Contains `MainWindow.xaml.cs`, the functional forensic engine logic.
* **`/test_material`**: Includes anonymized sample inputs and the resulting audit dashboard to demonstrate functional success.
* **`packages.config`**: Manifest of NuGet dependencies (OpenXML, SpreadsheetLight).

## **How to Review Results**
Open the `test_material/Audit_Result_Output.xlsx` file to see the engine in action:
* **🔵 Blue**: Standard Primary Duty hours.
* **🔴 Red**: Overlap Conflict (Potential Fraud) or Safety Limit violation.
* **🟣 Purple**: Duplicate primary employment entries (System Error).
