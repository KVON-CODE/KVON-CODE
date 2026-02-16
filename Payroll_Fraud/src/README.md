# 🛠️ C# Forensic Engine: MainWindow.xaml.cs

This file contains the core logic for the Payroll Fraud Detection Tool. It is designed to ingest legacy Excel data, sanitize inputs, and perform time-series analysis to identify employment conflicts.

### 🔬 Technical Highlights

#### 1. Regex Data Sanitization
The engine utilizes Regular Expressions and temporal logic to ensure "Zero-Error" data integrity when auditing legacy payroll exports.

* **Example 1:** `Regex.Replace(sl.GetCellValueAsString("B" + i), "[^0-9:AP]", "");`
* **Purpose:** This strips non-essential characters from source cells (like accidental spaces or symbols) to prevent parsing errors during the audit.

* **Example 2:** `Regex.Matches(fDay1, "[0-1]?[0-9]/[0-3]?[0-9]").Count > 0`
* **Purpose:** Identifies rows containing a date format (M/D or MM/DD) to act as a structural filter, skipping headers and metadata so the audit only processes valid entries.

* **Example 3:** `Regex.Matches(ft1, "[0-1]?[0-9]:[0-5][0-9][AP]")`
* **Purpose:** Extracts specific time patterns (e.g., "10:30A") from "dirty" cells so they can be normalized into standard DateTime objects for conflict analysis.

#### 2. Midnight Crossover Logic
A critical feature of the tool is its ability to handle shifts that span across two different calendar days.
* **The Logic:** The system detects if the end time is chronologically earlier than the start time or if the shift ends on a different `DayOfYear`.
* **The Fix:** It automatically migrates the data attribution to the next day's object, maintaining a continuous and accurate audit trail.
* **Example 4:** `if (FT1.Hour >= 12 && FT2.Hour < 12) { nextday++; FT2 = FT2.AddDays(nextday); }`
* **Purpose:** Implements "Midnight Crossover" logic by detecting when a shift ends in the early morning hours of the following day, ensuring the audit trail remains chronologically accurate.

#### 3. Forensic Color-Coding
The tool doesn't just process data; it visualizes risks using conditional formatting in the Excel output.
* **Fraud Detection:** It flags hours where an employee has both "DUTY" (Primary Duty) and "Detail" (Secondary Duty) entries in Red.
* **Safety Compliance:** It automatically calculates total daily duration and highlights any shift exceeding 16.5 hours in Red to signal a compliance violation.

---
[⬅️ Return to Main Folder](../README.md)
