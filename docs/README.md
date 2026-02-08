# 📉 Case Study: Expense Report Cycle Time Reduction
**Kennesaw State University | Six Sigma Green Belt Capstone**

This repository documents a simulated process improvement project focused on optimizing corporate reimbursement workflows. [cite_start]As the **Project Team Lead**, I applied the DMAIC framework to identify root causes of processing delays and modeled data-driven solutions[cite: 7, 15, 22].

---

### 📋 Project Overview
* [cite_start]**Problem Statement:** The average cycle time for expense reports was 12 days, exceeding the 5-day SLA. [cite_start]This caused a 25% employee dissatisfaction rate and $2,500 in annual late fees.
* [cite_start]**Goal Statement:** Reduce the average cycle time from 12 days to 5 days by March 13, 2026.
* [cite_start]**Scope:** Domestic travel reports submitted via the portal, from submission to funds deposited[cite: 7].

### 🔬 Statistical Analysis (The "Analyze" Phase)
[cite_start]I utilized **One-Way ANOVA** and **Pareto Analysis** to move beyond intuition and validate root causes[cite: 105, 107, 183]:

* [cite_start]**ANOVA Results:** Comparing correct reports (4.0 days) vs. defect reports (17.3 days) yielded a **P-Value < 0.001**[cite: 122, 123, 124]. [cite_start]This statistically proved the reprocessing cycle was the primary driver of delays[cite: 126].
* [cite_start]**Pareto Findings:** **67% of defects** were caused by "Missing Receipts"[cite: 130].
* [cite_start]**Root Cause (5 Whys):** Identified a lack of system error-proofing (Poka-Yoke); the "Submit" button remained active even if no files were attached[cite: 137, 138, 139].

### 🛠️ Improvements & Controls
* [cite_start]**Solution Selection:** Due to budget and time constraints, a policy update and standardization of a "Submission Checklist" were implemented[cite: 153, 155, 156].
* [cite_start]**Standardization:** Updated the SOP to administratively reject reports without receipts within 24 hours[cite: 157].
* [cite_start]**Sustainability:** Established a Control Plan to monitor reprocessing rates weekly and cycle times monthly[cite: 176].

### 📈 Simulated Results
| Metric | Baseline | Target | Post-Improvement |
| :--- | :--- | :--- | :--- |
| **Avg Cycle Time** | [cite_start]12.0 Days [cite: 51] | [cite_start]5.0 Days  | [cite_start]**3.4 Days**  |
| **Reprocessing Rate** | [cite_start]60% [cite: 52] | [cite_start]<10%  | [cite_start]**10%** [cite: 169] |

---

### 📂 Supporting Documents
* [Full Project Charter (PDF)](./Six-Sigma-Green-Belt-Project-Charter-Capstone-Project.pdf)
* [Process Map & SIPOC](./process_map.png)

> *"This experience demonstrated how the DMAIC framework can achieve a low-cost solution through data-driven decision-making."* 
---
[Return to Main Repository](../README.md)
