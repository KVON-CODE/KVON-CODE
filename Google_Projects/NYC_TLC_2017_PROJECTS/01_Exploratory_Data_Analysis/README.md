# NYC TLC: Exploratory Data Analysis & Visualization

[← Back to NYC TLC 2017 Projects](../README.md)

## Executive Summary

### Key Insights
Exploratory analysis identified significant data anomalies where high fares were recorded for trips with zero distance. These discrepancies and skewed ride durations pose a risk of producing biased results if the raw data is used directly for predictive modeling. By applying rigorous cleaning and visualizing these outliers, the dataset was successfully prepared for accurate revenue forecasting and business reporting.

![NYC TLC Scatter Plot](images/Course_3_Scatter_Plot.png)

### Details of the Visualization
* **[Course_3_Scatter_Plot.png](images/Course_3_Scatter_Plot.png):** This visualization identifies a meaningful linear trend between Total Amount and Trip Distance while clearly highlighting the zero-distance outliers that require sanitization.
* **Data Integrity Flags:** By highlighting data points where the total amount remains high despite zero distance, the chart identifies specific records requiring data sanitization before further analysis.

### Keys to Success
* **Data Integrity Verification:** Ensuring the data source is verifiable and accurate is the foundational step in preventing downstream errors in the analytical pipeline.
* **Systematic Outlier Management:** Developing a structured plan to distinguish between genuine errors and meaningful statistical outliers ensures that predictive models remain representative of real-world operations.

### Next Steps
* **Profitability Variable Analysis:** Further analyze secondary variables, such as payment type or vendor, to determine which specific factors most significantly impact overall profit margins.
* **Regression Model Development:** Transition from exploratory visualization to formal regression modeling to establish a quantitative forecasting tool for the NYC TLC management team.

---

## Technologies & Tools
* **Language:** Python
* **Libraries:** Pandas, NumPy, Matplotlib, Seaborn
* **Visualization:** Tableau Desktop
* **Documentation:** Jupyter Notebook

---

## Project Files
* **[Course_3_Project](Course_3_Project.ipynb):** The primary Jupyter Notebook containing data cleaning, structuring, and visualizations.
* **[Course_3_Project_PDF](Course_3_Project_PDF.pdf):** A static, readable export of the full notebook for quick technical review.
* **[Course_3_Executive_Summary](Course_3_Executive_Summary.pdf):** A high-level report designed for stakeholders and management.
* **[Course_3_Scatter_Plot](images/Course_3_Scatter_Plot.png):** High-resolution export of the Tableau visualization showing the relationship between trip distance and fare.

---

[← Back to NYC TLC 2017 Projects](../README.md)

