Excel-project

E-Commerce Sales Analysis Using Microsoft Excel

📌 Project Overview

This project focuses on creating, cleaning, analyzing, and visualizing an E-commerce dataset using Microsoft Excel.

The objective is to demonstrate practical data analysis skills by transforming raw data into organized information, performing calculations, summarizing data, and building an interactive dashboard.

This project is part of my learning journey as an **Aspiring Data Analyst**, where I am developing hands-on skills in data cleaning, Excel formulas, data analysis, and visualization.

🎯 Project Objectives

- Import and organize E-commerce data using Excel tables.
- Identify and remove duplicate records.
- Handle missing values and inconsistent data.
- Correct formatting errors and validate data quality.
- Remove irrelevant columns when necessary.
- Apply Excel formulas for calculations, summaries, and logical and text operations.
- Combine related tables using `XLOOKUP` and `VLOOKUP`.
- Summarize data using PivotTables.
- Calculate descriptive statistics.
- Create at least three types of charts.
- Build an interactive dashboard using PivotCharts and filters.
📂 Dataset Description
**Dataset:** [E-commerce Store Dataset](PASTE_DATASET_LINK_HERE)

The dataset is used to practice data preparation and analysis in Excel. The exact columns and record counts should be documented after checking the final dataset.

🛠️ Tools and Technologies

- **Microsoft Excel**
- Excel Tables
- Excel formulas and functions
- XLOOKUP and VLOOKUP
- PivotTables and PivotCharts
- Data cleaning and validation
- Data visualization
- Analysis ToolPak
- GitHub

🔄 Project Workflow

1. Data Import and Organization

- Imported the E-commerce dataset into Microsoft Excel.
- Organized data into tables for easier filtering and analysis.
- Reviewed column names and data structure.

2. Data Cleaning and Validation

- Checked for duplicate records and removed confirmed duplicates.
- Identified missing values and reviewed appropriate ways to handle them.
- Corrected inconsistent text and formatting.
- Used the `TRIM` function to remove unnecessary spaces.
- Applied appropriate data types and number formats.
- Validated the data to identify potential errors.
- Reviewed columns and removed irrelevant fields when justified by the project requirements.

**Note:** Missing values should not automatically be replaced with zero. The correct approach depends on the meaning of the column and the available information.

3. Excel Formulas and Functions

Applied Excel formulas to perform calculations, summarize information, evaluate conditions, and manipulate text.

| Formula | Purpose |
| `SUM()` | Calculates totals |
| `AVERAGE()` | Calculates the average |
| `MIN()` | Finds the minimum value |
| `MAX()` | Finds the maximum value |
| `COUNT()` | Counts numeric cells |
| `COUNTA()` | Counts non-empty cells |
| `COUNTIF()` | Counts values matching a condition |
| `SUMIF()` | Sums values matching a condition |
| `SUMIFS()` | Sums values matching multiple conditions |
| `AVERAGEIF()` | Calculates an average based on a condition |
| `IF()` | Returns a result based on a condition |
| `AND()` | Checks whether all conditions are true |
| `OR()` | Checks whether at least one condition is true |
| `IFERROR()` | Handles formula errors |
| `LEFT()` | Extracts characters from the left |
| `RIGHT()` | Extracts characters from the right |
| `LEN()` | Counts characters |
| `TRIM()` | Removes extra spaces |
| `UPPER()` / `LOWER()` | Changes text case |
| `CONCAT()` | Combines text values |

4. Combining Tables Using Lookup Functions

Used `XLOOKUP` or `VLOOKUP` to retrieve matching information from related tables using common identifiers.

These functions help combine relevant details from multiple tables for further analysis.

Example:

```excel
=XLOOKUP(A2,Product_Dim!A:A,Product_Dim!B:B,"Not Found")
```

This example looks for the ID in cell `A2` in column A of `Product_Dim` and returns the corresponding value from column B. Adjust the sheet, columns, and cell references to match your workbook.

5. PivotTables for Data Summarization

PivotTables are used to organize and summarize data based on relevant fields.

Possible analysis includes:

- Summarizing sales by category.
- Comparing sales across regions.
- Reviewing monthly sales trends.
- Grouping records by relevant product or customer fields.

The available analyses depend on the columns in the dataset.

6. Descriptive Statistics

Descriptive statistics help understand the distribution and variation of numerical data.

| Statistic | Excel Formula | Purpose |
|---|---|---|
| Mean | `=AVERAGE(range)` | Calculates the average |
| Median | `=MEDIAN(range)` | Finds the middle value |
| Mode | `=MODE.SNGL(range)` | Finds the most frequent value |
| Standard Deviation | `=STDEV.S(range)` | Measures data variability |
| Variance | `=VAR.S(range)` | Measures variation |
| Minimum | `=MIN(range)` | Finds the lowest value |
| Maximum | `=MAX(range)` | Finds the highest value |
| Count | `=COUNT(range)` | Counts numeric values |
| Sum | `=SUM(range)` | Calculates the total |
| Range | `=MAX(range)-MIN(range)` | Calculates the difference between maximum and minimum |

The formulas should be applied to the appropriate numerical column in the dataset.

Excel's **Analysis ToolPak** can also be used to generate descriptive statistics.

7. Data Visualization

Created or planned charts to present information in a clear and understandable format.

The project requires at least three chart types:

1. **Bar or Column Chart** — compares values across categories.
2. **Line Chart** — displays changes over time when date information is available.
3. **Pie or Doughnut Chart** — illustrates proportions across categories when appropriate.

The charts should use actual dataset values and meaningful titles, labels, and legends.

8. Interactive Dashboard and PivotCharts

Developed an Excel dashboard to bring key summaries and visualizations together in one place.

Dashboard features may include:

- Sales summary charts.
- Category-wise comparisons.
- Monthly trends, where date data is available.
- PivotCharts connected to PivotTables.
- Interactive filters or slicers, where supported.

The dashboard allows users to explore the available data and understand patterns more easily.

📊 Key Skills Demonstrated

- Data entry and organization
- Data cleaning and validation
- Excel formulas and functions
- Lookup functions and table combination
- PivotTables and PivotCharts
- Descriptive statistics
- Data visualization
- Interactive dashboard creation
- Analytical thinking and problem-solving

🚀 How to Use This Project

1. Clone or download this GitHub repository.
2. Open the Excel workbook in Microsoft Excel.
3. Review the source dataset and cleaned data.
4. Explore the formulas and lookup functions.
5. Review the descriptive statistics and PivotTables.
6. Open the dashboard to explore the charts and available filters.
7. Refresh PivotTables and PivotCharts if the source data changes.

✅ Project Outcome

This project demonstrates how Microsoft Excel can be used to clean, organize, analyze, and visualize E-commerce data.

It provides practical experience in spreadsheet-based data analysis, from data preparation and formula application to descriptive statistics and interactive dashboard development.

Specific business findings and numerical results should be added after verifying the completed analysis.



---

*Thank you for visiting my project!*
