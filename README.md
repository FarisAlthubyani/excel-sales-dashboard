# Excel Sales Dashboard

An interactive sales analysis project built in Microsoft Excel using Power Query, Power Pivot, DAX, Pivot Tables, Pivot Charts, and Slicers.

The project combines sales data from 2023 and 2024, cleans and transforms the source files, builds a relational data model, and presents the results in an interactive dashboard.

## Dashboard Preview

Add a screenshot of the dashboard here after uploading it to the repository:

<img width="1361" height="726" alt="image" src="https://github.com/user-attachments/assets/57569838-9c9a-4469-8cb1-c3ec63911e15" />
<img width="1490" height="750" alt="image" src="https://github.com/user-attachments/assets/93cdce32-1a13-43c3-97d1-751069e008d1" />


## Project Objectives

- Combine sales data from multiple years.
- Clean and prepare raw data using Power Query.
- Build relationships between sales and dimension tables.
- Create DAX measures for the main business metrics.
- Analyze sales by branch, month, category, brand, and payment method.
- Build an interactive Excel dashboard controlled by slicers.

## Tools Used

- Microsoft Excel
- Power Query
- Power Pivot
- DAX
- Pivot Tables
- Pivot Charts
- Slicers

## Data Preparation

Power Query was used to:

- Import data from multiple Excel files.
- Promote the first row to column headers.
- Assign appropriate data types.
- Remove unnecessary columns.
- Remove duplicate records.
- Split customer names and email addresses into separate columns.
- Fill missing country and product-category values.
- Remove rows containing invalid dates or quantities.
- Append the 2023 and 2024 sales datasets into one sales table.

## Data Model

The workbook uses a relational data model containing the following tables:

- Sales
- Calendar
- Customers
- Products
- Brands
- Branches
- Payment Methods

The central sales table is connected to the dimension tables through customer, product, branch, payment-method, and date fields.

## DAX Measures

```DAX
Total Sales =
SUM(Append1[TotalAmount])
```

```DAX
Active Customers =
DISTINCTCOUNT(Append1[CustomerID])
```

## Key Performance Indicators

- **Total Sales:** 974,693.16
- **Active Customers:** 638
- **Top-Selling Brand:** Brand D
- **Top-Selling Category:** Fashion
- **Top-Performing Branch:** Branch 3
- **Most-Used Payment Method by Sales:** PayPal
- **Highest-Sales Month:** January

## Dashboard Analysis

The dashboard includes:

- Total sales and active-customer KPI cards.
- Sales by branch.
- Monthly sales trend.
- Sales by payment method.
- Sales by category and brand.
- Interactive filters for branch, brand, category, and month.

All slicers are connected to the dashboard's Pivot Tables, allowing the user to filter the complete report interactively.

## Key Insights

- Brand D generated the highest sales among the five brands.
- Fashion was the highest-performing product category.
- Branch 3 recorded the highest total sales.
- PayPal generated the highest sales among the available payment methods.
- January was the strongest month, while November recorded the lowest monthly sales.

## How to Use the Workbook

1. Download `sales-dashboard.xlsx`.
2. Open the file using Microsoft Excel.
3. Enable content if Excel displays a security notice.
4. Open the `Dashboard` worksheet.
5. Use the slicers to filter the charts and KPIs.

> **Refresh note:** The dashboard and saved results can be viewed without refreshing. Power Query currently references source files stored on the original computer. To use `Refresh All` on another computer, download the source files and update the file paths in Power Query.

## Repository Structure

```text
excel-sales-dashboard/
├── README.md
├── sales-dashboard.xlsx
├── screenshots/
│   └── dashboard-overview.png
└── data/
    └── source files (optional)
```

## Data Privacy

The repository uses public or non-confidential project data. Review the workbook and source files before publishing future versions to ensure that they do not contain private customer, employee, or company information.

## Future Improvements

- Replace fixed Power Query file paths with a configurable folder parameter.
- Add year-over-year and month-over-month growth measures.
- Add average order value and total quantity KPIs.
- Improve dashboard printing and export settings.
- Add a dedicated data dictionary and model diagram.

## Author

Created by **Fares** as an Excel data analysis and dashboard portfolio project.
