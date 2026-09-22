# 🍽️ Restaurant Sales Analysis Dashboard — Excel

![Restaurant Dashboard](dashboard.png)

## 📊 Project Overview

This project is an **Interactive Restaurant Sales Analysis Dashboard developed in Microsoft Excel** using advanced Excel features such as **VLOOKUP, Pivot Tables, Pivot Charts, GETPIVOTDATA, Slicers, Filters, and Excel Macros/VBA**.

The objective of this project is to transform raw restaurant transaction data into an interactive business dashboard that provides meaningful insights into **sales performance, profitability, customers, food categories, dining preferences, payment methods, and operational performance**.

The dashboard allows users to interact with the report using **slicers and filters**, making it possible to dynamically analyze restaurant performance based on different business dimensions.

---

## 👨‍💻 Author

**Saquib Alam**

**Project:** Interactive Restaurant Sales Analysis Dashboard
**Tool:** Microsoft Excel
**File Format:** `.xlsm`
**Domain:** Data Analytics / Business Intelligence
**Techniques:** Excel Analytics, Pivot Tables, Pivot Charts, VLOOKUP, GETPIVOTDATA, Slicers & VBA

---

## 🎯 Project Objectives

The main objectives of this project are:

* Analyze restaurant sales performance.
* Understand customer purchasing behavior.
* Analyze sales across different food categories.
* Evaluate restaurant profitability.
* Identify customer trends.
* Analyze different dining types.
* Analyze payment-method preferences.
* Compare performance across cities.
* Analyze waiter-level performance.
* Analyze sales by month, quarter, and day.
* Create an interactive dashboard for business reporting.
* Convert raw transactional data into meaningful visual insights.

---

## 🗂️ Dataset Information

The main dataset contains **489 restaurant transaction records**.

The dataset includes information related to:

| Category              | Fields                                        |
| --------------------- | --------------------------------------------- |
| 🧾 Orders             | Order ID, Order Date                          |
| 📅 Time               | Month, Quarter, Order Time, Day               |
| 👤 Customers          | Customer ID, Customer Type, Gender, Age Group |
| 👨‍🍳 Staff           | Waiter                                        |
| 🌍 Location           | City                                          |
| 🪑 Restaurant         | Table No, Dining Type                         |
| 💳 Payments           | Payment Method                                |
| 🍔 Food               | Food Category, Item Name                      |
| 📦 Quantity           | Quantity                                      |
| 💰 Pricing            | Unit Price                                    |
| 🏷️ Discounts         | Discount %                                    |
| 🧾 Tax                | Tax %                                         |
| ⭐ Customer Experience | Rating                                        |
| 💵 Revenue            | Total Sales                                   |
| 💸 Cost               | Cost                                          |
| 📈 Profit             | Profit                                        |
| 👥 Customers          | Unique Customers                              |

---

# 🛠️ Tools & Technologies

### Microsoft Excel

The complete project was developed using Microsoft Excel.

### Excel Features Used

* VLOOKUP
* Pivot Tables
* Pivot Charts
* GETPIVOTDATA
* Slicers
* Filters
* Excel formulas
* Data cleaning
* Data aggregation
* KPI calculations
* Dashboard design
* Conditional formatting
* VBA / Macros

---

# 🔍 Excel Techniques Used

## 1. VLOOKUP

**VLOOKUP** was used to retrieve and match information between datasets based on a common identifier.

Example use cases include:

* Looking up customer-related information
* Matching order information
* Retrieving related attributes
* Supporting data preparation before analysis

Conceptually:

```excel
=VLOOKUP(lookup_value, table_array, column_index, FALSE)
```

VLOOKUP helped make the underlying data more organized and analysis-ready.

---

# 2. Pivot Tables

Pivot Tables were used to summarize the restaurant transaction data.

They allow the dataset to be analyzed by dimensions such as:

* Month
* Quarter
* City
* Food Category
* Item Name
* Customer Type
* Gender
* Age Group
* Dining Type
* Payment Method
* Waiter

Metrics such as the following were analyzed:

* Total Sales
* Total Profit
* Quantity
* Number of Orders
* Customers
* Ratings

---

# 3. Pivot Charts

Pivot Charts were used to convert summarized Pivot Table information into visual reports.

The dashboard can therefore provide a visual representation of restaurant performance instead of relying only on raw tables.

Examples of analysis include:

* Sales by Food Category
* Sales by City
* Sales by Customer Type
* Sales by Dining Type
* Sales by Payment Method
* Profit Analysis
* Monthly Sales
* Order Analysis

---

# 4. GETPIVOTDATA

The **GETPIVOTDATA** function was used to dynamically retrieve values from Pivot Tables.

This allows dashboard KPI values and summary information to remain connected to the underlying Pivot Tables.

Example:

```excel
=GETPIVOTDATA("Total sales",$A$3)
```

This makes the dashboard more dynamic because the displayed values can respond to changes in the Pivot Table and selected filters.

---

# 5. Slicers

Interactive **Slicers** were incorporated into the dashboard to allow users to filter the report visually.

Possible filtering dimensions include:

* Month
* Quarter
* City
* Customer Type
* Gender
* Age Group
* Dining Type
* Payment Method
* Food Category
* Waiter

Instead of manually modifying formulas, users can select slicer options and immediately analyze the corresponding results.

---

# 6. Filters

Excel filters were used to provide additional control over the dataset and dashboard analysis.

Filters can help isolate specific:

* Cities
* Months
* Customers
* Food categories
* Dining types
* Payment methods
* Waiters
* Customer segments

---

# 📊 Dashboard Analysis

The dashboard is designed to provide an overall view of restaurant performance.

## 💰 Sales Analysis

The dashboard analyzes:

* Total Sales
* Sales by Food Category
* Sales by Item
* Sales by City
* Monthly Sales
* Quarterly Sales
* Sales by Dining Type

This helps understand where restaurant revenue is being generated.

---

## 📈 Profit Analysis

Profit is analyzed using the relationship between:

```text
Sales - Cost = Profit
```

The dashboard can be used to examine:

* Overall Profit
* Profit by Food Category
* Profit by City
* Profit by Item
* Profit trends
* Profitability across different customer segments

---

## 👥 Customer Analysis

Customer information is analyzed using:

* Customer Type
* Gender
* Age Group
* Unique Customers
* Dining Type

This provides a better understanding of the restaurant's customer base.

---

## 🍔 Food Category Analysis

Restaurant products are categorized to analyze performance across different food groups.

Examples include:

* Food Category
* Item Name
* Quantity Sold
* Sales
* Profit
* Customer Ratings

This helps identify which types of food contribute to restaurant performance.

---

## 🪑 Dining Type Analysis

The dashboard analyzes different dining methods such as:

* Dine-In
* Takeaway
* Delivery

This provides an overview of how customers interact with the restaurant.

---

## 💳 Payment Method Analysis

Payment methods can be compared to understand customer payment preferences.

Examples include:

* Card
* Cash
* Wallet
* Other available payment methods in the dataset

---

## 🌍 City Analysis

The dataset contains restaurant transaction information across multiple cities.

The dashboard can be filtered by city to analyze:

* Sales
* Profit
* Orders
* Customers
* Food categories
* Dining types

This provides a geographic perspective of restaurant performance.

---

## 👨‍🍳 Waiter Performance

Waiter-level analysis can be used to compare operational performance.

The report can analyze:

* Orders handled
* Sales generated
* Customer ratings
* Customer interactions

This provides an additional operational perspective.

---

# 🎛️ Interactive Dashboard

One of the main features of this project is its interactive design.

Users can select different values from the slicers and filters to dynamically change the dashboard.

### Example workflow

```text
Raw Restaurant Data
        ↓
Data Preparation
        ↓
VLOOKUP / Excel Formulas
        ↓
Pivot Tables
        ↓
Pivot Charts
        ↓
GETPIVOTDATA
        ↓
Slicers & Filters
        ↓
Interactive Dashboard
```

---

# 📌 Key Dashboard Components

The dashboard focuses on the following analytical areas:

### KPI / Summary Section

* Total Sales
* Total Profit
* Total Orders
* Total Customers
* Average Rating
* Total Quantity

### Sales Analysis

* Sales by Category
* Sales by Item
* Sales by City
* Monthly Sales
* Quarterly Sales

### Customer Analysis

* Customer Type
* Gender
* Age Group
* Unique Customers

### Operational Analysis

* Waiter Performance
* Dining Type
* Payment Method

### Interactive Controls

* Slicers
* Filters
* Pivot Table controls

---

# 🧠 Business Insights Supported by the Dashboard

This dashboard can help answer questions such as:

1. What is the overall restaurant sales performance?
2. How much profit is being generated?
3. Which food categories generate the most sales?
4. Which items have higher demand?
5. Which cities generate more sales?
6. Which customer types contribute to sales?
7. What age groups are represented among customers?
8. Which dining type is used most frequently?
9. Which payment methods are commonly used?
10. How does sales performance change over time?
11. Which waiters handle more orders?
12. How do customer ratings vary?
13. Which food categories contribute more to profitability?
14. How do discounts affect sales and profitability?
15. How can restaurant performance be analyzed using different filters?

---

# 📷 Dashboard Preview

The main dashboard screenshot is stored as:

```text
dashboard.png
```

Place the image in the **same folder as README.md**.

Recommended GitHub project structure:

```text
Restaurant-Sales-Analysis/
│
├── README.md
│
├── dashboard.png
│
└── Restaurant_Data.xlsm
```

If you use a different image filename, update the following line in `README.md`:

```markdown
![Restaurant Dashboard](dashboard.png)
```

For example:

```markdown
![Restaurant Dashboard](restaurant_dashboard.png)
```

---

# 📁 Project Files

| File                   | Description                            |
| ---------------------- | -------------------------------------- |
| `README.md`            | Project documentation                  |
| `Restaurant_Data.xlsm` | Excel workbook containing the analysis |
| `dashboard.png`        | Dashboard screenshot                   |

---

# 🔄 Analytical Workflow

The project follows a complete data-analysis workflow:

### Step 1 — Data Collection

Restaurant transaction data was organized in Excel.

### Step 2 — Data Preparation

The dataset was reviewed and prepared for analysis.

### Step 3 — Data Lookup

VLOOKUP was used where required to retrieve and match information.

### Step 4 — Data Summarization

Pivot Tables were created to aggregate important business metrics.

### Step 5 — Visualization

Pivot Charts were created to visually represent the summarized data.

### Step 6 — Dynamic Dashboard Values

GETPIVOTDATA was used to retrieve dynamic values from Pivot Tables.

### Step 7 — Interactivity

Slicers and filters were added to allow users to interact with the report.

### Step 8 — Dashboard Design

All important KPIs and charts were organized into a single interactive dashboard.

---

# 💼 Skills Demonstrated

This project demonstrates practical skills in:

* 📊 Microsoft Excel
* 📈 Data Analysis
* 📉 Data Visualization
* 🔎 VLOOKUP
* 🔄 Pivot Tables
* 📊 Pivot Charts
* 🧮 GETPIVOTDATA
* 🎛️ Slicers
* 🔍 Data Filtering
* 🧹 Data Preparation
* 📋 Business Reporting
* 📌 KPI Development
* 📊 Dashboard Design
* ⚙️ Excel VBA / Macros
* 💡 Business Intelligence

---

# 🚀 How to Use the Dashboard

1. Download the `.xlsm` Excel file.
2. Open the workbook using Microsoft Excel.
3. Enable macros if Excel asks for permission.
4. Navigate to the **Dashboard** sheet.
5. Use the available slicers and filters.
6. Select different categories, cities, months, customer types, or other dimensions.
7. Observe how the dashboard updates.
8. Use the Pivot Tables and Pivot Charts for detailed analysis.

> **Note:** Because the workbook uses Excel features such as Pivot Tables, Slicers, GETPIVOTDATA, and VBA/Macros, Microsoft Excel is recommended for the full interactive experience.

---

# 📈 Project Value

This project demonstrates how raw restaurant transaction data can be transformed into an interactive business intelligence report using Excel.

Instead of simply presenting raw data, the dashboard provides an analytical layer that allows users to explore restaurant performance from multiple perspectives.

The combination of **Pivot Tables, Pivot Charts, GETPIVOTDATA, VLOOKUP, Slicers, Filters, and VBA/Macros** demonstrates practical Excel-based data analytics and dashboard development skills.

---

# 🎓 Portfolio / Academic Project

This project can be presented as an **Excel Data Analytics / Business Intelligence portfolio project** demonstrating practical experience with data preparation, analysis, visualization, dashboard development, and interactive reporting.

---

# 🔮 Future Improvements

Potential future improvements include:

* Connecting the dashboard to Power BI.
* Automating data refresh.
* Adding advanced Excel VBA automation.
* Adding monthly and yearly trend analysis.
* Adding more advanced KPI calculations.
* Adding customer retention analysis.
* Adding predictive sales analysis.
* Adding automated report generation.
* Connecting the dashboard to a live database.
* Creating an automated management reporting system.

---

# 🏷️ Project Tags

```text
Excel
Microsoft Excel
Data Analytics
Restaurant Analytics
Restaurant Sales
Business Intelligence
Dashboard
Excel Dashboard
Pivot Table
Pivot Chart
VLOOKUP
GETPIVOTDATA
Slicer
Excel VBA
Macros
Data Visualization
Business Analytics
KPI Dashboard
```

---

## 👨‍💻 Author

**Saquib Alam**

*Aspiring Data Analyst | Excel | Power BI | Data Visualization | Business Intelligence*

---

⭐ If you find this project useful, feel free to explore the workbook and dashboard.
