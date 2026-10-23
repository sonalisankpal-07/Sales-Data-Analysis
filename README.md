# 📊 Sales Data Analysis Dashboard

## 📌 Project Overview

The **Sales Data Analysis Dashboard** is a data analytics project developed using Python, Pandas, Excel, and Microsoft Power BI.

## 🎯 Project Objectives

* Load and explore sales data using Python and Pandas.
* Convert date columns into the appropriate datetime format.
* Filter sales records based on specified sales-value conditions.
* Separate the dataset into customer, product, and order tables.
* Remove duplicate customer and product records.
* Export the processed tables into an Excel workbook.
* Develop an interactive Power BI dashboard.
* Analyze sales and profit by category, sub-category, month, state, and product.

## 🛠️ Technologies Used

| Technology         | Purpose                                         |
| ------------------ | ----------------------------------------------- |
| Python             | Data preprocessing and transformation           |
| Pandas             | Data manipulation and cleaning                  |
| Jupyter Notebook   | Writing and executing Python code               |
| Excel              | Storing processed tables in separate worksheets |
| Microsoft Power BI | Interactive dashboard development               |
| CSV                | Source dataset format                           |

## 📁 Project Structure

```text
Sales-Data-Analysis/
│
├── README.md
├── SalesProject(2).csv
├── project 2(2).ipynb
└── Sales_Dashboard.png
```

**Files:**

* `SalesProject(2).csv` — Original sales dataset.
* `project 2(2).ipynb` — Python notebook for data preprocessing and transformation.
* `Sales_Dashboard.png` — Dashboard screenshot, if added to the repository.
* `README.md` — Project documentation.

*The dashboard screenshot filename above is a suggested name. Rename your screenshot or update the filename in this structure to match your actual repository.*

## 📂 Dataset Description

The dataset contains **9,994 rows and 21 columns** describing orders, customers, products, sales, discounts, and profits.

### Important Columns

| Column        | Description                              |
| ------------- | ---------------------------------------- |
| Row ID        | Unique row identifier                    |
| Order ID      | Order identification number              |
| Order Date    | Date the order was placed                |
| Ship Date     | Date the order was shipped               |
| Ship Mode     | Shipping method                          |
| Customer ID   | Customer identification number           |
| Customer Name | Customer name                            |
| Segment       | Customer segment                         |
| Country       | Country of the customer or order         |
| City          | City associated with the order           |
| State         | State associated with the order          |
| Region        | Geographical region                      |
| Product ID    | Product identification number            |
| Category      | Product category                         |
| Sub-Category  | Product sub-category                     |
| Product Name  | Name of the product                      |
| Sales         | Sales amount                             |
| Quantity      | Quantity ordered                         |
| Discount      | Discount applied                         |
| Profit        | Profit or loss associated with the order |

## ⚙️ Project Workflow

### Step 1: Import Libraries

Import Pandas to read, inspect, and manipulate the dataset.

```python
import pandas as pd
```

### Step 2: Load the Dataset

Load the CSV file into a Pandas DataFrame.

```python
df = pd.read_csv("SalesProject(2).csv", encoding="latin-1")
```

The `latin-1` encoding is used to read the uploaded dataset successfully.

### Step 3: Explore the Dataset

Inspect the dataset structure, column names, and data types.

```python
df.info()
```

This helps identify the number of records, available columns, and their data types.

### Step 4: Convert Date Columns

Convert the order and shipping dates into datetime format.

```python
df["Order Date"] = pd.to_datetime(df["Order Date"])
df["Ship Date"] = pd.to_datetime(df["Ship Date"])
```

This makes the date columns suitable for date-based analysis.

### Step 5: Filter Sales Records

Apply the sales-value conditions used in the notebook.

```python
df = df[(df["Sales"] >= 25) & (df["Sales"] < 100000)]
```

This retains records where sales are at least 25 and below 100,000.

### Step 6: Create the Customer Table

Select customer-related columns and remove duplicate customer IDs.

```python
customer_df = df[
    [
        "Customer ID",
        "Customer Name",
        "Segment",
        "Country",
        "City",
        "State",
        "Postal Code",
        "Region"
    ]
]

customer_df = customer_df.drop_duplicates(subset="Customer ID")
```

This creates a customer table containing distinct customer IDs.

### Step 7: Create the Product Table

Select product-related columns and remove duplicate product IDs.

```python
product_df = df[
    [
        "Product ID",
        "Category",
        "Sub-Category",
        "Product Name"
    ]
]

product_df = product_df.drop_duplicates(subset="Product ID")
```

This creates a product table with distinct product IDs.

### Step 8: Create the Orders Table

Select order-related columns for sales analysis.

```python
orders_df = df[
    [
        "Row ID",
        "Order ID",
        "Ship Date",
        "Ship Mode",
        "Customer ID",
        "Product ID",
        "Sales",
        "Quantity",
        "Discount",
        "Profit"
    ]
]
```

The orders table connects order information with customer IDs, product IDs, and financial measures.

### Step 9: Export the Processed Data

Export the three DataFrames into separate worksheets in an Excel workbook.

```python
with pd.ExcelWriter(
    "Project1.xlsx",
    engine="xlsxwriter"
) as writer:

    customer_df.to_excel(
        writer,
        sheet_name="customer",
        index=False
    )

    product_df.to_excel(
        writer,
        sheet_name="product",
        index=False
    )

    orders_df.to_excel(
        writer,
        sheet_name="orders",
        index=False
    )

print("File created")
```

The resulting workbook contains three worksheets:

* `customer`
* `product`
* `orders`

Install the required Excel-writing engine if necessary:

```python
pip install xlsxwriter
```

## 📊 Power BI Dashboard

The processed data is used to create an interactive sales dashboard in Microsoft Power BI.

### 1. KPI Cards

The dashboard displays five summary cards:

* Net
* Amount Profit
* Amount Discount
* Sum of Profit
* Gross Net

These cards provide a quick overview of the financial figures displayed in the report.

### 2. Category Slicer

The category slicer allows users to filter the report by:

* Furniture
* Office Supplies
* Technology

The selected category can be used to explore the corresponding sales and profit figures in the dashboard.

### 3. Segment and Region Slicer

This slicer provides filtering based on customer segment and geographical region.

It allows users to explore performance for selected customer groups and regions.

### 4. Net by Sub-Category

A pie chart displays net values across product sub-categories, such as Accessories, Copiers, Machines, and Phones.

It helps users compare the displayed contribution of different sub-categories.

### 5. Amount Profit by State

The map visual displays profit information geographically by state.

It helps users explore the geographical distribution of the profit values shown in the report.

### 6. Net and Amount Profit by Category

A clustered column chart compares net and profit across product categories.

It provides a visual comparison of the two measures for the selected category or categories.

### 7. Net and Amount Profit by Month

A line chart displays monthly net and profit trends.

It helps users examine changes in the displayed financial measures over the months.

### 8. Amount Profit by Product Name

A horizontal bar chart compares profit values across individual products.

It helps users identify products with higher or lower displayed profit values.

## 🚀 How to Run the Project

### Prerequisites

Install or prepare the following:

* Python 3
* Jupyter Notebook
* Pandas
* XlsxWriter
* Microsoft Power BI Desktop

### Execution Steps

1. Clone or download this GitHub repository.
2. Keep the CSV dataset and Jupyter Notebook in the same project folder.
3. Open the notebook using Jupyter Notebook or JupyterLab.
4. Ensure the CSV filename in the notebook matches the actual dataset filename.
5. Run the notebook cells in sequence.
6. Verify that the customer, product, and orders DataFrames are created.
7. Export the processed tables to the Excel workbook.
8. Open Power BI Desktop.
9. Import the processed Excel workbook using **Home → Get Data → Excel workbook**.
10. Load the customer, product, and orders worksheets.
11. Create appropriate relationships between the tables using their IDs.
12. Add the KPI cards, slicers, pie chart, map, column chart, line chart, and horizontal bar chart.
13. Configure the fields, aggregations, formatting, and interactions for each visual.
14. Save the Power BI report.

## 📈 Key Outcomes

* Explored a sales dataset containing 9,994 records and 21 columns.
* Converted order and shipping dates to datetime format.
* Filtered records using specified sales-value conditions.
* Created separate customer, product, and order tables.
* Removed duplicate customer IDs and product IDs from the corresponding tables.
* Exported the processed tables into a structured Excel workbook.
* Presented sales and profit information through an interactive Power BI dashboard.

## 💡 Skills Demonstrated

* Python programming
* Pandas data manipulation
* Data preprocessing
* Date and data-type conversion
* Duplicate removal
* Conditional filtering
* Data organization and table creation
* Excel data export
* Power BI dashboard development
* Data visualization
* Business data analysis

##  Future Improvements

* Add a dedicated date table for more reliable time-based analysis.
* Create DAX measures for clearly defined sales, profit, discount, and order KPIs.
* Add a profit-margin KPI.
* Include year, quarter, and month filters.
* Add a top-performing products analysis.
* Improve the dashboard layout, color consistency, and tooltips.
* Document the data model and relationships.
* Publish the report using an appropriate Power BI sharing option.

Data Analytics Project.
