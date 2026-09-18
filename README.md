# DATA-ANALYSIS-ASS--1# 📊 Excel Data Exploration & Analysis – Product Dataset

## 📌 Project Overview

This project is a beginner-level **Data Analysis assignment using Microsoft Excel**. The objective is to perform basic data exploration, summarization, conditional analysis, logical categorization, and text manipulation on a Product Dataset.

This project helped me build foundational skills that are essential for a career in **Data Analytics**, especially working with Excel formulas and structured datasets.

---

## 🎯 Objectives

The main objectives of this project are:

* Explore and summarize product data
* Perform basic numerical calculations
* Analyze minimum and maximum values
* Apply logical conditions using the `IF` function
* Perform conditional calculations using `SUMIF` and `COUNTIF`
* Extract information from Product IDs using `LEFT`, `RIGHT`, and `MID`
* Organize analytical results clearly in Excel

---

## 🗂️ Dataset Description

The dataset contains information about different products.

### Dataset Attributes

| Column       | Description                        |
| ------------ | ---------------------------------- |
| Product ID   | Unique identifier for each product |
| Product Name | Name of the product                |
| Brand Name   | Brand associated with the product  |
| Quantity     | Number of products                 |
| Category     | Product category                   |
| Price        | Price of the product               |

---

## 🔍 Tasks Performed

### 1. Basic Data Exploration

Used Excel functions to calculate:

* **Total Price** using `SUM`
* **Number of Products** using `COUNT`
* **Average Price** using `AVERAGE`

### 2. Minimum & Maximum Analysis

Used:

* `MIN` – to identify the minimum product price
* `MAX` – to identify the maximum product price

### 3. Price Categorization Using IF

Created a new column called **Price Range**.

The `IF` function was used with the following condition:

* Price ≥ $500 → **High Price**
* Price < $500 → **Standard Price**

Example formula:

```excel
=IF(F2>=500,"High Price","Standard Price")
```

---

### 4. Conditional Analysis

Used `SUMIF` and `COUNTIF` to perform conditional calculations.

#### SUMIF

Calculated the total price of products belonging to the **Electronics** category.

Example:

```excel
=SUMIF(E:E,"Electronics",F:F)
```

#### COUNTIF

Calculated the number of products with a price below **$100**.

Example:

```excel
=COUNTIF(F:F,"<100")
```

---

### 5. Text Manipulation

Extracted information from the **Product ID** using Excel text functions.

Created the following columns:

#### Day

Extracted the first 2 characters using `LEFT`.

```excel
=LEFT(A2,2)
```

#### Country Code

Extracted the last 2 characters using `RIGHT`.

```excel
=RIGHT(A2,2)
```

#### Month

Extracted characters 4 to 6 using `MID`.

```excel
=MID(A2,4,3)
```

---

## 🛠️ Tools & Skills Used

**Tool:**

* Microsoft Excel

**Excel Functions:**

* `SUM`
* `COUNT`
* `AVERAGE`
* `MIN`
* `MAX`
* `IF`
* `SUMIF`
* `COUNTIF`
* `LEFT`
* `RIGHT`
* `MID`

**Skills Developed:**

* Data Exploration
* Data Summarization
* Logical Analysis
* Conditional Aggregation
* Text Manipulation
* Spreadsheet Analysis
* Formula Application

---

## 📁 Project Files

```text
Excel-Data-Exploration/
│
├── 📊 Excel Assignment 1 - Data Exploration.xlsx
├── 📄 Assignment Documentation.pdf
└── 📖 README.md
```

### Excel Workbook

The Excel workbook contains:

* Original product dataset
* Calculated results
* Price Range column
* Day column
* Country Code column
* Month column
* Applied Excel formulas

### PDF Documentation

The PDF contains screenshots showing:

* Calculation results
* Newly created columns
* Excel formulas
* Formula bar
* Completed analysis

---

## 📈 Key Learning Outcomes

Through this assignment, I learned how to:

* Summarize numerical data using Excel
* Apply statistical functions to a dataset
* Create categories using logical conditions
* Perform conditional calculations
* Extract useful information from text fields
* Use Excel formulas for basic data analysis
* Structure an Excel workbook for analytical purposes

---

## 🚀 Portfolio Journey

This project is part of my journey toward becoming a **Data Analyst**.

I am currently building my skills in:

* Excel
* Data Analysis
* SQL
* Python
* Power BI
* Data Visualization

I will continue adding more projects to this GitHub repository as I progress through my Data Analytics learning journey.

---

## 👨‍💻 Author

**Afsar**

Aspiring Data Analyst | B.Com Student | Data Analytics Learner

---

⭐ This repository documents my learning journey and practical experience with data analysis tools and techniques.
