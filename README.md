# Healthcare-Data-Analysis-Excel
Healthcare data analysis and interactive dashboard using Excel, Power Query, VLOOKUP, PivotTables, charts, and slicers.
# Healthcare Data Analysis and Insights

## Overview

This project focuses on analyzing a healthcare dataset using Microsoft Excel to identify meaningful patterns in patient health profiles, medical history, and healthcare costs.

The project combines data from **Customer Names, Medical Examinations, and Hospitalisation Details** tables using **Customer ID** as the common field. The data was cleaned and transformed before performing analysis using **PivotTables, PivotCharts, and interactive slicers**.

## Project Workflow

The analysis was completed through the following steps:

1. **Data Cleaning**

   * Identified missing values represented by `?`.
   * Replaced missing values in relevant fields using appropriate values or strategies.
   * Checked inconsistencies in fields such as **Smoker** and **Heart Issues**.
   * Converted `NumberOfMajorSurgeries` into numerical data.
   * Handled missing State ID values appropriately.

2. **Data Transformation**

   * Split the customer `Name` field into:

     * Title
     * First Name
     * Last Name
   * Created **Weight Status** based on BMI.
   * Created **Diabetes Status** based on HbA1C.
   * Combined Year, Month and Date into **Date of Birth**.
   * Calculated **Age** based on Date of Birth and the dataset collection date.
   * Formatted healthcare `Charges` as currency.

3. **Data Merging**

   * Combined the required information from the three tables using **Customer ID**.
   * Used **VLOOKUP** to retrieve related fields and create a consolidated healthcare dataset.

4. **Data Analysis**
   PivotTables were created to analyze:

   * Cancer history among smokers and non-smokers.
   * Major surgeries and average HbA1C based on transplant history.
   * Healthcare charges by Weight Status and Diabetes Status.
   * Average healthcare charges for different Hospital Tiers within each State.
   * Relationship between Age and BMI.
   * Relationship between Age and HbA1C.
   * Relationship between Age and healthcare Charges.

5. **Visualization & Dashboard**

   * Created Pie/Donut, Column/Bar, and Line charts to visualize the analysis.
   * Built an interactive healthcare dashboard.
   * Added **Diabetes Status** and **Weight Status** slicers to enable interactive filtering across the dashboard.

## Excel Functions & Formulas Used

### 1. VLOOKUP

Used to combine information from the different healthcare tables using **Customer ID**.

```excel
=VLOOKUP(A2,TableRange,ColumnNumber,FALSE)
```

### 2. IF Function – Weight Status

Used to categorize patients according to their BMI.

```excel
=IF(BMI<18.5,"Underweight",IF(BMI<25,"Normal Weight",IF(BMI<30,"Overweight","Obesity")))
```

Categories used:

* Below 18.5 → Underweight
* 18.5–24.9 → Normal Weight
* 25.0–29.9 → Overweight
* 30.0 and above → Obesity

### 3. IF Function – Diabetes Status

Used to classify patients based on HbA1C.

```excel
=IF(HBA1C<5.7,"Normal",IF(HBA1C<6.5,"Prediabetes","Diabetes"))
```

Categories used:

* Below 5.7 → Normal
* 5.7–6.4 → Prediabetes
* 6.5 and above → Diabetes

### 4. Age Calculation

Age was calculated from the patient's Date of Birth using the dataset collection date of **8 June 2023**.

```excel
=DATEDIF(Date_of_Birth,DATE(2023,6,8),"Y")
```

### 5. Data Cleaning & Transformation

Power Query was used where appropriate for:

* Splitting columns
* Replacing missing values
* Changing data types
* Transforming columns
* Creating calculated/custom columns
* Cleaning and preparing the source data for analysis

## PivotTable Analysis

The cleaned and consolidated dataset was used to create PivotTables for the required business questions.

### Key analyses included:

* **Cancer History Distribution by Smoking Status**
* **Major Surgeries & Average HbA1C by Transplant History**
* **Average Healthcare Charges by Weight Status and Diabetes Status**
* **Average Charges by Hospital Tier within State**
* **Age vs BMI**
* **Age vs HbA1C**
* **Age vs Healthcare Charges**

These PivotTables were then connected to suitable charts for easier interpretation.

## Dashboard

An interactive dashboard was created to consolidate the major findings from the analysis.

The dashboard includes:

* Healthcare charge analysis
* Patient health-related comparisons
* Medical history analysis
* Age-based analysis
* Interactive charts
* **Diabetes Status slicer**
* **Weight Status slicer**

The slicers allow the user to dynamically filter the dashboard and compare healthcare outcomes and charges across different health categories.

## Tools Used

* Microsoft Excel
* Power Query
* VLOOKUP
* IF
* DATEDIF
* PivotTables
* PivotCharts
* Slicers
* Data Cleaning & Transformation

## Project Outcome

The completed workbook transforms raw healthcare data into a structured analytical dataset and interactive dashboard. The analysis helps identify relationships between patient characteristics, health conditions, medical history, and healthcare charges through Excel-based data analysis and visualization.
