# Excel_Manufacturing

## Integrated Operations Analysis: A Case Study on ElectroniTech Manufacturing Co.

**Link to project**

https://drive.google.com/drive/folders/1kU7pr1HKnNnrLSXXC7yf-3ZmOJK0zUQ_?usp=sharing

## Overview

The project aims to analyze ElectroniTech Manufacturing Co. 's operational datasets to gain insights into customer behavior, employee efficiency, shipment logistics, and financial transactions. Students are required to perform comprehensive data analysis using Excel to identify trends, inefficiencies, and potential areas for improvement. Key tasks include data cleaning and integration, detailed analysis of customer and employee data, shipment tracking, financial analysis, and the development of a dynamic dashboard to visualize key metrics. This project will aid in streamlining ElectroniTech Manufacturing Co.'s operations, optimizing supply chain management, and enhancing customer relations, ultimately contributing to the company's strategic decision-making and growth.

## Background

ElectroniTech Manufacturing Co. is a leading company in the manufacturing and distribution of electronic components. The company has a diverse clientele ranging from large-scale wholesalers to small retail businesses, and it operates across multiple manufacturing sites and regional offices. As the company expands, it faces challenges in managing its increasing volume of data, related to customers, employees, shipments, and financial transactions. Efficient handling and analysis of this data are crucial for maintaining operational excellence, customer satisfaction, and financial health.

## Excel Data Analysis

1. **Handling Missing and Irrelevant Data:** Identify and handle missing data in both datasets. Address duplicate entries and irrelevant data points, ensuring data quality. 
3. **Name and Email Normalization:** Normalize the customer names in **`Customer.csv`** to have the first letter capitalized and the rest lower. Standardize the email addresses  to follow a consistent format. For example, ensure all email addresses are in lowercase.
4. **Customer-Membership Merge:** Merge **`Customer.csv`** with **`Membership.csv`** using a common key. What key did you use and why?
5. **Shipment-Status Integration:** Create a new dataset by joining **`Shipment_Details.csv`** with **`Status.csv`**. Describe the method used for merging these datasets.
6. **Payment Categorization:** In **`Payment_Details.csv`**, create a new column that categorizes payments into 'High', 'Medium', and 'Low' based on the amount.
7. **Average Shipment Weight:** Calculate the average shipment weight for each type of service in **`Shipment_Details.csv`**.
8. **Membership Duration Calculation:** Use a date function to find out how long each customer has been a member (from **`Membership.csv`**).
9. **Total Customer Payment:** Create a formula to calculate the total amount paid by each customer in **`Payment_Details.csv`**.
10.  **Shipment Domain Comparison:** Analyze the **`Shipment_Details.csv`** dataset to find out which shipment domain (Domestic/International) has more shipments.
11. **Top Employees Characterization:** What are the common characteristics of the top 5 employees managing the highest number of shipments?
12. **Conditional Formatting for Weight:** Use conditional formatting in **`Shipment_Details.csv`** to highlight shipments weighing over 500 units.
13. **Employee Detail Dropdown:** Create a dynamic drop-down list in **`Employee_Details.csv`** to select and display details of an employee.
14. **Payment Summary Pivot Table:** Implement a pivot table to summarize the average payment amount by customer type and payment mode.
15. **Shipment Weight Categorization:** Categorize shipments in **`Shipment_Details.csv`** as 'Heavy' or 'Light' based on weight.
16. **Employee Shipment Count:** In **`employee_manages_shipment.csv`**, use an array formula to count the number of shipments managed by each employee.
17. **Dynamic Payment Analysis:** In **`Payment_Details.csv`**, create a dynamic named range for each payment mode (e.g., CARD PAYMENT, COD). Then use a formula to calculate the average payment amount for each mode.
18. **Average Delivery Time Analysis:** Analyze the combined data of **`Shipment_Details.csv`** and **`Status.csv`** to determine the average delivery time for shipments.(Consider using the **`DATEDIF`** function or similar.) How does the average delivery time vary between domestic and international shipments?
19. **Designation-wise Shipment Charge:** Using a pivot table, analyze the **`employee_manages_shipment.csv`** and **`Shipment_Details.csv`** datasets to find out which employee designation has the highest average shipment charges. Incorporate slicers or timeline filters for dynamic data exploration.
20. **Shipment Efficiency Analysis:** Analyze the **`Shipment_Details.csv`** to calculate the efficiency of shipments. Define efficiency as the ratio of **`SH_WEIGHT`** to **`SH_CHARGES`**, categorized by **`SH_DOMAIN`** (Domestic/International) and **`SER_TYPE`** (Regular/Express). Identify the most and least efficient categories.
21. **Payment Forecast:** Using the **`Payment_Details.csv`**, create a time-series forecast to predict future payment amounts for the next year based on historical data. Break down the forecast by payment mode. (**Hint**: Utilize Excel's forecasting tools, like the Forecast Sheet feature. Ensure the data is in a chronological order and consider using a line chart for visualization. Analyze trends and seasonality in the payment data.)
22. **Customer Segmentation:** Segment customers in **`Customer.csv`** based on their transaction behavior from **`Payment_Details.csv`** and membership duration from **`Membership.csv`**. Create segments such as 'High Value - Long Term', 'Low Value - Short Term', etc., based on total payment amounts and membership lengths.
## Dashboard
Interactive excel dashboard was designed to provide comprehensive insights into various aspects of ecommerce operations. Each dashboard leverages advanced Excel techniques to visualize key metrics, trends, and patterns, offering valuable information for data-driven decision-making.


![image](https://github.com/user-attachments/assets/b7495fa6-e2f3-47e8-a394-5b22a1e36275)


