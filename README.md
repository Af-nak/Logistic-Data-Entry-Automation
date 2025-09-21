# Logistics Form User Interface (Excel Automation)

##  Table of Contents
## Table of Contents  
- [Project Overview](#project-overview)  
- [Key Insights](#key-insights)  
- [Data Sources](#data-sources)  
- [Tools](#tools)  
- [Exploratory Data Analysis](#exploratory-data-analysis)  
- [Data Analysis](#data-analysis)  
- [Results / Findings](#results--findings)  
- [Recommendations](#recommendations)  
- [Limitations](#limitations)  
- [Collaboration and Connection](#collaboration-and-connection)
  
---

## Project Overview  
This project is an **Excel-based logistics form interface** designed to optimize data entry and automate inbuilt calculations. It provides an alternative to traditional manual Excel tables or forms.  

Key features:  
- Data entry is handled through **two user interfaces** for shipment details (when sent) and shipment arrivals (when received).  
- Automatically populates forecast values (shipping fee and arrival date) from a **Forecast Table** using lookup functions and working-day calculations.  
- **Macros** streamline navigation and automate form submissions.  
- A **Dashboard template** provides insights into key logistics performance metrics.  
- **Power Query modeling** connects three related tables for comprehensive analysis.  

---

## Key Insights  
- Data entry is streamlined into **Logistics** and **Shipment** tables.  
- Forecasted values (fees and arrival dates) are automated with lookup and working-day calculations.  
- Macros enable fast navigation and responsive user experience.  
- A dashboard displays **KPIs and visuals** for decision-making.  
- Relationships are modeled in Power Query & Power Pivot for **data integration and reporting**.  

---

## Data Sources  
The system uses three core tables:  

1. **Logistics Table**  
   - Captures product and shipment details when a shipment is **sent**.  
   - Data entered through the **Logistics Entry Form**.  
   - Columns include:  
     `Order ID`, `Product Name`, `Purchase Price`, `Tracking ID`, `Shipping Mode`, `Weight`, `Quantity`, `Date Sent`, `Forecast Shipping Fee`, `Forecast Arrival Date`.  

2. **Shipment Table**  
   - Records actual shipment details when shipment **arrive**.  
   - Data entered through the **Shipment Entry Form**.  
   - Captures **Actual Shipping Fee** and **Actual Arrival Date**.  

3. **Forecast Table**  
   - Contains company-provided reference data from past shipment records.  
   - Used to estimate **Forecast Arrival Date** and **Forecast Shipping Fee**.  

---

## Tools  
- **Microsoft Excel**  
- **Functions:** `XLOOKUP`/ `VLOOKUP`, `NETWORKDAYS`  
- **Charts:** Line, Doughnut, Bar, Clustered Column, Scatter, KPI cards  
- **Power Query & Data Model-Power Pivot**  
- **VBA Macros** for automation  

---

## Exploratory Data Analysis
The following guiding questions were explored to generate insights:  

1. **Top Products by Cost**  
   - What are the top products the company spends most on, in terms of total cost?  

2. **Shipping Mode Usage**  
   - What percentage of each shipping mode is most frequently used to ship products?  

3. **Shipping vs. Product Costs**  
   - Is the company paying more on shipping fees compared to product costs annually?  
   - What is the difference between the two?  

4. **Forecast Accuracy**  
   - Since shipping fees are forecasted ahead, what is the trend in forecast accuracy over time?  

5. **Weight vs. Shipping Cost**  
   - For each shipment in a given month, what is the relationship between shipment weight and shipping cost?  

---

## Data Analysis  
- Designed a **streamlined dashboard** with 5 visuals:  
  1. **Line Chart** – Forecast vs. Actual arrival trends  
  2. **Doughnut Chart** – Distribution of shipping modes  
  3. **Bar Chart** – Total cost by top products  
  4. **Scatter Chart** – Relationship between weight and shipping cost  
  5. **KPI Cards** – On-time forecast %, Total Orders, Total Cost  

- Implemented **Power Query merge queries** to unify the three tables.  
- Built a **date table** to support time-based analysis.  
- Applied **DAX measures** to calculate metrics such as:  
  - Total Shipping Cost  
  - Forecast Accuracy %  
- Recorded **macros** for responsive navigation and form automation.  

👉 *Alternative:* Instead of using **Power Query merge queries**, you can also use **Power Pivot** to directly model relationships between the Logistics, Shipment, and Forecast tables. This approach allows for more advanced and scalable DAX-based calculations.   

---

## Results / Findings  
- Data entry and processing are now streamlined through form-based automation.  
- Forecast shipping fees and arrival dates are more accurate with working-day adjustments.  
- Dashboard visualizations reveal:  
  - High-spend products  
  - Most frequently used shipping modes  
  - Comparison of shipping vs. product costs  
  - Forecast accuracy trends  
  - Relationship between weight and cost  

---

## Recommendations  
- Keep the **Forecast Table** updated regularly for better prediction accuracy.  
- Expand to **Power BI** for more advanced visualization and real-time updates.  
- Implement **error checks** in forms to reduce incorrect data entry.  
- Introduce historical trend analysis to support cost-optimization decisions.  

---

## Limitations  
- The system is limited to **Excel**, which may not scale for large enterprise logistics operations.  
- Accuracy heavily depends on how frequently the **Forecast Table** is updated.  
- VBA macros may not work in all environments (e.g., web-based Excel).  
- Static templates may need adjustments if company processes change.  

---
## Collaboration and Connection  
I am passionate about building data-driven solutions and continuously improving my skills in analytics and automation.  
If you find this project useful or have ideas for improvements, feel free to connect with me or collaborate. 
[Connect with me](https://linktr.ee/_afnak) 
