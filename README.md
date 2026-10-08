# Glam Studio - Salon Management System

BCA Major Project | MCM DAV College for Women, Chandigarh (Panjab University)  
Technologies Used: VB.NET, MS Access Database (.accdb), OleDb Engine, ADO.NET Architecture

# Project Overview
Glam Studio is a Windows desktop application built to automate day-to-day operations in salon management. It manages client records, service billing, employee details, shop inventory, and monthly profit/loss reports.
Note: This GitHub repository serves as a technical documentation archive for the BCA Major Project. It includes database query structures, data flow diagrams (DFD), UI design screens, and extracted VB.NET source code modules directly from the verified project report.

# Key Modules and Features

1. Dashboard and Login (Form1.vb and landpage.vb)
   - Multi-user desktop login interface.
   - Quick navigation across all core modules.

2. Client Management (clientpage.vb)
   - Insert, Update, Delete, and Search client records using SQL queries.
   - Track client pre-service history and total billing via Client ID lookup.

3. Dynamic Service Billing (homepage.vb and services.vb)
   - Dynamic billing calculation for salon services.
   - Update service catalogs and pricing lists.

4. Employee Management (employeepage.vb)
   - Manage staff profiles, designated departments, and monthly salary tiers.

5. Retail Shop and Inventory (shop.vb and items.vb)
   - Point-of-sale module to sell retail beauty products (Serums, Shampoos, Toners, etc.).
   - Real-time product billing linked directly to income records.

6. Financial Accounting (records.vb and expenses.vb)
   - Automated monthly profit calculation formula: Profit = Total Income - Total Salaries - Total Expenses.
   - Automated profit or loss alert popups based on calculated monthly totals.


# Database Architecture

The application connects to an MS Access (.accdb) database using Microsoft.ACE.OLEDB.12.0 connection drivers.

# Table Schema Specifications:

| Table Name | Fields and Data Types | Description |
| client | ClientID (Number), First_Name (Text), Last_Name (Text), Email (Text), Phone_Number (Number), Address (Text), Pre_Services (Text), Total_Price (Number) | Stores client details and service history |
| employee | EmpID (Number), First_Name (Text), Last_Name (Text), Department (Text), Salary (Number), Email (Text), Phone_Number (Number), Address (Text) | Stores staff details and salary tiers |
| services | ServiceID (Number), Service_Name (Text), Price (Number) | Catalog of available salon services |
| shop | ItemID (Number), Item_Name (Text), Item_Price (Number) | Catalog of retail beauty products |
| income | Income_ID (Number), Income (Number), Type (Text), Date (Date/Time), Month (Number), Year (Number) | Tracks service and shop transaction incomes |
| expense | ID (Number), Expense (Text), Price (Number), Quantity (Number), Total_Amount (Number), Date (Date/Time), Month (Number), Year (Number) | Tracks salon operational and inventory expenses |

# Complete Documentation
The official project report submitted to Panjab University is archived in the repository root directory:  
Glam_Studio_Project_Report.pdf
