-- ================================================================
-- VANGUARD: Retail Intelligence SQL Query Library
-- ================================================================
-- Description: Advanced analytical queries for retail data analysis
-- Database: Microsoft Access / SQL Server compatible
-- Author: Data Analytics Portfolio Project
-- Last Updated: February 2026
-- ================================================================

-- ================================================================
-- SECTION 1: DATABASE SCHEMA CREATION
-- ================================================================

-- Create Customers Table
CREATE TABLE Customers (
    Customer_ID VARCHAR(50) PRIMARY KEY,
    Customer_Name VARCHAR(100) NOT NULL,
    Email VARCHAR(100),
    Phone VARCHAR(20),
    Registration_Date DATE,
    Customer_Segment VARCHAR(20),
    Lifetime_Value DECIMAL(10,2),
    Total_Transactions INT,
    Avg_Transaction_Value DECIMAL(10,2),
    Region VARCHAR(20),
    Age_Group VARCHAR(20),
    Preferred_Category VARCHAR(50)
);

-- Create Products Table
CREATE TABLE Products (
    Product_SKU VARCHAR(50) PRIMARY KEY,
    Product_Name VARCHAR(100) NOT NULL,
    Category VARCHAR(50),
    Unit_Price DECIMAL(10,2),
    Cost_of_Goods DECIMAL(10,2),
    Profit_Margin DECIMAL(5,2),
    Stock_Level INT,
    Reorder_Point INT,
    Supplier VARCHAR(100),
    Launch_Date DATE,
    Status VARCHAR(20)
);

-- Create Transactions Table
CREATE TABLE Transactions (
    Transaction_ID VARCHAR(50) PRIMARY KEY,
    Customer_ID VARCHAR(50),
    Product_SKU VARCHAR(50),
    Product_Category VARCHAR(50),
    Transaction_Date DATE,
    Unit_Price DECIMAL(10,2),
    Quantity INT,
    Revenue DECIMAL(10,2),
    Variable_Cost DECIMAL(10,2),
    COGS DECIMAL(10,2),
    Marketing_Spend DECIMAL(10,2),
    Region VARCHAR(20),
    Payment_Method VARCHAR(50),
    FOREIGN KEY (Customer_ID) REFERENCES Customers(Customer_ID),
    FOREIGN KEY (Product_SKU) REFERENCES Products(Product_SKU)
);

-- ================================================================
-- SECTION 2: KEY PERFORMANCE INDICATOR (KPI) QUERIES
-- ================================================================

-- Query 1: Customer Lifetime Value Calculation
-- Purpose: Calculate actual CLV based on historical transactions
SELECT 
    c.Customer_ID,
    c.Customer_Name,
    c.Customer_Segment,
    COUNT(t.Transaction_ID) AS Total_Transactions,
    SUM(t.Revenue) AS Lifetime_Revenue,
    AVG(t.Revenue) AS Avg_Transaction_Value,
    SUM(t.Revenue) / COUNT(t.Transaction_ID) AS Revenue_Per_Transaction,
    DATEDIFF(day, c.Registration_Date, GETDATE()) AS Days_As_Customer,
    SUM(t.Revenue) / NULLIF(DATEDIFF(day, c.Registration_Date, GETDATE()), 0) AS Daily_Revenue_Rate
FROM Customers c
LEFT JOIN Transactions t ON c.Customer_ID = t.Customer_ID
GROUP BY c.Customer_ID, c.Customer_Name, c.Customer_Segment, c.Registration_Date
ORDER BY Lifetime_Revenue DESC;

-- Query 2: Revenue Velocity by Month
-- Purpose: Calculate month-over-month revenue growth rate
WITH MonthlyRevenue AS (
    SELECT 
        FORMAT(Transaction_Date, 'yyyy-MM') AS Month,
        SUM(Revenue) AS Monthly_Revenue
    FROM Transactions
    GROUP BY FORMAT(Transaction_Date, 'yyyy-MM')
),
RevenueWithLag AS (
    SELECT 
        Month,
        Monthly_Revenue,
        LAG(Monthly_Revenue, 1) OVER (ORDER BY Month) AS Previous_Month_Revenue
    FROM MonthlyRevenue
)
SELECT 
    Month,
    Monthly_Revenue,
    Previous_Month_Revenue,
    CASE 
        WHEN Previous_Month_Revenue IS NULL THEN NULL
        ELSE ((Monthly_Revenue - Previous_Month_Revenue) / Previous_Month_Revenue) * 100
    END AS Revenue_Velocity_Percent
FROM RevenueWithLag
ORDER BY Month;

-- Query 3: Category Concentration Analysis
-- Purpose: Identify revenue concentration risk by category
WITH CategoryRevenue AS (
    SELECT 
        Product_Category,
        SUM(Revenue) AS Category_Revenue
    FROM Transactions
    GROUP BY Product_Category
),
TotalRevenue AS (
    SELECT SUM(Revenue) AS Total_Revenue
    FROM Transactions
),
RankedCategories AS (
    SELECT 
        cr.Product_Category,
        cr.Category_Revenue,
        tr.Total_Revenue,
        (cr.Category_Revenue / tr.Total_Revenue) * 100 AS Revenue_Percentage,
        ROW_NUMBER() OVER (ORDER BY cr.Category_Revenue DESC) AS Category_Rank
    FROM CategoryRevenue cr
    CROSS JOIN TotalRevenue tr
)
SELECT 
    Product_Category,
    Category_Revenue,
    Revenue_Percentage,
    Category_Rank,
    CASE 
        WHEN Category_Rank <= 3 THEN 'Top 3 Category'
        ELSE 'Other'
    END AS Category_Group
FROM RankedCategories
ORDER BY Category_Revenue DESC;

-- Query 4: Price Elasticity Analysis
-- Purpose: Analyze relationship between price changes and volume
WITH PriceVolume AS (
    SELECT 
        Product_SKU,
        Transaction_Date,
        Unit_Price,
        SUM(Quantity) AS Total_Quantity,
        AVG(Unit_Price) OVER (PARTITION BY Product_SKU ORDER BY Transaction_Date 
            ROWS BETWEEN 30 PRECEDING AND CURRENT ROW) AS Moving_Avg_Price
    FROM Transactions
    GROUP BY Product_SKU, Transaction_Date, Unit_Price
)
SELECT 
    Product_SKU,
    COUNT(*) AS Number_of_Transactions,
    AVG(Unit_Price) AS Avg_Price,
    STDEV(Unit_Price) AS Price_StdDev,
    AVG(Total_Quantity) AS Avg_Quantity,
    STDEV(Total_Quantity) AS Quantity_StdDev,
    -- Simplified elasticity indicator: negative correlation suggests elastic demand
    CASE 
        WHEN AVG(Unit_Price) > LAG(AVG(Unit_Price)) OVER (PARTITION BY Product_SKU ORDER BY Transaction_Date) 
            AND AVG(Total_Quantity) < LAG(AVG(Total_Quantity)) OVER (PARTITION BY Product_SKU ORDER BY Transaction_Date)
        THEN 'Elastic'
        ELSE 'Inelastic'
    END AS Elasticity_Type
FROM PriceVolume
GROUP BY Product_SKU, Transaction_Date
ORDER BY Product_SKU, Transaction_Date;

-- Query 5: Contribution Margin by Product
-- Purpose: Calculate profitability metrics for each product
SELECT 
    p.Product_SKU,
    p.Product_Name,
    p.Category,
    SUM(t.Revenue) AS Total_Revenue,
    SUM(t.Variable_Cost) AS Total_Variable_Cost,
    SUM(t.Revenue - t.Variable_Cost) AS Total_Contribution,
    (SUM(t.Revenue - t.Variable_Cost) / NULLIF(SUM(t.Revenue), 0)) * 100 AS Contribution_Margin_Percent,
    COUNT(t.Transaction_ID) AS Number_of_Sales,
    SUM(t.Quantity) AS Units_Sold,
    p.Stock_Level,
    CASE 
        WHEN (SUM(t.Revenue - t.Variable_Cost) / NULLIF(SUM(t.Revenue), 0)) * 100 > 60 THEN 'High Margin'
        WHEN (SUM(t.Revenue - t.Variable_Cost) / NULLIF(SUM(t.Revenue), 0)) * 100 BETWEEN 40 AND 60 THEN 'Medium Margin'
        ELSE 'Low Margin'
    END AS Margin_Category
FROM Products p
LEFT JOIN Transactions t ON p.Product_SKU = t.Product_SKU
GROUP BY p.Product_SKU, p.Product_Name, p.Category, p.Stock_Level
ORDER BY Contribution_Margin_Percent DESC;

-- ================================================================
-- SECTION 3: CUSTOMER SEGMENTATION & COHORT ANALYSIS
-- ================================================================

-- Query 6: RFM Analysis (Recency, Frequency, Monetary)
-- Purpose: Segment customers for targeted marketing
WITH CustomerRFM AS (
    SELECT 
        c.Customer_ID,
        c.Customer_Name,
        DATEDIFF(day, MAX(t.Transaction_Date), GETDATE()) AS Recency_Days,
        COUNT(t.Transaction_ID) AS Frequency,
        SUM(t.Revenue) AS Monetary
    FROM Customers c
    LEFT JOIN Transactions t ON c.Customer_ID = t.Customer_ID
    GROUP BY c.Customer_ID, c.Customer_Name
),
RFMScores AS (
    SELECT 
        Customer_ID,
        Customer_Name,
        Recency_Days,
        Frequency,
        Monetary,
        NTILE(5) OVER (ORDER BY Recency_Days) AS R_Score,
        NTILE(5) OVER (ORDER BY Frequency DESC) AS F_Score,
        NTILE(5) OVER (ORDER BY Monetary DESC) AS M_Score
    FROM CustomerRFM
)
SELECT 
    Customer_ID,
    Customer_Name,
    Recency_Days,
    Frequency,
    Monetary,
    R_Score,
    F_Score,
    M_Score,
    (R_Score + F_Score + M_Score) AS RFM_Total_Score,
    CASE 
        WHEN (R_Score + F_Score + M_Score) >= 12 THEN 'Champions'
        WHEN (R_Score + F_Score + M_Score) >= 9 THEN 'Loyal Customers'
        WHEN (R_Score + F_Score + M_Score) >= 6 THEN 'Potential Loyalists'
        WHEN R_Score >= 4 AND (F_Score + M_Score) <= 4 THEN 'At Risk'
        ELSE 'Needs Attention'
    END AS Customer_Segment
FROM RFMScores
ORDER BY RFM_Total_Score DESC;

-- Query 7: Cohort Analysis by Registration Month
-- Purpose: Track customer retention and value over time
WITH CustomerCohorts AS (
    SELECT 
        Customer_ID,
        FORMAT(Registration_Date, 'yyyy-MM') AS Cohort_Month,
        Registration_Date
    FROM Customers
),
CohortActivity AS (
    SELECT 
        cc.Cohort_Month,
        FORMAT(t.Transaction_Date, 'yyyy-MM') AS Activity_Month,
        COUNT(DISTINCT t.Customer_ID) AS Active_Customers,
        SUM(t.Revenue) AS Cohort_Revenue
    FROM CustomerCohorts cc
    JOIN Transactions t ON cc.Customer_ID = t.Customer_ID
    GROUP BY cc.Cohort_Month, FORMAT(t.Transaction_Date, 'yyyy-MM')
)
SELECT 
    Cohort_Month,
    Activity_Month,
    Active_Customers,
    Cohort_Revenue,
    DATEDIFF(month, CAST(Cohort_Month + '-01' AS DATE), CAST(Activity_Month + '-01' AS DATE)) AS Months_Since_Registration
FROM CohortActivity
ORDER BY Cohort_Month, Activity_Month;

-- ================================================================
-- SECTION 4: INVENTORY & OPERATIONAL ANALYTICS
-- ================================================================

-- Query 8: Inventory Turnover Ratio
-- Purpose: Measure inventory efficiency by product
SELECT 
    p.Product_SKU,
    p.Product_Name,
    p.Category,
    SUM(t.COGS) AS Total_COGS,
    AVG(p.Stock_Level) AS Avg_Inventory,
    SUM(t.COGS) / NULLIF(AVG(p.Stock_Level * p.Cost_of_Goods), 0) AS Inventory_Turnover_Ratio,
    365 / NULLIF((SUM(t.COGS) / NULLIF(AVG(p.Stock_Level * p.Cost_of_Goods), 0)), 0) AS Days_Inventory_Outstanding,
    CASE 
        WHEN SUM(t.COGS) / NULLIF(AVG(p.Stock_Level * p.Cost_of_Goods), 0) > 8 THEN 'Optimal'
        WHEN SUM(t.COGS) / NULLIF(AVG(p.Stock_Level * p.Cost_of_Goods), 0) BETWEEN 4 AND 8 THEN 'Acceptable'
        ELSE 'Slow Moving'
    END AS Turnover_Status
FROM Products p
LEFT JOIN Transactions t ON p.Product_SKU = t.Product_SKU
GROUP BY p.Product_SKU, p.Product_Name, p.Category
ORDER BY Inventory_Turnover_Ratio DESC;

-- Query 9: Basket Analysis (Market Basket / Cross-Sell Opportunities)
-- Purpose: Identify products frequently purchased together
WITH TransactionBaskets AS (
    SELECT 
        t1.Transaction_ID,
        t1.Product_Category AS Category_A,
        t2.Product_Category AS Category_B
    FROM Transactions t1
    JOIN Transactions t2 ON t1.Transaction_ID = t2.Transaction_ID 
        AND t1.Product_Category < t2.Product_Category
)
SELECT 
    Category_A,
    Category_B,
    COUNT(*) AS Co_Purchase_Count,
    (COUNT(*) * 100.0 / (SELECT COUNT(DISTINCT Transaction_ID) FROM Transactions)) AS Co_Purchase_Percentage
FROM TransactionBaskets
GROUP BY Category_A, Category_B
HAVING COUNT(*) > 1
ORDER BY Co_Purchase_Count DESC;

-- Query 10: Regional Performance Dashboard
-- Purpose: Compare performance metrics across regions
SELECT 
    Region,
    COUNT(DISTINCT Customer_ID) AS Total_Customers,
    COUNT(Transaction_ID) AS Total_Transactions,
    SUM(Revenue) AS Total_Revenue,
    AVG(Revenue) AS Avg_Transaction_Value,
    SUM(Revenue - Variable_Cost) AS Total_Profit,
    (SUM(Revenue - Variable_Cost) / NULLIF(SUM(Revenue), 0)) * 100 AS Profit_Margin_Percent,
    SUM(Marketing_Spend) AS Total_Marketing_Spend,
    SUM(Revenue) / NULLIF(SUM(Marketing_Spend), 0) AS Marketing_ROI,
    RANK() OVER (ORDER BY SUM(Revenue) DESC) AS Revenue_Rank
FROM Transactions
GROUP BY Region
ORDER BY Total_Revenue DESC;

-- ================================================================
-- SECTION 5: ADVANCED PREDICTIVE ANALYTICS
-- ================================================================

-- Query 11: Customer Churn Risk Prediction
-- Purpose: Identify customers at risk of churning
WITH CustomerActivity AS (
    SELECT 
        c.Customer_ID,
        c.Customer_Name,
        c.Customer_Segment,
        DATEDIFF(day, MAX(t.Transaction_Date), GETDATE()) AS Days_Since_Last_Purchase,
        COUNT(t.Transaction_ID) AS Total_Transactions,
        AVG(t.Revenue) AS Avg_Purchase_Value,
        DATEDIFF(day, MIN(t.Transaction_Date), MAX(t.Transaction_Date)) AS Customer_Lifespan_Days
    FROM Customers c
    LEFT JOIN Transactions t ON c.Customer_ID = t.Customer_ID
    GROUP BY c.Customer_ID, c.Customer_Name, c.Customer_Segment
)
SELECT 
    Customer_ID,
    Customer_Name,
    Customer_Segment,
    Days_Since_Last_Purchase,
    Total_Transactions,
    Avg_Purchase_Value,
    Customer_Lifespan_Days,
    CASE 
        WHEN Days_Since_Last_Purchase > 90 AND Total_Transactions > 3 THEN 'High Churn Risk'
        WHEN Days_Since_Last_Purchase > 60 AND Total_Transactions > 1 THEN 'Medium Churn Risk'
        WHEN Days_Since_Last_Purchase > 45 THEN 'Low Churn Risk'
        ELSE 'Active'
    END AS Churn_Risk_Category,
    Avg_Purchase_Value * Total_Transactions AS Potential_Lost_Revenue
FROM CustomerActivity
WHERE Days_Since_Last_Purchase > 30
ORDER BY 
    CASE 
        WHEN Days_Since_Last_Purchase > 90 AND Total_Transactions > 3 THEN 1
        WHEN Days_Since_Last_Purchase > 60 AND Total_Transactions > 1 THEN 2
        WHEN Days_Since_Last_Purchase > 45 THEN 3
        ELSE 4
    END,
    Potential_Lost_Revenue DESC;

-- Query 12: Seasonal Trend Analysis
-- Purpose: Identify seasonality patterns for forecasting
SELECT 
    DATEPART(year, Transaction_Date) AS Year,
    DATEPART(month, Transaction_Date) AS Month,
    DATENAME(month, Transaction_Date) AS Month_Name,
    Product_Category,
    SUM(Revenue) AS Monthly_Revenue,
    COUNT(Transaction_ID) AS Transaction_Count,
    AVG(SUM(Revenue)) OVER (PARTITION BY Product_Category) AS Category_Avg_Revenue,
    (SUM(Revenue) / NULLIF(AVG(SUM(Revenue)) OVER (PARTITION BY Product_Category), 0)) * 100 AS Seasonal_Index,
    CASE 
        WHEN (SUM(Revenue) / NULLIF(AVG(SUM(Revenue)) OVER (PARTITION BY Product_Category), 0)) * 100 > 120 THEN 'Peak Season'
        WHEN (SUM(Revenue) / NULLIF(AVG(SUM(Revenue)) OVER (PARTITION BY Product_Category), 0)) * 100 < 80 THEN 'Off Season'
        ELSE 'Normal Season'
    END AS Season_Category
FROM Transactions
GROUP BY DATEPART(year, Transaction_Date), DATEPART(month, Transaction_Date), DATENAME(month, Transaction_Date), Product_Category
ORDER BY Product_Category, Year, Month;

-- ================================================================
-- SECTION 6: EXECUTIVE SUMMARY QUERIES
-- ================================================================

-- Query 13: Top 10 Products by Revenue
SELECT TOP 10
    p.Product_SKU,
    p.Product_Name,
    p.Category,
    SUM(t.Revenue) AS Total_Revenue,
    SUM(t.Quantity) AS Units_Sold,
    AVG(t.Unit_Price) AS Avg_Price,
    SUM(t.Revenue - t.Variable_Cost) AS Total_Profit,
    (SUM(t.Revenue - t.Variable_Cost) / NULLIF(SUM(t.Revenue), 0)) * 100 AS Profit_Margin
FROM Products p
JOIN Transactions t ON p.Product_SKU = t.Product_SKU
GROUP BY p.Product_SKU, p.Product_Name, p.Category
ORDER BY Total_Revenue DESC;

-- Query 14: Customer Acquisition Efficiency
SELECT 
    FORMAT(Transaction_Date, 'yyyy-MM') AS Month,
    COUNT(DISTINCT CASE WHEN c.Registration_Date >= DATEADD(month, -1, Transaction_Date) 
        AND c.Registration_Date <= Transaction_Date THEN c.Customer_ID END) AS New_Customers,
    SUM(CASE WHEN c.Registration_Date >= DATEADD(month, -1, Transaction_Date) 
        AND c.Registration_Date <= Transaction_Date THEN t.Revenue ELSE 0 END) AS New_Customer_Revenue,
    SUM(t.Marketing_Spend) AS Total_Marketing_Spend,
    SUM(CASE WHEN c.Registration_Date >= DATEADD(month, -1, Transaction_Date) 
        AND c.Registration_Date <= Transaction_Date THEN t.Revenue ELSE 0 END) / 
        NULLIF(SUM(t.Marketing_Spend), 0) AS Customer_Acquisition_Efficiency
FROM Transactions t
JOIN Customers c ON t.Customer_ID = c.Customer_ID
GROUP BY FORMAT(Transaction_Date, 'yyyy-MM')
ORDER BY Month;

-- ================================================================
-- END OF QUERY LIBRARY
-- ================================================================
-- Note: These queries are designed to work with Microsoft Access
-- and SQL Server. Some syntax may need adjustment for other DBMS.
-- ================================================================
