# SQL Schema for Dashboard Temporary Tables (MS Access)

## Create Tables

```sql
-- Dashboard Stats Table
CREATE TABLE tblDashboardStats (
    StatID AUTOINCREMENT PRIMARY KEY,
    StatName TEXT(50),
    Value TEXT(50),
    Trend TEXT(10),
    TrendUp YESNO
);

-- Stats Mini Chart Data Table
CREATE TABLE tblStatsMiniChartData (
    DataID AUTOINCREMENT PRIMARY KEY,
    StatID LONG,
    DataValue INTEGER,
    DataOrder INTEGER,
    FOREIGN KEY (StatID) REFERENCES tblDashboardStats(StatID)
);

-- Sales Chart Data Table
CREATE TABLE tblSalesChartData (
    DataID AUTOINCREMENT PRIMARY KEY,
    MonthOrder INTEGER,
    OnlineValue INTEGER,
    OfflineValue INTEGER
);

-- Donut Chart Table
CREATE TABLE tblDonutChart (
    SegmentID AUTOINCREMENT PRIMARY KEY,
    SegmentName TEXT(50),
    DataValue INTEGER
);

-- Customer Chart Data Table
CREATE TABLE tblCustomerChartData (
    DataID AUTOINCREMENT PRIMARY KEY,
    MonthOrder INTEGER,
    LoyalValue INTEGER,
    NewValue INTEGER
);

-- Recent Orders Table
CREATE TABLE tblRecentOrders (
    OrderID AUTOINCREMENT PRIMARY KEY,
    Product TEXT(100),
    OrderDate TEXT(50),
    Price TEXT(20),
    Status TEXT(50)
);
```

## Sample Data Insert (Optional)

```sql
-- Insert Dashboard Stats
INSERT INTO tblDashboardStats (StatName, Value, Trend, TrendUp) VALUES
('Total orders', '947', '+5%', True),
('Total sales', '$28,407', '+3%', True),
('Online sessions', '54,778', '+2%', True),
('Average order', '$89.99', '-1%', False);

-- Insert Stats Mini Chart Data (for Total orders - StatID=1)
INSERT INTO tblStatsMiniChartData (StatID, DataValue, DataOrder) VALUES
(1, 5, 1), (1, 7, 2), (1, 6, 3), (1, 8, 4), (1, 9, 5), (1, 12, 6);

-- Insert Stats Mini Chart Data (for Total sales - StatID=2)
INSERT INTO tblStatsMiniChartData (StatID, DataValue, DataOrder) VALUES
(2, 8, 1), (2, 9, 2), (2, 7, 3), (2, 10, 4), (2, 11, 5), (2, 14, 6);

-- Insert Stats Mini Chart Data (for Online sessions - StatID=3)
INSERT INTO tblStatsMiniChartData (StatID, DataValue, DataOrder) VALUES
(3, 6, 1), (3, 8, 2), (3, 7, 3), (3, 9, 4), (3, 10, 5), (3, 13, 6);

-- Insert Stats Mini Chart Data (for Average order - StatID=4)
INSERT INTO tblStatsMiniChartData (StatID, DataValue, DataOrder) VALUES
(4, 12, 1), (4, 10, 2), (4, 11, 3), (4, 9, 4), (4, 8, 5), (4, 6, 6);

-- Insert Sales Chart Data (12 months)
INSERT INTO tblSalesChartData (MonthOrder, OnlineValue, OfflineValue) VALUES
(1, 22, 7), (2, 19, 5), (3, 26, 8), (4, 23, 8), (5, 21, 9), (6, 25, 6),
(7, 23, 5), (8, 17, 4), (9, 20, 5), (10, 19, 6), (11, 25, 8), (12, 26, 3);

-- Insert Donut Chart Data
INSERT INTO tblDonutChart (SegmentName, DataValue) VALUES
('Online sales', 55), ('Offline sales', 30), ('Returns', 15);

-- Insert Customer Chart Data (6 months)
INSERT INTO tblCustomerChartData (MonthOrder, LoyalValue, NewValue) VALUES
(1, 500, 400), (2, 2300, 1800), (3, 1500, 2400), (4, 2700, 1000), (5, 1800, 1800), (6, 2800, 1700);

-- Insert Recent Orders
INSERT INTO tblRecentOrders (Product, OrderDate, Price, Status) VALUES
('Rust Linen Blazer', 'Jan 25, 2025', '$149.99', 'Shipping'),
('Crop Tank', 'Jan 25, 2025', '$49.99', 'Received'),
('Oversized Blazer', 'Jan 24, 2025', '$185.99', 'Received');
```

## Notes

- All tables use AutoNumber for primary keys
- DataOrder fields maintain array order for JSON serialization
- StatName values must match: "Total orders", "Total sales", "Online sessions", "Average order"
- SegmentName values must match: "Online sales", "Offline sales", "Returns"
- MonthOrder: 1-12 for sales chart (Jan-Dec), 1-6 for customer chart (Jul-Dec)
