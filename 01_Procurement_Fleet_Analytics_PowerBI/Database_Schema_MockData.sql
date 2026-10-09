-- Create ProcRep Table
CREATE TABLE ProcRep (
	Contract VARCHAR(50) PRIMARY KEY,
    ContractName VARCHAR(50),
    Ord INT,
    OrdType VARCHAR(50),
    SupplierName VARCHAR(50),
    Orderdate Date, 
    Exvat Decimal(10,2),
    StatusCategory VARCHAR(50)
);


-- Create dummy data for this table
-- Change delimiter to create stored procedure
DELIMITER //
DROP PROCEDURE IF EXISTS GenerateProcRepData //

CREATE PROCEDURE GenerateProcRepData()
BEGIN
    DECLARE i INT DEFAULT 1;
    
    -- Declare variables to match the table structure
    DECLARE v_Contract VARCHAR(50);
    DECLARE v_ContractName VARCHAR(50);
    DECLARE v_Ord INT;
    DECLARE v_OrdType VARCHAR(50);
    DECLARE v_SupplierName VARCHAR(50);
    DECLARE v_Orderdate DATE;
    DECLARE v_Exvat DECIMAL(10,2);
    DECLARE v_StatusCategory VARCHAR(50);
    
    -- Random number to control status logic
    DECLARE rand_status INT;

    -- Loop to generate 100 rows of data
    WHILE i <= 100 DO
        -- 1. Generate unique primary key and base string (Contract)
        SET v_Contract = CONCAT('CTR-2025-', 10000 + i);
        SET v_ContractName = CONCAT('Facility Project ', CHAR(FLOOR(65 + (RAND() * 26))));
        
        -- 2. Generate order number (Ord) and order type (OrdType: M and S)
        SET v_Ord = 80000 + i;
        SET v_OrdType = IF(RAND() > 0.5, 'M', 'S');
        
        -- 3. Randomly assign supplier name 
        SET v_SupplierName = ELT(FLOOR(1 + (RAND() * 5)), 'Alpha Maintenance', 'Beta FM Supplies', 'Global Tech Services', 'Nordic Builders', 'Eco Facility Sol');
        
        -- 4. Generate a random date between 2025-01-01 and 2025-12-31
        SET v_Orderdate = DATE_ADD('2025-01-01', INTERVAL FLOOR(RAND() * 364) DAY);
        
        -- 5. Generate amount excluding VAT
        SET v_Exvat = ROUND((RAND() * 8500) + 150, 2);
        
        -- 6. Business logic mapping: Keep Status and StatusCategory strongly correlated
        SET v_StatusCategory = ELT(FLOOR(1 + (RAND() * 3)), 'Fully Invoiced', 'Open Orders','Part Invoiced');
        

-- Insert data
        INSERT INTO ProcRep (Contract, ContractName, Ord, OrdType, SupplierName, Orderdate, Exvat, StatusCategory)
        VALUES (v_Contract, v_ContractName, v_Ord, v_OrdType, v_SupplierName, v_Orderdate, v_Exvat, v_StatusCategory);
        
        SET i = i + 1;
    END WHILE;
END //

-- Reset default delimiter
DELIMITER ;

-- Clear table to prevent primary key conflicts on re-execution, then call procedure to generate data
TRUNCATE TABLE ProcRep;
CALL GenerateProcRepData();

-- Verify generated data
SELECT * FROM ProcRep ORDER BY Orderdate DESC LIMIT 10;
