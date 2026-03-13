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
-- 更改结束符以便创建存储过程
DELIMITER //
DROP PROCEDURE IF EXISTS GenerateProcRepData //

CREATE PROCEDURE GenerateProcRepData()
BEGIN
    DECLARE i INT DEFAULT 1;
    
    -- 定义变量匹配你的表结构 
    DECLARE v_Contract VARCHAR(50);
    DECLARE v_ContractName VARCHAR(50);
    DECLARE v_Ord INT;
    DECLARE v_OrdType VARCHAR(50);
    DECLARE v_SupplierName VARCHAR(50);
    DECLARE v_Orderdate DATE;
    DECLARE v_Exvat DECIMAL(10,2);
    DECLARE v_StatusCategory VARCHAR(50);
    
    -- 用于控制状态逻辑的随机数
    DECLARE rand_status INT;

    -- 循环生成 100 条数据 (你可以修改这个数字生成更多)
    WHILE i <= 100 DO
        -- 1. 生成唯一主键和基础文本 (Contract) 
        SET v_Contract = CONCAT('CTR-2025-', 10000 + i);
        SET v_ContractName = CONCAT('Facility Project ', CHAR(FLOOR(65 + (RAND() * 26))));
        
        -- 2. 生成订单号 (Ord) 和订单类型 (OrdType: 仅 M 和 S) 
        SET v_Ord = 80000 + i;
        SET v_OrdType = IF(RAND() > 0.5, 'M', 'S');
        
        -- 3. 随机分配供应商名称 
        SET v_SupplierName = ELT(FLOOR(1 + (RAND() * 5)), 'Alpha Maintenance', 'Beta FM Supplies', 'Global Tech Services', 'Nordic Builders', 'Eco Facility Sol');
        
        -- 4. 生成 2025-01-01 到 2025-12-31 之间的随机日期 
        SET v_Orderdate = DATE_ADD('2025-01-01', INTERVAL FLOOR(RAND() * 364) DAY);
        
        -- 5. 生成不含税金额 (Exvat) 
        SET v_Exvat = ROUND((RAND() * 8500) + 150, 2);
        
        -- 6. 业务逻辑映射：保持 Status 和 StatusCategory 强相关 
        SET v_StatusCategory = ELT(FLOOR(1 + (RAND() * 3)), 'Fully Invoiced', 'Open Orders','Part Invoiced');
        


        -- 插入数据 
        INSERT INTO ProcRep (Contract, ContractName, Ord, OrdType, SupplierName, Orderdate, Exvat, StatusCategory)
        VALUES (v_Contract, v_ContractName, v_Ord, v_OrdType, v_SupplierName, v_Orderdate, v_Exvat, v_StatusCategory);
        
        SET i = i + 1;
    END WHILE;
END //

-- 恢复默认结束符
DELIMITER ;

-- 清空表（防止重复运行主键冲突），然后调用存储过程生成数据
TRUNCATE TABLE ProcRep;
CALL GenerateProcRepData();

-- 检查生成的数据
SELECT * FROM ProcRep ORDER BY Orderdate DESC LIMIT 10;
