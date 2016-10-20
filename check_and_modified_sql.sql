
# Availablibity Check

UPDATE `zpar_master_data` SET `Availability_Check_change` = 'KP' WHERE `MRP Type` = 'ZS';
UPDATE `zpar_master_data` SET `Availability_Check_change` = 'Z4' WHERE `Material Type` = 'FERT' AND (`Plant` = 5578 OR `Plant` = 0538 OR `Plant` = 9216);
UPDATE `zpar_master_data` SET `Availability_Check_change` = '01' WHERE `Material Type` = 'ROH';
UPDATE `zpar_master_data` SET `Availability_Check_change` = '02' WHERE `Material Type` = 'FERT' AND `Plant` != 5578 AND `Plant` != 0538 AND `Plant` != 9216;
UPDATE `zpar_master_data` SET `Availability_Check_change` = '02' WHERE `Material Type` = 'HALB';

# Maximum stock level

UPDATE `zpar_master_data` SET `Maximum stock level change` = 0 WHERE `Material Type` = 'HALB' OR `Material Type` = 'FERT';
UPDATE `zpar_master_data` SET `Maximum stock level change` = 30000 WHERE `MRP Controller` = 'F04' AND (`Material Type` = 'HALB' OR `Material Type` = 'FERT');

# Maximum Lot size

UPDATE `zpar_master_data` SET `Maximum Lot size change`='POSS' WHERE `Material Type` = 'FERT' OR `Material Type` = 'HALB';

# Minimum Lot size

UPDATE `zpar_master_data` SET `Minimum Lot size change` = 'POSS' WHERE `Material Type` = 'FERT' OR `Material Type` = 'HALB';

# Procurement Type

UPDATE `zpar_master_data` SET `Procurement Type change` = 'F' WHERE `Material Type` = 'ROH' OR `Material Type` = 'UNBW';  
UPDATE `zpar_master_data` SET `Procurement Type change` = 'F' WHERE `Material Type` = 'FERT';
UPDATE `zpar_master_data` SET `Procurement Type change` = 'X' WHERE `Material Type` = 'FERT' AND `MRP Type` = 'PD';

# Quota Arrangement us

UPDATE `zpar_master_data` SET `Quota Arrangement us change` = 3;

# InhseProdn Tiem

UPDATe `zpar_master_data` SET `InhseProdnTime change` = 0;

# Goods receipt procc

# 当前字段记录中没有空值
UPDATE `zpar_master_data` SET `Goods receipt procc change` = 0 WHERE `Material Type` = 'HALB';
UPDATE `zpar_master_data` SET `Goods receipt procc change` = 0 WHERE `Goods receipt procc change` = '' AND `Material Type` = 'HERT' AND (`Plant` = 0386 OR `Plant` = 1864);
UPDATE `zpar_master_data` SET `Goods receipt procc change` = 1 WHERE `Goods receipt procc change` = '' AND `Material Type` = 'HERT' AND (`Plant` = 5578 OR `Plant` = 0538 OR `Plant` = 9216);

# Schedule margin key

UPDATE `zpar_master_data` SET `Schedule Margin key change` = '000' WHERE `Material Type` = 'FERT' OR `Material Type` = 'HALB';

# Period Indicator

UPDATE `zpar_master_data` SET `Period Indicator change` = 'M';

# Planned delivery tim

UPDATE `zpar_master_data` SET `Planned delivery tim change` = 0;
UPDATE `zpar_master_data` SET `Planned delivery tim change` = 1 WHERE `Material Type` = 'HALB' AND (`Plant` = 0538 OR `Plant` = 5578 OR `Plant` = 9216);

# Selection Method 
UPDATE `zpar_master_data` SET `Selection Method change` = 3 WHERE `Material Type` = 'FERT' OR `Material Type` = 'HALB';
UPDATE `zpar_master_data` SET `Selection Method change` = '' WHERE `Material Type` = 'ROH' OR `Material Type` = 'UNBW';

# Prod.Sched.Profile

UPDATE `zpar_master_data` SET `Prod.Sched.Profile change` = 'PI01' WHERE `Material Type` = 'FERT' OR `Material Type` = 'HALB';
UPDATE `zpar_master_data` SET `Prod.Sched.Profile change` = '' WHERE `Material Type` = 'ROH' OR `Material Type` = 'UNBW';