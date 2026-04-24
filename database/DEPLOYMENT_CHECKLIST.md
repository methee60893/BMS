# BMS SQL Server Deployment Checklist

## 1. ก่อน deploy

- ยืนยันว่า SQL Server version รองรับ `CREATE OR ALTER`, `STRING_SPLIT`, `MERGE`
- เตรียมสิทธิ์สำหรับ create database/schema/object
- สำรองฐานข้อมูลเดิมก่อน หากเป็นการลงทับ environment ที่มีข้อมูล
- ตรวจสอบ connection string ใน [Web.config](/D:/CIE/BMS/Web.config) ให้ชี้ instance เป้าหมายถูกต้อง
- เตรียมข้อมูลจริงสำหรับ master data โดยแก้ไฟล์ [04_seed_master_data_template.sql](/D:/CIE/BMS/database/04_seed_master_data_template.sql)

## 2. ลำดับการรันสคริปต์

รันตามนี้ใน SSMS หรือ sqlcmd:

1. [00_create_database.sql](/D:/CIE/BMS/database/00_create_database.sql)
2. [01_create_tables.sql](/D:/CIE/BMS/database/01_create_tables.sql)
3. [02_create_views.sql](/D:/CIE/BMS/database/02_create_views.sql)
4. [03_create_stored_procedures.sql](/D:/CIE/BMS/database/03_create_stored_procedures.sql)
5. [04_seed_master_data_template.sql](/D:/CIE/BMS/database/04_seed_master_data_template.sql)
6. [05_post_deploy_verify.sql](/D:/CIE/BMS/database/05_post_deploy_verify.sql)

## 3. ตรวจสอบหลัง deploy

รัน query เช็ก object:

```sql
USE [BMS];

SELECT name, type_desc
FROM sys.objects
WHERE name IN (
    'MS_Month','MS_Year','MS_Company','MS_Category','MS_Segment','MS_Brand','MS_Vendor','MS_CCY','MS_Version',
    'MS_User','MS_Role','MS_Menu','Map_User_Role','Map_Role_Permission',
    'Template_Upload_Draft_OTB','OTB_Transaction','OTB_Switching_Transaction','Draft_PO_Transaction',
    'Actual_PO_Staging','Actual_PO_Summary',
    'View_OTB_Draft','View_UserRole',
    'SP_Get_Actual_PO_List','SP_Search_Approved_OTB','SP_Search_SWitch_OTB',
    'SP_Sync_Actual_PO_Summary','SP_Auto_Match_Actual_Draft','SP_Approve_Draft_OTB','SP_Deleted_Draft_OTB',
    'SP_Sync_User_From_AD','SP_Get_Users_List','SP_Admin_Save_User'
)
ORDER BY type_desc, name;
```

- ตรวจสอบว่า master data มีอย่างน้อย company/category/segment/brand/vendor/month/version
- ทดสอบ view:

```sql
SELECT TOP (10) *
FROM dbo.View_OTB_Draft
ORDER BY RunNo DESC;
```

- ทดสอบ stored procedures:

```sql
EXEC dbo.SP_Search_Approved_OTB;
EXEC dbo.SP_Search_SWitch_OTB;
EXEC dbo.SP_Get_Actual_PO_List;
```

- ทดสอบ post-deploy verify:

```sql
:r D:\CIE\BMS\database\05_post_deploy_verify.sql
```

## 4. Smoke test เชิงธุรกิจ

- เพิ่ม master data จริงให้ครบตามธุรกิจ
- ทดสอบ upload Draft OTB 1 รายการ
- ทดสอบ approve Draft OTB 1 รายการ
- ทดสอบ create Draft PO 1 รายการ
- ทดสอบ sync Actual PO staging/summary
- ทดสอบ auto match และ manual match
- ทดสอบหน้ารายงานที่ใช้ filter Company/Category/Segment/Brand/Vendor

## 5. จุดที่ต้องระวัง

- `MS_Vendor` ต้องผูก `VendorCode + SegmentCode` ให้ตรง ไม่เช่นนั้น dropdown และ join บางหน้าจะไม่ขึ้นชื่อ vendor
- `MS_Version` ต้องมีอย่างน้อย `A1` สำหรับ `Original` และ `R1` สำหรับ `Revise`
- ถ้า SAP ส่ง segment มาในรูปแบบมี wrapper เช่น `(S01)` หรือ `OS010` ระบบมี logic แปลงบางส่วนในโค้ด แต่ master data ยังต้องเก็บ code มาตรฐานฝั่ง BMS
- ถ้า production มีข้อมูล volume สูง ควรทบทวน index เพิ่มตาม query plan จริงอีกครั้ง

## 6. ตัวอย่างคำสั่ง sqlcmd

```powershell
sqlcmd -S YOUR_SERVER -E -i D:\CIE\BMS\database\00_create_database.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\01_create_tables.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\02_create_views.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\03_create_stored_procedures.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\04_seed_master_data_template.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\05_post_deploy_verify.sql
```

ถ้าใช้ SQL Login:

```powershell
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -d BMS -i D:\CIE\BMS\database\04_seed_master_data_template.sql
```

## 7. Python Compare

แบบ Windows auth:

```powershell
python D:\CIE\BMS\database\compare_excel_to_db.py --server YOUR_SERVER
```

แบบ SQL auth:

```powershell
python D:\CIE\BMS\database\compare_excel_to_db.py --server YOUR_SERVER --database BMS --username sa --password YOUR_PASSWORD
```

แบบส่ง connection string ตรง:

```powershell
python D:\CIE\BMS\database\compare_excel_to_db.py --connection-string "DRIVER={ODBC Driver 17 for SQL Server};SERVER=YOUR_SERVER;DATABASE=BMS;Trusted_Connection=yes;TrustServerCertificate=yes;"
```
