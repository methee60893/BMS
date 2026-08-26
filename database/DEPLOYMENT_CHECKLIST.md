# BMS SQL Server Deployment Checklist

## 1. ก่อน deploy

- ยืนยันว่า SQL Server version รองรับ `CREATE OR ALTER`, `STRING_SPLIT`, `MERGE`
- เตรียมสิทธิ์สำหรับ create database/schema/object
- สำรองฐานข้อมูลเดิมก่อน หากเป็นการลงทับ environment ที่มีข้อมูล
- ตรวจสอบ connection string ใน [Web.config](/D:/CIE/BMS/Web.config) ให้ชี้ instance เป้าหมายถูกต้อง
- เตรียมข้อมูลจริงสำหรับ master data โดยแก้ไฟล์ [04_seed_master_data_template.sql](/D:/CIE/BMS/database/04_seed_master_data_template.sql)
- ยืนยันว่า config ของ environment มี request envelope 25 MB (`maxRequestLength="25600"` และ `maxAllowedContentLength="26214400"`) โดยไม่คัดลอก secret จาก environment อื่น
- สร้าง `App_Data\DraftOtbJobs` และให้ Application Pool identity มีสิทธิ์ **Modify** เฉพาะโฟลเดอร์นี้

## 2. ลำดับการรันสคริปต์

รันตามนี้ใน SSMS หรือ sqlcmd:

> สำหรับอัปเกรด PRD เดิมเพื่อ Sprint 1 ให้สำรองฐาน, รอให้งาน Draft OTB จบ, หยุด App Pool แล้วรัน [08_deploy_sprint1_prd.sql](/D:/CIE/BMS/database/08_deploy_sprint1_prd.sql) ผ่าน TCP ที่ `10.3.152.155` ก่อน deploy application ไม่ต้องรันชุดสร้างฐานใหม่ข้อ 1-5 ซ้ำบน PRD

- ต้องเห็นผลลัพธ์ `DeploymentStatus = PASS` ก่อนนำ application ขึ้น หากสคริปต์หยุดเพราะ Draft key ผิด/ซ้ำ หรือมี job/claim ค้าง ให้แก้สาเหตุแล้วรันไฟล์เดิมใหม่ ห้ามข้าม guard
- สคริปต์ไม่แก้ค่าข้อมูลธุรกิจเดิม แต่การสร้าง index จะอ่านและ lock ตาราง Draft รวมถึงใช้พื้นที่ transaction log จึงควรรันใน maintenance window
- ยืนยันว่า Application DB account มีสิทธิ์อ่าน/เพิ่ม/แก้ไข `dbo.Draft_OTB_Background_Job` และอ่าน/เพิ่ม/ลบ `dbo.Draft_OTB_Approval_Claim` ตามแนวทางสิทธิ์ของ PRD

1. [00_create_database.sql](/D:/CIE/BMS/database/00_create_database.sql)
2. [01_create_tables.sql](/D:/CIE/BMS/database/01_create_tables.sql)
3. [02_create_views.sql](/D:/CIE/BMS/database/02_create_views.sql)
4. [03_create_stored_procedures.sql](/D:/CIE/BMS/database/03_create_stored_procedures.sql)
5. [04_seed_master_data_template.sql](/D:/CIE/BMS/database/04_seed_master_data_template.sql)
6. [07_create_draft_otb_background_jobs.sql](/D:/CIE/BMS/database/07_create_draft_otb_background_jobs.sql)
7. [05_post_deploy_verify.sql](/D:/CIE/BMS/database/05_post_deploy_verify.sql)

## 3. ตรวจสอบหลัง deploy

รัน query เช็ก object:

```sql
USE [BMS];

SELECT name, type_desc
FROM sys.objects
WHERE name IN (
    'MS_Month','MS_Year','MS_Company','MS_Category','MS_Segment','MS_Brand','MS_Vendor','MS_CCY','MS_Version',
    'MS_User','MS_Role','MS_Menu','Map_User_Role','Map_Role_Permission',
    'Template_Upload_Draft_OTB','OTB_Transaction','OTB_Switching_Transaction','Draft_PO_Transaction','Draft_OTB_Background_Job','Draft_OTB_Approval_Claim',
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
- ทดสอบ Draft OTB ที่ 300, 301, 10,000 และ 15,000 แถว รวมทั้งยืนยันว่า 15,001 แถวถูกปฏิเสธทั้งไฟล์
- ทดสอบไฟล์ที่มีบางแถวผิด: ต้องไม่มีแถวใดถูกบันทึกและต้องดาวน์โหลดรายงานข้อผิดพลาดได้
- ระหว่าง upload/approve ให้ Refresh และปิด/เปิดหน้าใหม่ แล้วตรวจว่า progress และ Job ID เดิมกลับมา
- ทดสอบ retry หลังจำลอง network timeout: ต้องได้ Job ID เดิมจาก client request ID เดิมและต้องไม่มีงานซ้ำ

## 5. จุดที่ต้องระวัง

- `MS_Vendor` ต้องผูก `VendorCode + SegmentCode` ให้ตรง ไม่เช่นนั้น dropdown และ join บางหน้าจะไม่ขึ้นชื่อ vendor
- `MS_Version` ต้องมีอย่างน้อย `A1` สำหรับ `Original` และ `R1` สำหรับ `Revise`
- ถ้า SAP ส่ง segment มาในรูปแบบมี wrapper เช่น `(S01)` หรือ `OS010` ระบบมี logic แปลงบางส่วนในโค้ด แต่ master data ยังต้องเก็บ code มาตรฐานฝั่ง BMS
- ถ้า production มีข้อมูล volume สูง ควรทบทวน index เพิ่มตาม query plan จริงอีกครั้ง

## 6. งานปฏิบัติการสำหรับ Background Job

- หน้าจอปิดหรือ Refresh ได้โดยงานยังทำต่อใน server process และสถานะถูกเก็บในฐานข้อมูล
- หลีกเลี่ยง deploy, IIS reset และ Application Pool recycle ขณะมีสถานะ `Queued`, `ReceivingUpload`, `Validating`, `QueuedForSave`, `Saving`, `PreparingSap`, `SendingToSap` หรือ `SavingApproval`; worker แบบ in-process ไม่รับประกันว่าจะทำต่อหลัง recycle
- หาก recycle เกิดช่วงก่อนส่ง SAP ให้ตรวจสถานะและเริ่มใหม่ได้เฉพาะเมื่อยืนยันว่าไม่มี side effect; หากอยู่ช่วงส่ง/หลังส่ง SAP ต้องจัดเป็น `ReconciliationRequired` และห้าม retry อัตโนมัติ
- เมื่อพบ `ReconciliationRequired` ให้บันทึก Job ID, ตรวจ payload/ผลตอบกลับ, ยืนยันสถานะเอกสารใน SAP และเทียบข้อมูล BMS จากนั้นจึงให้ผู้มีสิทธิ์รับทราบบนหน้าจอ
- ตั้ง scheduled housekeeping แยกต่างหาก: ลบเฉพาะงาน terminal ที่พ้น retention และไฟล์ orphan ใน `App_Data\DraftOtbJobs`; ห้ามลบงาน active หรือ `ReconciliationRequired` ที่ยังไม่ปิดการตรวจสอบ
- ตรวจ disk space และสิทธิ์ Modify ของ `App_Data\DraftOtbJobs` หลัง deploy ทุก environment

## 7. ตัวอย่างคำสั่ง sqlcmd

```powershell
sqlcmd -S YOUR_SERVER -E -i D:\CIE\BMS\database\00_create_database.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\01_create_tables.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\02_create_views.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\03_create_stored_procedures.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\04_seed_master_data_template.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\07_create_draft_otb_background_jobs.sql
sqlcmd -S YOUR_SERVER -E -d BMS -i D:\CIE\BMS\database\05_post_deploy_verify.sql
```

ถ้าใช้ SQL Login:

```powershell
sqlcmd -S YOUR_SERVER -U sa -P YOUR_PASSWORD -d BMS -i D:\CIE\BMS\database\04_seed_master_data_template.sql
```

## 8. Python Compare

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
