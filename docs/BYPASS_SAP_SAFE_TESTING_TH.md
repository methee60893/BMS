# วิธีทดสอบ Bypass SAP โดยไม่แตะข้อมูล Production

เอกสารนี้ใช้สำหรับทดสอบ flow ใหม่ของ Switch Upload และ Draft OTB Approval ที่ bypass SAP แล้วบันทึกลงฐานข้อมูล BMS โดยตรง

## หลักที่ต้องยึด

- ห้ามทดสอบ save/approve/upload กับ `Initial Catalog=BMS` ที่เป็น production
- ห้ามใช้ SQL server ตัวเดียวกับ production สำหรับทดสอบ ยกเว้น DBA อนุมัติและแยก database ชัดเจน
- ให้ใช้ database แยก เช่น `BMS_TEST`, `BMS_UAT`, หรือ `BMS_SANDBOX`
- ก่อนเปิดหน้าเว็บเพื่อทดสอบ ต้องเช็ก `Web.config` ว่า `BMSConnectionString` ไม่ได้ชี้ production

## สิ่งที่ตรวจแล้วใน repo

- `Web.config` ปัจจุบันชี้ `Initial Catalog=BMS` บน server production
- มี template สำหรับ test ที่ `Web.Test.config.example`
- มี script เตรียม test environment:
  - `run_test_step_by_step.ps1`
  - `run_full_isolated_test.ps1`
- มี SQL test scripts อยู่ที่ `database\test`
- มี mock SAP server อยู่ที่ `test_support\mock_sap_server.py`

## Safety guard ที่เพิ่มไว้

เพิ่มไฟล์ `test_support\isolated_test_safety.ps1` และผูกเข้ากับ script ทดสอบแล้ว

guard นี้จะหยุดการทำงานถ้า:

- `-TestDatabase` เป็น `BMS`
- `-TestDatabase` ตรงกับ database ปัจจุบันใน `Web.config` และชื่อนั้นไม่ใช่ test/dev/uat/sandbox
- `-SqlServer` ตรงกับ server ปัจจุบันใน `Web.config` โดยไม่ได้ระบุ `-AllowProductionServer`
- ยังใช้ค่า placeholder `YOUR_SERVER`

ดังนั้นถ้าเผลอสั่ง test script ไปหา production server/script จะหยุดก่อนแก้ config หรือรัน SQL

## วิธีทดสอบที่ปลอดภัยที่สุด

1. เตรียม SQL Server แยกจาก production เช่น local SQL Server, SQL Express, หรือ UAT SQL Server

2. สร้าง/seed database test ด้วย script ใน repo

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -RunSql -SqlServer YOUR_TEST_SQL_SERVER -TestDatabase BMS_TEST -DbUser YOUR_TEST_USER -DbPassword YOUR_TEST_PASSWORD
```

3. ถ้าอยากดูก่อนว่า script จะทำอะไร ให้ใช้ dry run

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -DryRun -RunSql -SqlServer YOUR_TEST_SQL_SERVER -TestDatabase BMS_TEST -DbUser YOUR_TEST_USER -DbPassword YOUR_TEST_PASSWORD
```

4. ตรวจสถานะ config ก่อนเปิดเว็บ

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_test_step_by_step.ps1 status
```

ต้องเห็นว่า `BMSConnectionString` ชี้ `BMS_TEST` หรือ database ทดสอบเท่านั้น

5. ทดสอบ flow ในเว็บ

- Switch Upload: upload Excel, validate/preview, save
- Extra Upload: upload Excel, validate/preview, save
- Draft OTB: upload draft, tick checkbox, submit/approve

6. ตรวจข้อมูลใน test DB เท่านั้น

- `OTB_Switching_Transaction`
- `OTB_Transaction`
- `Template_Upload_Draft_OTB`

7. เมื่อทดสอบเสร็จ ให้ restore config กลับ

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -RestoreOnly
```

## กรณีต้องใช้ SQL server เดียวกับ production จริง ๆ

ไม่แนะนำ แต่ถ้า DBA ยืนยันว่าอนุญาตให้มี database ทดสอบบน server เดียวกัน ให้ใช้ database แยกเท่านั้น เช่น `BMS_TEST` และต้องระบุ flag นี้เอง:

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -RunSql -SqlServer PROD_SQL_SERVER -TestDatabase BMS_TEST -DbUser TEST_USER -DbPassword TEST_PASSWORD -AllowProductionServer
```

ห้ามใช้ `-TestDatabase BMS` ทุกกรณี

## วิธีทดสอบตอนนี้โดยไม่แตะ DB เลย

ในสถานะที่ยังไม่มี test DB สามารถทำได้เฉพาะ non-write checks:

- build project เพื่อตรวจ compile
- ตรวจ static diff ว่าไม่มีการเรียก SAP API ใน flow ที่ bypass แล้ว
- ตรวจ mapping logic จาก code review
- ใช้หน้า upload เฉพาะขั้น preview/validate ถ้ามั่นใจว่า handler นั้นยังไม่ save DB ในขั้น preview

สำหรับ Draft OTB ไม่ควรทดสอบผ่านหน้าจอ production เพราะ flow upload/approve มีการเขียน DB
