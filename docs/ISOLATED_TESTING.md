# Isolated Testing Plan

เป้าหมายคือทดสอบ BMS โดยไม่กระทบ `BMS` database หลัก และไม่ยิง API ไป SAP จริง

## หลักการแยก environment

- ใช้ฐานข้อมูลแยกชื่อ `BMS_TEST`
- ใช้ `Web.Test.config.example` เป็น template สำหรับ config ฝั่ง test
- ใช้ mock SAP server ที่ [mock_sap_server.py](/D:/CIE/BMS/test_support/mock_sap_server.py)
- ไม่แก้ `Web.config` ตัวหลักโดยตรง ถ้าจะทดสอบให้ copy เป็นไฟล์ทำงานชั่วคราวก่อน

## ส่วนประกอบ

- SQL สำหรับสร้าง test database: [00_create_database_test.sql](/D:/CIE/BMS/database/00_create_database_test.sql)
- SQL schema/seeds เดิม:
  - [01_create_tables.sql](/D:/CIE/BMS/database/01_create_tables.sql)
  - [02_create_views.sql](/D:/CIE/BMS/database/02_create_views.sql)
  - [03_create_stored_procedures.sql](/D:/CIE/BMS/database/03_create_stored_procedures.sql)
  - [04_seed_master_data_template.sql](/D:/CIE/BMS/database/04_seed_master_data_template.sql)
- SQL verify:
  - [05_post_deploy_verify.sql](/D:/CIE/BMS/database/05_post_deploy_verify.sql)
- Mock SAP:
  - [mock_sap_server.py](/D:/CIE/BMS/test_support/mock_sap_server.py)
  - [po_sample.json](/D:/CIE/BMS/test_support/mock_data/po_sample.json)

## วิธีทดสอบแบบไม่กระทบระบบหลัก

1. สร้าง `BMS_TEST`

```powershell
sqlcmd -S YOUR_SERVER -E -i D:\CIE\BMS\database\00_create_database_test.sql
```

2. สร้าง schema และ seed ลง `BMS_TEST`

ให้ generate ชุด test scripts ก่อน:

```powershell
python D:\CIE\BMS\database\generate_test_db_scripts.py
```

แล้วรันไฟล์ในโฟลเดอร์ `database\test`

3. สตาร์ต mock SAP

```powershell
python D:\CIE\BMS\test_support\mock_sap_server.py --host 127.0.0.1 --port 18080
```

4. ใช้ config test

- copy [Web.Test.config.example](/D:/CIE/BMS/Web.Test.config.example) ไปเป็นไฟล์ config สำหรับรันทดสอบ
- ตั้ง `BMSConnectionString` ให้ชี้ `BMS_TEST`
- ตั้ง `SAPAPI_BASEURL` เป็น `http://127.0.0.1:18080`

5. รัน smoke test

- Upload Draft OTB
- Approve Draft OTB
- Save Extra / Save Switching
- Sync / Match PO
- Export Draft PO / Actual PO / Summary

## พฤติกรรมของ mock SAP

- `POST /ZPaymentPlan/OTBPlanUpload`
  - echo payload กลับพร้อม `messageType = S`
- `POST /ZPaymentPlan/OTBPlanSwitch`
  - echo payload กลับพร้อม `messageType = S`
- `GET /sap/opu/odata/SAP/ZBBIK_API_2_SRV/PoSet`
  - คืน OData response จาก `po_sample.json`
  - รองรับ filter แบบ `Po eq '...'` เบื้องต้น

## ขอบเขตที่ทดสอบได้ปลอดภัย

- flow ภายในระบบ BMS
- schema / seed / stored procedure
- validation logic
- matching logic
- การจัดการ response จาก SAP

## สิ่งที่ยังไม่ใช่การทดสอบ production จริง

- SAP authorization จริง
- network / certificate / firewall จริง
- response edge cases จาก SAP ที่ยังไม่ได้ใส่เพิ่มใน mock data

## แนะนำเพิ่มในรอบถัดไป

- เพิ่ม fixture หลายชุด เช่น success / partial fail / hard fail
- เพิ่ม script generate SQL test copies สำหรับ `BMS_TEST` อัตโนมัติ
- เพิ่ม integration test runner ที่ยิง handlers ตาม action สำคัญ
