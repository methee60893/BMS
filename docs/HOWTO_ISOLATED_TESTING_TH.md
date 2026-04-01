# How To: ทดสอบแบบไม่กระทบระบบหลัก

คู่มือนี้เป็นลำดับใช้งานจริง ตั้งแต่เตรียม environment, เริ่มทดสอบ, ระหว่างทดสอบ, และหลังทดสอบเสร็จ

## 1. เตรียมก่อนทดสอบ

สิ่งที่ต้องมี

- SQL Server สำหรับฐานทดสอบ
- `python`
- `sqlcmd`
- ไฟล์ Excel master ที่ใช้ seed
- source code ล่าสุดของโปรเจกต์

ไฟล์ที่เกี่ยวข้อง

- [run_full_isolated_test.ps1](/D:/CIE/BMS/run_full_isolated_test.ps1)
- [run_test_step_by_step.ps1](/D:/CIE/BMS/run_test_step_by_step.ps1)
- [mock_sap_server.py](/D:/CIE/BMS/test_support/mock_sap_server.py)
- [generate_test_db_scripts.py](/D:/CIE/BMS/database/generate_test_db_scripts.py)
- [compare_excel_to_db.py](/D:/CIE/BMS/database/compare_excel_to_db.py)

## 2. ตรวจความพร้อมก่อนเริ่ม

รัน preflight ก่อนทุกครั้ง

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -PreflightOnly -SqlServer YOUR_SERVER -TestDatabase BMS_TEST -DbUser sa -DbPassword YOUR_PASSWORD
```

สิ่งที่ควรดูในผลลัพธ์

- `PASS` สำหรับไฟล์สำคัญ
- `PASS` หรือ `WARN` สำหรับ `sqlcmd`
- `PASS` หรือ `WARN` สำหรับ `SQL connectivity`
- `PASS` หรือ `WARN` สำหรับ `Mock SAP port`

หมายเหตุ

- ถ้าเป็น `FAIL` ต้องแก้ก่อนเริ่ม
- ถ้า `Mock SAP port` เป็น `WARN` แปลว่ามี process ใช้ port นั้นอยู่แล้ว อาจเป็น mock ตัวเดิม

## 3. ดูขั้นตอนแบบไม่แตะอะไรจริง

ให้ dry run ก่อน

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -DryRun -RunSql -StartMockSap -SqlServer YOUR_SERVER -TestDatabase BMS_TEST -DbUser sa -DbPassword YOUR_PASSWORD
```

สิ่งที่ dry run จะทำ

- ตรวจสถานะ config ปัจจุบัน
- generate test SQL
- พิมพ์คำสั่ง `sqlcmd`
- จำลองการ switch config
- จำลองการ start mock SAP

## 4. เริ่มทดสอบจริง

### 4.1 รัน full isolated test

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -RunSql -StartMockSap -SqlServer YOUR_SERVER -TestDatabase BMS_TEST -DbUser sa -DbPassword YOUR_PASSWORD
```

สิ่งที่จะเกิดขึ้น

1. ตรวจ preflight
2. generate test SQL
3. สร้าง/อัปเดต schema และ seed ใน `BMS_TEST`
4. backup `Web.config`
5. สลับ `Web.config` ให้ชี้ test DB และ mock SAP
6. เปิด mock SAP ถ้า port ยังว่าง

### 4.2 เปิดเว็บแล้วทดสอบ flow

เปิดโปรเจกต์ผ่าน Visual Studio หรือ IIS Express ตามปกติ แล้วทดสอบ flow สำคัญ

- Upload Draft OTB
- Approve Draft OTB
- Save Extra
- Save Switching
- Sync PO
- Match / Manual Match
- Export report

## 5. ระหว่างทดสอบ

ตรวจ log ได้ที่

- [run_full_isolated_test.log](/D:/CIE/BMS/.codex-test-state/run_full_isolated_test.log)
- [run_test_step_by_step.log](/D:/CIE/BMS/.codex-test-state/run_test_step_by_step.log)

ถ้าต้องเช็กจำนวนข้อมูลหลัง seed

```powershell
python D:\CIE\BMS\database\compare_excel_to_db.py --server YOUR_SERVER --database BMS_TEST
```

## 6. หลังทดสอบเสร็จ

คืนค่า environment ปกติ

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_full_isolated_test.ps1 -RestoreOnly
```

สิ่งที่จะเกิดขึ้น

- `Web.config` ถูก restore จาก backup
- snapshot/state ชั่วคราวถูกล้าง
- environment กลับไปใช้ค่าปกติเดิม

## 7. เช็กล่าสุดหลัง restore

```powershell
powershell -ExecutionPolicy Bypass -File D:\CIE\BMS\run_test_step_by_step.ps1 status
```

ควรเห็นว่า

- `Backup exists: False`
- `Snapshot exists: False`
- `BMSConnectionString` กลับไปชี้ฐานเดิม
- `SAPAPI_BASEURL` กลับไปชี้ QAS/ค่าปกติเดิม

## 8. กรณีแนะนำการใช้งานจริง

ก่อนเริ่มรอบใหม่

1. รัน `-PreflightOnly`
2. รัน `-DryRun`
3. ค่อยรันจริง

หลังจบรอบทดสอบ

1. เก็บผล log
2. รัน compare ถ้าต้องตรวจจำนวนแถว
3. รัน `-RestoreOnly`
