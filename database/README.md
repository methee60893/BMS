# BMS Database Scripts

ลำดับการรันแนะนำ

1. `00_create_database.sql`
2. `01_create_tables.sql`
3. `02_create_views.sql`
4. `03_create_stored_procedures.sql`
5. `04_seed_master_data_template.sql`
6. `07_create_draft_otb_background_jobs.sql` (required before deploying Sprint 1 Draft OTB upload/approval progress)
7. `05_post_deploy_verify.sql`

หมายเหตุสำคัญ

- สคริปต์ชุดนี้ออกแบบจาก source code ในโปรเจกต์ปัจจุบัน เพื่อให้ object หลักที่ระบบเรียกใช้งานมีครบ
- หลังสร้าง table แล้ว ต้อง seed ข้อมูล master อย่างน้อยใน `MS_Month`, `MS_Year`, `MS_Company`, `MS_Category`, `MS_Segment`, `MS_Brand`, `MS_Vendor`, `MS_CCY`, `MS_Version`
- มี template พร้อมแก้ไขได้ใน `04_seed_master_data_template.sql`
- มี post-check หลัง deploy ใน `05_post_deploy_verify.sql`
- ถ้าต้อง regenerate seed และ verify sql จาก Excel ล่าสุด ให้ใช้ `generate_seed_from_excel.py`
- ถ้าต้อง compare จำนวนแถว Excel กับ DB โดยตรง ให้ใช้ `compare_excel_to_db.py`
- ชุด seed ปัจจุบันครอบคลุม user/role/menu/permission สำหรับระบบหลังบ้านด้วย
- Stored procedures ถูกทำให้ idempotent และเน้นความปลอดภัยของข้อมูลมากกว่าการ match แบบ aggressive
- ถ้าจะขึ้น production ควรเพิ่ม user/role/security policy แยกต่างหาก
- ดูขั้นตอน deploy แบบพร้อมใช้ได้ใน `DEPLOYMENT_CHECKLIST.md`

## Draft OTB background jobs (Sprint 1)

- ต้องรัน `07_create_draft_otb_background_jobs.sql` ก่อนนำหน้าจอ Draft OTB รุ่นนี้ขึ้นใช้งาน
- สำหรับฐาน PRD เดิมที่ `10.3.152.155/BMS` ให้ใช้ `08_deploy_sprint1_prd.sql` แทนการรันไฟล์ `07` โดยตรง สคริปต์ PRD มี target guard, data preflight, transaction และ post-deploy verification ครบในไฟล์เดียว
- ต้องรัน SQL ให้สำเร็จก่อน deploy application; โค้ด Sprint 1 อ้างถึงตาราง claim และ index ตามชื่อระหว่าง upload/approve/cancel/switch
- การ deploy SQL ไม่แก้ค่าข้อมูลธุรกิจเดิม แต่เพิ่ม operational tables/indexes และจะหยุดโดยไม่ deploy หากพบ Draft key ผิด/ซ้ำหรือมี job/claim ค้าง
- ตั้ง IIS request envelope เป็น 25 MB ตาม `Web.Test.config.example`; ตัวระบบยังจำกัดไฟล์ Draft OTB ที่ 20 MB และไม่เกิน 15,000 แถว
- Application Pool identity ต้องมีสิทธิ์ **Modify** (อ่าน/เขียน/สร้าง/ลบไฟล์) ที่ `App_Data\DraftOtbJobs` โดยไม่ต้องให้สิทธิ์กับโฟลเดอร์อื่น
- ข้อมูลสถานะอยู่ใน `dbo.Draft_OTB_Background_Job` ทำให้ปิดหรือ Refresh browser แล้วกลับมาดูงานเดิมได้ แต่ worker ทำงานใน IIS process ดังนั้น IIS/App Pool recycle หรือ deploy ระหว่างประมวลผลอาจหยุด worker
- ห้าม retry งานสถานะ `ReconciliationRequired` เพราะ SAP อาจรับข้อมูลแล้ว ให้ใช้ Job ID ตรวจสอบ SAP เทียบกับ BMS และบันทึกผลการ reconcile ก่อนทุกครั้ง
- งาน `ValidationFailed` ไม่มีผลต่อข้อมูลทั้งไฟล์และมีรายงานข้อผิดพลาดให้ดาวน์โหลด ผู้ใช้ต้องตรวจ/รับทราบก่อนเริ่มงานใหม่
- กำหนดงานดูแลเพื่อลบ job/payload/result และไฟล์ชั่วคราวที่หมดอายุตามนโยบายองค์กร แนะนำเก็บอย่างน้อย 30 วันสำหรับงานปกติ และเก็บ `ReconciliationRequired` จนปิดการตรวจสอบแล้ว ห้ามลบงานที่ยังไม่จบ
