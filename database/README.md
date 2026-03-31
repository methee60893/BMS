# BMS Database Scripts

ลำดับการรันแนะนำ

1. `00_create_database.sql`
2. `01_create_tables.sql`
3. `02_create_views.sql`
4. `03_create_stored_procedures.sql`
5. `04_seed_master_data_template.sql`
6. `05_post_deploy_verify.sql`

หมายเหตุสำคัญ

- สคริปต์ชุดนี้ออกแบบจาก source code ในโปรเจกต์ปัจจุบัน เพื่อให้ object หลักที่ระบบเรียกใช้งานมีครบ
- หลังสร้าง table แล้ว ต้อง seed ข้อมูล master อย่างน้อยใน `MS_Month`, `MS_Year`, `MS_Company`, `MS_Category`, `MS_Segment`, `MS_Brand`, `MS_Vendor`, `MS_CCY`, `MS_Version`
- มี template พร้อมแก้ไขได้ใน `04_seed_master_data_template.sql`
- มี post-check หลัง deploy ใน `05_post_deploy_verify.sql`
- ถ้าต้อง regenerate seed และ verify sql จาก Excel ล่าสุด ให้ใช้ `generate_seed_from_excel.py`
- ถ้าต้อง compare จำนวนแถว Excel กับ DB โดยตรง ให้ใช้ `compare_excel_to_db.py`
- Stored procedures ถูกทำให้ idempotent และเน้นความปลอดภัยของข้อมูลมากกว่าการ match แบบ aggressive
- ถ้าจะขึ้น production ควรเพิ่ม user/role/security policy แยกต่างหาก
- ดูขั้นตอน deploy แบบพร้อมใช้ได้ใน `DEPLOYMENT_CHECKLIST.md`
