# Store 68 Daily Dashboard

Dashboard ประจำวันของ Makro สาขา 68 Salaya (District 5) แบบ BIC: Menu ด้านซ้าย, Filter Date Period / Division, Card, กราฟ และตาราง

**เปิดดู:** https://nextgensalaya.github.io/Store-68-Daily-Dashboard/

## หัวข้อใน Menu

| หัวข้อ | หน้า |
|---|---|
| Overall Trend | ยอดขาย & ลูกค้า, Category, Class & สินค้า, Sale by Item & Top Drop Customer, Fresh Food Sale & Top 50, Hamper & Soft-serve |
| Stock & Order | Stock & Stock Days, Stock ใกล้หมด & ติดลบ, Over Stock & On Order, Aging NBS & Clearance, Top Items & Item List |
| Service Level · KVI | Service level, KVI · OB · Makro Mail Monitoring, KVI Price Survey List, Pack & Component |
| Shrinkage | Shrinkage ภาพรวม & แผนก, Shrinkage by Code |
| O2O | O2O ภาพรวม & Orders, Focus Top 1000 O2O, Item Shortage & OOS |

## อัปเดตข้อมูลทุกวัน

### ผ่านหน้าเว็บ

กด **ใส่ไฟล์ Excel / อัปเดตข้อมูล** ใต้วันที่บน Dashboard หรือเปิด [หน้านำเข้าไฟล์](https://nextgensalaya.github.io/Store-68-Daily-Dashboard/upload/)

1. เลือกทุกกลุ่มหรือเฉพาะกลุ่มที่ต้องการอัปเดต แล้วเลือกไฟล์ Excel หลายไฟล์หรือทั้งโฟลเดอร์ ระบบจะแจ้งไฟล์ที่ขาดหรือซ้ำ
2. กด **ตรวจสอบและแปลงข้อมูล** ไฟล์ Excel ประมวลผลในเบราว์เซอร์ด้วย Pyodide และตัวแปลง Python เดิม ไม่ส่งไฟล์ Excel ต้นฉบับขึ้น GitHub
3. ตรวจสอบจำนวนชุดข้อมูล รายการ และวันที่ในรายงาน
4. ใส่ Fine-grained GitHub personal access token ที่อนุญาตเฉพาะ repository นี้ และมีสิทธิ์ **Contents: Read and write** รหัสนี้อยู่ในหน่วยความจำเฉพาะตอนส่งข้อมูล ไม่บันทึกไว้ใน localStorage/sessionStorage หรือซอร์ส
5. ยืนยันการเผยแพร่ แล้วกด **เผยแพร่ข้อมูลใหม่** ระบบส่งเฉพาะ JSON ที่เปลี่ยนเป็น commit เดียว โดยไม่ทับไฟล์หน้าเว็บหรือข้อมูลกลุ่มอื่น และรอให้เว็บยืนยัน `data-release.json` เวอร์ชันใหม่ก่อนแสดงว่าสำเร็จ

รองรับ 31 ไฟล์ต้นฉบับตามชื่อ/โครงสร้างรายงานเดิม รวม `.xlsx`, `.xlsm`, `.xls`, `.xlsb` ต้องเลือกไฟล์ครบสำหรับแต่ละกลุ่ม ไม่ประมวลผลบางส่วนเมื่อไฟล์ขาด ใช้ค่าผลลัพธ์สูตรที่ Excel บันทึกไว้ จึงควรเปิดคำนวณและบันทึกไฟล์ใน Excel ก่อนนำเข้า แนะนำใช้คอมพิวเตอร์และเปิดแท็บไว้ระหว่างประมวลผล

การเผยแพร่รวมข้อมูลลูกค้ารายคนตามสิทธิ์ที่เจ้าของยืนยันไว้ ผู้เผยแพร่ต้องตรวจสอบข้อมูลและยืนยันในหน้าจอทุกครั้ง ไม่มีการอัปโหลดอัตโนมัติเมื่อเลือกไฟล์ หาก token ไม่มีสิทธิ์หรือหมดอายุ GitHub จะปฏิเสธการเขียน

ตรวจสอบตัวนำเข้าและการเผยแพร่จำลองด้วย `node --test tests/importer.test.mjs` การทดสอบจำลองไม่แทนการทดสอบรหัสสิทธิ์จริงของผู้ใช้

### ผ่านสคริปต์เดิม

1. วาง File Excel ของวันนี้ไว้ในโฟลเดอร์งาน (อยู่ในโฟลเดอร์ย่อยก็ได้)
2. ดับเบิลคลิก `update.bat` จะเปิดโฟลเดอร์ `publish` และหน้า Upload files ของ GitHub ให้
3. ใน `publish` กด Ctrl+A ลากทุกไฟล์ (index.html และ .json เท่านั้น) ไปวางในหน้า GitHub
4. รอให้ขึ้นครบ แล้วกด **Commit changes** ประมาณ 1 นาทีต่อมากด Refresh หน้าเว็บ

**ห้ามลากไฟล์ .xlsx / .xlsm / .xls / .xlsb ขึ้น GitHub** เพราะมีชื่อลูกค้า พนักงาน และอีเมล

## ไฟล์ใน Repository

| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` + `*.json` | เว็บที่ GitHub Pages เปิดให้ดู (หน้าเว็บ และข้อมูลแต่ละชุด) |
| `board/index.html` | หน้า Dashboard (ต้นฉบับ) |
| `build_all_dashboard.py`, `build_topic2.py` … `build_topic5.py` | อ่าน Excel แต่ละเรื่องแล้วสร้างข้อมูล JSON |
| `make_site.py` | รวมหน้าและข้อมูลเป็นเว็บในโฟลเดอร์ `publish` |
| `docs/` | เว็บรุ่นเดิม (ไม่ได้ใช้แล้ว) |

## ตารางลูกค้าที่ซื้อลดลงรายคน

หน้า `Sale by Item & Top Drop Customer` ใช้ `t5_drop_cust.json` ที่โฟลเดอร์หลักของเว็บไซต์เพื่อแสดงตารางรายคน พร้อมตัวกรอง Department / ประเภทลูกค้า / สถานะ และการค้นหา

เจ้าของอนุมัติให้เผยแพร่ชื่อ เลขสมาชิก และยอดซื้อรายลูกค้าจริงบนเว็บและ repository สาธารณะเมื่อ 2 พฤศจิกายน 2026 ไฟล์นี้จึงเปิดอ่านและดาวน์โหลดได้โดยไม่ต้องเข้าสู่ระบบ

`make_site.py` คัดลอก `data_private/t5_drop_cust.json` (สร้างจาก `Top drop Customer by class1.xlsx`) ไปไว้ใน `publish` ให้เอง และใส่ลิงก์ **ใส่ไฟล์ Excel / อัปเดตข้อมูล** ใน `index.html` ของเว็บ จึงลากทุกไฟล์ใน `publish` ขึ้นได้เลย
