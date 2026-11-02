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

เมื่ออัปเดตข้อมูลรายวัน ให้คัดลอก `data_private/t5_drop_cust.json` ที่สร้างจาก `Top drop Customer by class1.xlsx` ไปไว้ใน `publish/t5_drop_cust.json` หลังสร้างเว็บ แล้วอัปโหลดไฟล์นี้พร้อมข้อมูลอื่นด้วย สคริปต์สร้างเว็บเดิมไม่ได้คัดลอกโฟลเดอร์ `data_private` โดยอัตโนมัติ
