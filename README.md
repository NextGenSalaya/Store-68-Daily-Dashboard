# Store 68 Daily Dashboard

Dashboard ประจำวันของ Makro สาขา 68 Salaya (District 5) แบบ BIC: Menu ด้านซ้าย, Filter Date Period / Division, Card, กราฟ และตาราง

**เปิดดู:** https://nextgensalaya.github.io/Store-68-Daily-Dashboard/

## หัวข้อใน Menu

| หัวข้อ | หน้า |
|---|---|
| ยอดขาย · Sales | ยอดขาย & ลูกค้า, Category & Class, สินค้า & Top Drop Customer, Fresh Food Sale & Top 50, Hamper & Soft-serve |
| Stock & Order | Stock & Stock Days, Stock ใกล้หมด & ติดลบ, Over Stock & On Order, Aging NBS & Clearance, Top Items & Item List |
| Service Level · KVI | Service level, KVI · OB · Makro Mail Monitoring, KVI Price Survey List, Pack & Component |
| Shrinkage | Shrinkage ภาพรวม & แผนก, Shrinkage by Code |
| O2O | O2O ภาพรวม & Orders, Focus Top 1000 O2O, Item Shortage & OOS |

## อัปเดตข้อมูลทุกวัน

1. วาง File Excel ของวันนี้ไว้ในโฟลเดอร์นี้ (อยู่ในโฟลเดอร์ย่อยก็ได้)
2. ดับเบิลคลิก `update.bat`
3. รอประมาณ 1 นาที แล้วกด Refresh หน้าเว็บ

`update.bat` จะอ่าน Excel ด้วย `build_*.py` → สร้างเว็บใน `docs/` ด้วย `make_site.py` → Push ขึ้น GitHub

## ไฟล์ใน Repository

| ไฟล์ | หน้าที่ |
|---|---|
| `board/index.html` | หน้า Dashboard (ต้นฉบับ) |
| `build_all_dashboard.py`, `build_topic2.py` … `build_topic5.py` | อ่าน Excel แต่ละเรื่องแล้วสร้างข้อมูล JSON |
| `make_site.py` | รวมหน้าและข้อมูลเป็นเว็บใน `docs/` |
| `docs/` | เว็บที่ GitHub Pages เปิดให้ดู (`docs/data/` คือข้อมูล) |

File Excel ต้นฉบับไม่ถูกอัปโหลด (`.gitignore` กันไว้) ชื่อลูกค้า เลขสมาชิก ชื่อพนักงาน และอีเมล ไม่อยู่ในข้อมูล
