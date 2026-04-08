# ระบบทะเบียนคุมลูกหนี้เงินบำรุง

โปรเจกต์นี้เป็นชุดไฟล์ Google Apps Script สำหรับใช้กับ Google Sheet ต้นฉบับที่มีชีต:

- `ใบแจ้งอัตราค่าบริการ`
- `ใบแจ้งค่าบริการ`
- `Test No.`
- `ทะเบียนลูกหนี้`

สคริปต์จะช่วยทำงานดังนี้:

- รับข้อมูลจากหน้า Web App แทนชีตกรอกข้อมูลเดิม
- สร้างชีตระบบแบบซ่อนชื่อ `ข้อมูลฟอร์มระบบ` เพื่อเป็นตัวกลางให้แบบพิมพ์ทำงาน
- รัน `เลขที่ใบแจ้งอัตราค่าบริการ` อัตโนมัติในรูปแบบ `001/69`
- บันทึกประวัติลง `ทะเบียนลูกหนี้`
- อัปเดตสถานะ `รอชำระ / ชำระแล้ว / ยกเลิกใบแจ้งหนี้`
- พิมพ์ PDF จากชีต `ใบแจ้งอัตราค่าบริการ` และ `ใบแจ้งค่าบริการ`
- เก็บลิงก์ PDF กลับไปที่ข้อมูลลูกหนี้
- โหลดรายการเก่ากลับมาพิมพ์ซ้ำได้
- แสดงข้อมูลสรุปต่อจากคอลัมน์ H ในชีต `ทะเบียนลูกหนี้`
- หลังบันทึกจากฟอร์ม ระบบจะพาไปหน้า `DocumentDetail.html` เพื่อดูรายละเอียด แก้ไข และพิมพ์เอกสาร

## วิธีติดตั้ง

1. เปิด Google Sheet ของคุณ
2. ไปที่ `Extensions > Apps Script`
3. สร้างไฟล์ตามนี้แล้ววางโค้ดจากโฟลเดอร์นี้:
   - `appsscript.json`
   - `Code.gs`
   - `Dashboard.html`
   - `InvoiceForm.html`
   - `DocumentDetail.html`
   - `config.js`
   - `site.css`
4. บันทึกโปรเจกต์
5. รันฟังก์ชัน `setupSystem` 1 ครั้งเพื่ออนุญาตสิทธิ์และตั้งค่าระบบ
6. กลับมาที่ Google Sheet แล้วรีเฟรชหน้า
7. จะเห็นเมนูใหม่ชื่อ `ระบบลูกหนี้`
8. หลังจากรัน `setupSystem` แล้ว จึงค่อยลบหรือเลิกใช้งานชีต `กรุณากรอกข้อมูล`

## ขั้นตอนย้ายระบบจากชีตเดิม

1. วางโค้ดเวอร์ชันใหม่ให้ครบ
2. รัน `setupSystem`
3. ระบบจะสร้างชีตซ่อน `ข้อมูลฟอร์มระบบ`
4. ระบบจะปรับสูตรของชีต `ใบแจ้งอัตราค่าบริการ` และ `ใบแจ้งค่าบริการ` ให้ไปอ้างอิงชีตซ่อนแทน
5. หลังจากนั้นคุณสามารถไม่ให้ผู้ใช้เข้าชีต `กรุณากรอกข้อมูล` หรือจะลบชีตนั้นออกก็ได้

## การเปิดผ่านลิงก์ Web App

ถ้าคุณ `Deploy > New deployment > Web app` แล้วเปิดผ่านลิงก์ของ Apps Script:

- ระบบจะเปิดหน้า Web App ที่มีทั้งฟอร์มกรอกข้อมูลและ Dashboard
- หน้า `Dashboard` ใช้ดูภาพรวมและกดปุ่ม `เพิ่มใบแจ้งหนี้`
- เมื่อกด `เพิ่มใบแจ้งหนี้` ระบบจะเปิดหน้า `InvoiceForm.html`
- หน้า `InvoiceForm` ใช้สำหรับกรอกข้อมูลและกด `บันทึกข้อมูล` เพียงปุ่มเดียว
- หลังบันทึก ระบบจะสร้าง PDF ทั้ง 2 ใบอัตโนมัติ แล้วพาไปหน้า `DocumentDetail.html`
- หน้า `DocumentDetail` ใช้ดูรายละเอียดเอกสาร, กดแก้ไขข้อมูล, และกดพิมพ์ `ใบแจ้งอัตราค่าบริการ` กับ `ใบแจ้งค่าบริการ`
- ชีต `ทะเบียนลูกหนี้` ยังคงใช้สำหรับเก็บประวัติและอัปเดตวันที่ชำระหนี้

## การใช้งานแบบ API สำหรับเว็บที่โฮสต์บน GitHub

ระบบรองรับ route API เพิ่มแล้ว โดยยังใช้ Apps Script ตัวเดิมเป็น backend ได้เลย

- Base URL ที่ใส่ใน `config.js`: `https://script.google.com/macros/s/.../exec`
- ระบบจะเติม `mode=api` ให้อัตโนมัติ
- ตรวจสอบระบบ: `action=health`
- ดึงข้อมูล dashboard: `action=dashboard`
- ดึงข้อมูลฟอร์ม: `action=formBootstrap`
- ดึงรายละเอียดเอกสาร: `action=detail&recordId=...`
- ดึง URL PDF ที่เก็บไว้แล้ว: `action=documentUrl&recordId=...&type=rateStatement`
- ขอเลขที่ใหม่สำหรับเริ่มฟอร์ม: `action=newInvoice`
- บันทึกข้อมูล: `POST action=save`
- ยกเลิกใบแจ้งหนี้: `POST action=cancel`

ตัวอย่าง `GET` จากหน้าเว็บ static:

```js
const API_BASE = 'https://script.google.com/macros/s/DEPLOYMENT_ID/exec?mode=api';
const response = await fetch(`${API_BASE}&action=dashboard`);
const result = await response.json();
console.log(result.data);
```

ตัวอย่าง `POST` บันทึกข้อมูล:

```js
const API_BASE = 'https://script.google.com/macros/s/DEPLOYMENT_ID/exec?mode=api&action=save';

const response = await fetch(API_BASE, {
  method: 'POST',
  headers: {
    'Content-Type': 'text/plain;charset=utf-8'
  },
  body: JSON.stringify({
    payload: formPayload
  })
});

const result = await response.json();
console.log(result.data);
```

หมายเหตุ:

- ใช้ `text/plain` ในตัวอย่าง `POST` เพื่อให้เรียกจากเว็บ static ได้ง่ายและลดปัญหา preflight ของ browser
- API แบบ `GET` รองรับ `callback=ชื่อฟังก์ชัน` สำหรับการเรียกแบบ JSONP ได้ด้วย ถ้าหน้าเว็บ static ต้องการทางเลือกสำรอง
- ถ้าจะโฮสต์หน้าเว็บบน GitHub Pages แล้วเปิดใช้ API แบบ public จริง ควรพิจารณาเรื่องสิทธิ์การเข้าถึงเพิ่มเติม เพราะ URL ของ Apps Script API จะถูกเรียกจากหน้าเว็บฝั่ง client โดยตรง
- ถ้าต้องการความปลอดภัยสูงกว่านี้ในอนาคต แนะนำแยก backend ไปที่ Cloud Run, Firebase Functions หรือระบบที่รองรับ auth/API key ได้ชัดเจนกว่า

## การเปิดหน้า HTML บน GitHub Pages

ตอนนี้ไฟล์ [Dashboard.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/Dashboard.html), [InvoiceForm.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/InvoiceForm.html), และ [DocumentDetail.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/DocumentDetail.html) ถูกปรับให้ใช้งานได้ 2 แบบแล้ว:

- ถ้าเปิดจาก Apps Script Web App จะใช้ `google.script.run`
- ถ้าเปิดจาก GitHub Pages จะสลับไปเรียก `Apps Script API` อัตโนมัติ
- หน้า static จะอ่านค่า API จาก [config.js](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/config.js) เป็นหลัก
- หน้าเว็บทั้งหมดใช้สไตล์กลางจาก [site.css](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/site.css)

วิธีใช้งานบน GitHub Pages:

1. Deploy Apps Script เป็น Web App เวอร์ชันล่าสุด
2. เปิดใช้หน้าเว็บจาก GitHub Pages ผ่าน [Dashboard.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/Dashboard.html)
3. เปิดไฟล์ [config.js](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/config.js)
4. เปลี่ยนค่า `apiBase` ให้เป็นลิงก์จริงรูปแบบ `https://script.google.com/macros/s/.../exec`
5. ระบบจะอ่านค่า `apiBase` จาก `config.js` อัตโนมัติในทุกหน้า
6. ถ้าไฟล์ `config.js` โหลดไม่สำเร็จจริง ระบบค่อย fallback ไปถามผู้ใช้และจำค่าไว้ใน browser

ถ้าต้องการส่ง URL เข้าไปเลยตั้งแต่ลิงก์แรก สามารถเปิดแบบนี้ได้:

```text
https://<github-pages-url>/Dashboard.html?apiBase=https://script.google.com/macros/s/.../exec
```

หมายเหตุ:

- หน้า `Dashboard` และ `DocumentDetail` รองรับการอ่านข้อมูลผ่านทั้ง `fetch` และ fallback แบบ `JSONP`
- การกด `บันทึกข้อมูล` จากหน้า `InvoiceForm` บน GitHub Pages จะลอง `POST` ไปยัง Apps Script API ก่อน
- ถ้า browser บล็อก `POST` ข้ามโดเมน ระบบจะ fallback ไปส่งฟอร์มตรงเข้า Apps Script แล้วเด้งกลับมาหน้า `DocumentDetail` ให้อัตโนมัติ
- ถ้า deployment ของ Apps Script ยังบล็อกการเรียกข้ามโดเมนมากกว่านี้อีก อาจต้องทำ proxy หรือ bridge เพิ่มในรอบถัดไป

## วิธีใช้งานหลัก

1. เปิดลิงก์ Web App
2. กดปุ่ม `เพิ่มใบแจ้งหนี้`
3. กรอกชื่อผู้ใช้บริการ, ที่อยู่, รายการใบแจ้งอัตราค่าบริการ, และรายการใบแจ้งค่าบริการ
4. กด `บันทึกข้อมูล`
5. ระบบจะพาไปหน้ารายละเอียดเอกสารโดยอัตโนมัติ
6. ถ้าต้องการแก้ไขหรือพิมพ์ซ้ำ ให้เปิดจากหน้า `DocumentDetail` หรือจากรายการใน Dashboard
7. เมื่อมีการกรอกวันที่ชำระหนี้ในคอลัมน์ G ของชีต `ทะเบียนลูกหนี้`
   ระบบจะอัปเดตสถานะเป็น `ชำระแล้ว` ให้อัตโนมัติ
8. หากต้องการยกเลิกใบแจ้งหนี้ สามารถใช้เมนูใน Google Sheet ได้ตามเดิม

## หมายเหตุสำคัญ

- สคริปต์จะแสดงข้อมูลสรุปต่อจากคอลัมน์ H ในชีต `ทะเบียนลูกหนี้`
  และซ่อนเฉพาะคอลัมน์ข้อมูลรายการแบบ JSON ของระบบ
- สคริปต์จะสร้างชีตซ่อน `ข้อมูลฟอร์มระบบ` เพื่อใช้เป็น buffer สำหรับแบบพิมพ์
- ถ้าจะพิมพ์ PDF ให้ภาพในแม่แบบปกติ ควรใช้รูปที่พื้นขาวจริงใน Google Sheet
  โดยคุณสามารถใช้ไฟล์ในโฟลเดอร์ [print-assets](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/print-assets) ไปแทนรูปเดิมในชีตแม่แบบได้
- ระบบใช้ lock ฝั่ง Apps Script ตอนบันทึกเลขที่ใบแจ้งและตอนสร้าง PDF
  เพื่อช่วยลดปัญหาหลายเครื่องใช้งานพร้อมกันแล้วชนกันที่เลขที่เอกสารหรือแม่แบบพิมพ์
- ถ้ามีผู้ใช้เปิดรายการเดียวกันไว้คนละเครื่อง แล้วอีกเครื่องบันทึกไปก่อน
  ระบบจะกันการบันทึกทับด้วย token เวลาแก้ไขล่าสุด และให้เปิดข้อมูลใหม่อีกครั้งก่อนบันทึก
- วันที่บนแบบพิมพ์จะถูกส่งเป็นข้อความไทยเต็ม เช่น `7 เมษายน พ.ศ. 2569`
- จำนวนเงินหลักในคอลัมน์ F และยอดลูกหนี้ของระบบจะอิงจากยอดรวมของ `ใบแจ้งค่าบริการ` เท่านั้น
- ยอดของ `ใบแจ้งอัตราค่าบริการ` ใช้เพื่อประกอบเอกสารฝั่งอัตราค่าบริการ และจะไม่ถูกนำไปรวมเป็นยอดลูกหนี้
- ระบบนี้ตั้งใจ “ไม่ลบใบแจ้งหนี้” แต่ให้ใช้สถานะ “ยกเลิกใบแจ้งหนี้” แทน

## โครงสร้างไฟล์

- [appsscript.json](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/appsscript.json)
- [Code.gs](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/Code.gs)
- [Dashboard.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/Dashboard.html)
- [InvoiceForm.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/InvoiceForm.html)
- [DocumentDetail.html](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/DocumentDetail.html)
- [config.js](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/config.js)
- [site.css](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/site.css)
- [print-assets](/Users/9phoomphi/Desktop/ระบบทะเบียนคุมลูกหนี้เงินบำรุง/print-assets)
