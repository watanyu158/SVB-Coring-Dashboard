# Coring Progress Dashboard

Dashboard สำหรับติดตามความคืบหน้างาน Cabling & Coring — สนามบินสุวรรณภูมิ (SVB)

## Deploy on Render
- Build: `cd backend && npm install`
- Start: `node backend/server.js`
- Port: `3002` (or `$PORT`)

## SharePoint (optional)
Set environment variable: `SHAREPOINT_URL=<your-link>`

## Static fallback
ถ้าไม่มี SharePoint URL หรือ Excel → ใช้ข้อมูล hardcode จาก SVB-Coring sheet
