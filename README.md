# NCR Dashboard – Y2026

Dashboard hiển thị dữ liệu chất lượng NCR, đọc trực tiếp từ file Excel trên Google Drive.

## Cấu trúc project

```
ncr-dashboard/
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
└── src/
    ├── main.jsx
    └── App.jsx        ← toàn bộ dashboard ở đây
```

## Cài đặt local

```bash
npm install
npm run dev
```

Mở http://localhost:5173

## Deploy lên Vercel

1. Push code lên GitHub
2. Vào https://vercel.com → New Project → Import repo
3. Vercel tự detect Vite → bấm Deploy

## Cập nhật dữ liệu

- Mở file Excel trên Google Drive → chỉnh sửa → Save
- Dashboard tự động đọc dữ liệu mới khi refresh trang

## Yêu cầu

File Excel trên Google Drive phải được set **"Anyone with the link can view"**

## Tech stack

- React 18 + Vite
- Recharts (biểu đồ)
- SheetJS/xlsx (đọc Excel)
- corsproxy.io (bypass CORS khi fetch từ Google Drive)
