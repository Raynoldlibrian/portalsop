# Portal SOP — Kabupaten Indragiri Hulu

Registri Standar Operasional Prosedur (SOP) untuk seluruh OPD di lingkungan
Kabupaten Indragiri Hulu. Dibangun dengan React + Vite + Tailwind CSS,
backend-nya Google Sheets + Google Apps Script.

## Menjalankan di lokal

```bash
npm install
npm run dev
```

Buka `http://localhost:5173`.

## Build untuk production

```bash
npm run build
```

Hasilnya ada di folder `dist/`.

## Konfigurasi backend

URL API (Google Apps Script Web App) di-set di `src/App.jsx`, pada konstanta
`API_BASE_URL`. Kalau backend-nya di-deploy ulang (misalnya bikin "New
deployment" baru di Apps Script), URL-nya berubah dan perlu diupdate di sini.

## Deploy ke Vercel

1. Push repo ini ke GitHub
2. Buka [vercel.com](https://vercel.com) → **New Project** → import repo GitHub ini
3. Vercel otomatis mendeteksi ini project Vite — biarkan pengaturan default
   (Build Command: `npm run build`, Output Directory: `dist`)
4. Klik **Deploy**

## Struktur project

```
portal-sop/
├── index.html          # entry HTML, judul & meta tag
├── src/
│   ├── main.jsx         # React entry point
│   ├── App.jsx           # seluruh komponen Portal SOP
│   └── index.css         # Tailwind directives
├── tailwind.config.js
├── postcss.config.js
├── vite.config.js
└── package.json
```
