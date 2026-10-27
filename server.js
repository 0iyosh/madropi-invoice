/**
 * MADROPI - Server Lokal dengan Integrasi Template Excel Asli
 * Menggunakan modul bawaan Node.js dan ExcelJS untuk replikasi 100% presisi
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const ExcelJS = require('exceljs');

const PORT = process.env.PORT || 3000;
const TEMPLATE_FILE = path.join(__dirname, 'template.xlsx');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
};

// Bulan dalam Bahasa Indonesia
const NAMA_BULAN = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

function formatTanggalIndo(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return `${d.getDate()} ${NAMA_BULAN[d.getMonth()]} ${d.getFullYear()}`;
}

const server = http.createServer(async (req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const parsedUrl = req.url.split('?')[0];

  // API ENDPOINT: POST /api/generate-xlsx
  if (req.method === 'POST' && parsedUrl === '/api/generate-xlsx') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', async () => {
      try {
        const payload = JSON.parse(body);
        const buffer = await generateExcelFromTemplate(payload);

        const safeProject = (payload.lokasiProyek || 'Makadam').replace(/[^a-zA-Z0-9_-]/g, '_');
        const filename = `Invoice_${safeProject}_${payload.tanggalInvoice || 'tanggal'}.xlsx`;

        res.writeHead(200, {
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          'Content-Disposition': `attachment; filename="${filename}"`,
          'Content-Length': buffer.length
        });
        res.end(buffer);
      } catch (err) {
        console.error('Error generating Excel:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: err.message || 'Gagal menghasilkan file Excel' }));
      }
    });
    return;
  }

  // Melayani file statis
  let filePath = path.join(__dirname, parsedUrl === '/' ? 'index.html' : parsedUrl);

  // Mencegah directory traversal
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403);
    res.end('Akses Ditolak');
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('File tidak ditemukan');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  });
});

/**
 * Mengisi template.xlsx dengan data invoice dari pengguna
 * Mempertahankan 100% style, merge cell, font, warna, border, dan formula asli!
 */
async function generateExcelFromTemplate(data) {
  if (!fs.existsSync(TEMPLATE_FILE)) {
    throw new Error('File master template.xlsx tidak ditemukan di direktori projek!');
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TEMPLATE_FILE);
  const ws = wb.getWorksheet(1);

  // 1. Lokasi Proyek di sel A2
  const lokasiProyek = data.lokasiProyek || 'CONDO VILLAS SUMMARECON';
  ws.getCell('A2').value = `Lokasi Proyek : ${lokasiProyek}`;

  // 2. Baris Data (Maksimal slot standar adalah baris 6 sampai baris 14 = 9 baris)
  const items = data.items || [];
  let grandTotalCalculated = 0;

  for (let r = 6; r <= 14; r++) {
    const itemIndex = r - 6;
    const item = items[itemIndex];

    if (item) {
      const vol = Number(item.volume) || ((item.panjang * item.lebar * item.tinggi) / 1000000);
      const volRounded = Math.round(vol * 1000) / 1000;
      const total = Number(item.total) || Math.round(volRounded * 200000);
      grandTotalCalculated += total;

      ws.getCell(`A${r}`).value = itemIndex + 1;
      ws.getCell(`B${r}`).value = formatTanggalIndo(item.tanggal);
      ws.getCell(`C${r}`).value = item.plat || '';
      ws.getCell(`D${r}`).value = Number(item.panjang) || null;
      ws.getCell(`E${r}`).value = Number(item.lebar) || null;
      ws.getCell(`F${r}`).value = Number(item.tinggi) || null;

      // Formula volume
      ws.getCell(`G${r}`).value = {
        formula: `ROUND(D${r}*E${r}*F${r}/1000000,3)`,
        result: volRounded
      };

      // Harga pasti 200.000
      ws.getCell(`H${r}`).value = 200000;

      // Formula total
      ws.getCell(`I${r}`).value = {
        formula: `G${r}*H${r}`,
        result: total
      };

      ws.getCell(`J${r}`).value = item.keterangan || '';
    } else {
      // Baris kosong bergaris (dipertahankan borders dan stylenya tapi nilainya null)
      ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'].forEach(col => {
        ws.getCell(`${col}${r}`).value = null;
      });
    }
  }

  // 3. Grand Total di sel I15
  ws.getCell('A15').value = 'GRAND TOTAL';
  ws.getCell('I15').value = {
    formula: 'SUM(I6:I14)',
    result: grandTotalCalculated
  };

  // 4. Terbilang di sel A17
  const terbilangText = data.terbilang || '';
  ws.getCell('A17').value = `Terbilang :  ${terbilangText}`;

  // 5. Tanggal Invoice & Kota di sel G24 (merged G24:J24)
  const kota = data.kotaInvoice || 'Tangerang';
  const tglFormatted = formatTanggalIndo(data.tanggalInvoice);
  ws.getCell('G24').value = `${kota}, ${tglFormatted}`;

  // 6. Penandatangan di sel G33 (merged G33:J33)
  const namaTtd = data.namaPenandatangan || 'MADROPI';
  ws.getCell('G33').value = namaTtd;

  // Catatan rekening di A19, A20, A21 dibiarkan utuh sesuai template asli

  // Tulis workbook ke buffer
  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// Menjalankan Server
server.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log('========================================================');
  console.log(` MADROPI - Sistem Rekapitulasi & Invoice Material`);
  console.log(` Terintegrasi langsung dengan template Excel: template.xlsx`);
  console.log(` Server berjalan di: ${url}`);
  console.log(' Tekan Ctrl+C untuk menghentikan server');
  console.log('========================================================');

  // Buka browser jika dijalankan secara lokal
  exec(`start ${url}`);
});
