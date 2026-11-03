/**
 * =====================================================================
 * BACKEND REST API JACNET (PT Citra Nusantara Cyberindo - Site Sumedang)
 * =====================================================================
 * Script ini bertindak sebagai REST API backend untuk Dashboard JACNET.
 * Database menggunakan Google Sheets dengan 2 Tab utama:
 *   1. "CLIENTS"  -> Data master profil & teknis pelanggan (47 kolom)
 *   2. "TAGIHAN"  -> Rekap tagihan & pembayaran bulanan (Januari - Desember)
 * 
 * CARA DEPLOY:
 * 1. Buka Google Sheet master Anda di browser.
 * 2. Klik menu "Ekstensi" (Extensions) > "Apps Script".
 * 3. Hapus seluruh kode bawaan di editor Apps Script.
 * 4. Salin (copy) seluruh kode di file ini, lalu tempel (paste).
 * 5. Klik icon disket Simpan (Save) atau tekan Ctrl + S.
 * 6. Klik tombol biru "Deploy" (Terapkan) di kanan atas > "New deployment".
 * 7. Klik icon roda gigi (Gear) di sebelah 'Select type' > pilih "Web app".
 * 8. Konfigurasi:
 *    - Description: v1.0 Backend JACNET
 *    - Execute as: Me (email Anda)
 *    - Who has access: Anyone (Wajib pilih Anyone agar dashboard bisa membaca/menulis)
 *    - 9. Klik "Deploy", izinkan review permissions, lalu salin "Web app URL" (akhiran /exec).
 * =====================================================================
 */

const SHEET_CLIENTS = 'CLIENTS';
const SHEET_TAGIHAN = 'TAGIHAN';

const MONTH_NAMES = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

// Nama cluster/wilayah per awalan ID Pelanggan — sesuai data di sheet TAGIHAN
// (PMT=Pematang, PRM=Perum, CDRS=Canarusa, dst.) dan label cluster di dashboard.
const CLUSTER_NAMES = {
  PMT: 'Pamatang', CKL: 'Cuklik (Sawahdadap)', CJK: 'Cijengkol',
  CCL: 'Cuklik Cimuncang', CDRS: 'Canarusa', CSPN: 'Cisurupan',
  CBT: 'Cibitung', CGLPK: 'Cigulampok', NRGL: 'Naringgul',
  BBKN: 'Babakan', PRM: 'Perum', PRMT: 'Sindang Pakuwon', JAC: 'Cimanggung'
};

function clusterNameForId(idPelanggan, fallback) {
  // Konvensi GSheet: seri ID PMT-Sxx adalah wilayah 'Pamatang Solo'
  if (/^PMT-S\d/i.test(String(idPelanggan || ''))) return 'Pamatang Solo';
  const m = String(idPelanggan || '').toUpperCase().match(/^([A-Z]+)/);
  if (m && CLUSTER_NAMES[m[1]]) return CLUSTER_NAMES[m[1]];
  return fallback || 'Cluster';
}

/**
 * Helper untuk response JSON dengan MIME type JSON yang mendukung CORS
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle HTTP GET Requests
 * Contoh endpoint:
 *   GET ?action=getData   -> Memuat data master klien, billing, & KPI
 *   GET ?action=ping      -> Cek koneksi server
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getData';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'ping') {
      return createJsonResponse({
        status: 'success',
        message: 'Backend Google Apps Script JACNET aktif!',
        timestamp: new Date().toISOString()
      });
    }

    if (action === 'getData' || action === 'getClients' || action === 'getBillings') {
      const clientsSheet = ss.getSheetByName(SHEET_CLIENTS);
      const tagihanSheet = ss.getSheetByName(SHEET_TAGIHAN);

      if (!clientsSheet || !tagihanSheet) {
        return createJsonResponse({
          status: 'error',
          message: 'Sheet "CLIENTS" atau "TAGIHAN" tidak ditemukan. Pastikan nama tab sheet sudah sesuai.'
        });
      }

      // 1. Baca data Master Klien
      const clientData = clientsSheet.getDataRange().getValues();
      const customerRows = [];
      if (clientData.length > 1) {
        for (let i = 1; i < clientData.length; i++) {
          const row = clientData[i];
          if (!row || (!row[6] && !row[8])) continue; // Lewati jika ID dan Nama kosong
          
          // Pastikan panjang array minimal 47 kolom dan format string aman
          const cleanRow = [];
          for (let c = 0; c < 47; c++) {
            let val = row[c] !== undefined && row[c] !== null ? row[c] : '';
            if (val instanceof Date) {
              val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
            cleanRow.push(String(val));
          }
          cleanRow[0] = String(customerRows.length + 1); // Penomoran urut otomatis
          customerRows.push(cleanRow);
        }
      }

      // 2. Baca data Rekap Tagihan Bulanan
      const tagihanData = tagihanSheet.getDataRange().getValues();
      const billingRows = [];
      const currentYear = new Date().getFullYear();

      if (tagihanData.length > 1) {
        for (let i = 1; i < tagihanData.length; i++) {
          const row = tagihanData[i];
          if (!row || (!row[1] && !row[2])) continue;

          const no = String(billingRows.length + 1);
          const idPel = String(row[1] || '').trim();
          const nama = String(row[2] || '').trim();
          const cluster = String(row[3] || '').trim();
          const paket = String(row[4] || '').trim(); // tanpa fallback karangan: kosong di sheet = kosong di web
          const users = String(row[5] || '').trim(); // tanpa fallback karangan
          const dueAmount = Number(row[6]) || 0;

          // Bulan 1 - 12 (Januari s/d Desember) berada di kolom index 7 s/d 18
          const monthlyPayments = [];
          for (let m = 0; m < 12; m++) {
            const rawVal = row[7 + m];
            if (rawVal === '' || rawVal === null || rawVal === undefined) {
              monthlyPayments.push(null);
            } else {
              const paidAmount = Number(rawVal);
              if (isNaN(paidAmount) || paidAmount <= 0) {
                monthlyPayments.push({
                  paid: 0,
                  due: dueAmount,
                  status: 'Belum Bayar',
                  year: currentYear
                });
              } else if (paidAmount >= dueAmount && dueAmount > 0) {
                monthlyPayments.push({
                  paid: paidAmount,
                  due: dueAmount,
                  status: 'Lunas',
                  year: currentYear
                });
              } else {
                monthlyPayments.push({
                  paid: paidAmount,
                  due: dueAmount,
                  status: 'Belum Lunas',
                  year: currentYear
                });
              }
            }
          }

          // Format baris billing: [no, nama, idPel, paket, users, dueAmount, jan, feb, ...]
          const bRow = [no, nama, idPel, paket, users, dueAmount].concat(monthlyPayments);
          billingRows.push(bRow);
        }
      }

      // 3. Hitung ringkasan KPI untuk Dashboard
      const currentMonth = new Date().getMonth();
      let collected = 0;
      let outstanding = 0;
      let unpaid = 0;

      billingRows.forEach(b => {
        const pay = b[6 + currentMonth];
        if (pay) {
          collected += Number(pay.paid || 0);
          if (pay.status === 'Belum Lunas') outstanding++;
          if (pay.status === 'Belum Bayar') unpaid++;
        }
      });

      const totalClients = customerRows.length;
      const activeClients = customerRows.filter(r => r[22] === 'Aktif').length;
      const promoClients = customerRows.filter(r => r[19] === 'Ya').length;

      return createJsonResponse({
        status: 'success',
        data: {
          customerRows: customerRows,
          billingRows: billingRows,
          kpi: {
            totalClients: totalClients,
            activeClients: activeClients,
            promoClients: promoClients,
            paymentCollected: collected,
            paymentOutstanding: outstanding,
            paymentUnpaid: unpaid,
            currentMonthName: MONTH_NAMES[currentMonth],
            year: currentYear
          }
        }
      });
    }

    return createJsonResponse({ status: 'error', message: 'Action tidak dikenal: ' + action });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Handle HTTP POST Requests
 * Digunakan untuk:
 *   - Tambah Pelanggan Baru (addCustomer)
 *   - Edit Pelanggan (updateCustomer)
 *   - Hapus Pelanggan (deleteCustomer)
 *   - Catat / Ubah Pembayaran Iuran (updatePayment)
 */
// ==========================================
// HIGHLIGHT PERUBAHAN TERAKHIR DI GOOGLE SHEET
// - Tambah data baru : seluruh baris baru diwarnai HIJAU JRENG
// - Edit data        : seluruh baris yang diedit diwarnai BIRU JRENG
//   (permintaan Capt 2026-11-03: edit = 1 baris full, bukan hanya 1 sel)
// - Setiap ada tambah/edit baru, highlight sebelumnya dibersihkan dulu,
//   jadi yang berwarna selalu hanya perubahan TERAKHIR.
// ==========================================
const HL_COLOR_ADD = '#00b050';  // hijau jreng (permintaan Capt: warna mencolok)
const HL_COLOR_EDIT = '#00b0f0'; // biru jreng
const HL_PROP_KEY = 'jacnet_last_highlight';
const HL_COLOR_PAY = '#ffff00'; // kuning jreng = uang masuk (pembayaran)
const HL_PROP_PAY = 'jacnet_last_payment';

function colToA1(col) { // nomor kolom 1-based -> huruf A1 (1=A, 27=AA, 47=AU)
  let out = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    out = String.fromCharCode(65 + m) + out;
    col = Math.floor((col - 1) / 26);
  }
  return out;
}

function hlNorm(v) { // samakan representasi nilai untuk perbandingan sebelum/sesudah edit
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

function clearLastHighlight(ss, key) {
  key = key || HL_PROP_KEY;
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(key);
    if (!raw) return;
    JSON.parse(raw).forEach(function (item) {
      const sh = ss.getSheetByName(item.sheet);
      if (sh) sh.getRange(item.range).setBackground(null);
    });
    props.deleteProperty(key);
  } catch (e) { /* highlight hanya kosmetik — jangan pernah gagalkan operasi utama */ }
}

function clearPaymentHighlight(ss) { clearLastHighlight(ss, HL_PROP_PAY); }

function applyHighlight(ss, items, key) {
  key = key || HL_PROP_KEY;
  // items: [{ sheet: 'CLIENTS', range: 'A800:AU800' / 'C5,F5', color: HL_COLOR_ADD }]
  try {
    const store = [];
    items.forEach(function (item) {
      const sh = ss.getSheetByName(item.sheet);
      if (!sh) return;
      sh.getRange(item.range).setBackground(item.color);
      store.push({ sheet: item.sheet, range: item.range });
    });
    if (store.length) {
      PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(store));
    }
  } catch (e) { /* abaikan */ }
}

// Kolom teks murni di CLIENTS (1-based): NIK 8, RW 11, NO TLP 13, SSID 37, PW 38,
// Perangkat (SN) 40, USER PPPOE 41. Tanpa ini, Google Sheet mengubah isian yang
// tampak seperti angka menjadi angka — nomor HP '0812...' kehilangan nol depannya.
// Mekanisme sesuai arahan Capt: cukup awali nilai dengan tanda petik (').
const TEXT_COLS_CLIENTS = [8, 11, 13, 37, 38, 40, 41];
function forceTextCells(sheet, rowIdx, client) {
  TEXT_COLS_CLIENTS.forEach(function(col) {
    const v = client[col - 1];
    if (v === undefined || v === null || v === '') return;
    const s = String(v);
    sheet.getRange(rowIdx, col).setValue(s.charAt(0) === "'" ? s : "'" + s);
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  // Tunggu hingga 30 detik untuk menghindari konflik data bersamaan
  if (!lock.tryLock(30000)) {
    return createJsonResponse({ status: 'error', message: 'Server sedang sibuk, silakan coba sesaat lagi.' });
  }

  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (ex) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const clientsSheet = ss.getSheetByName(SHEET_CLIENTS);
    const tagihanSheet = ss.getSheetByName(SHEET_TAGIHAN);

    if (!clientsSheet || !tagihanSheet) {
      return createJsonResponse({ status: 'error', message: 'Sheet CLIENTS atau TAGIHAN tidak ditemukan.' });
    }

    // ==========================================
    // 1. ACTION: ADD CUSTOMER
    // ==========================================
    if (action === 'addCustomer') {
      const client = payload.client; // Array 47 kolom
      if (!client || !client[8]) {
        return createJsonResponse({ status: 'error', message: 'Nama Pelanggan wajib diisi.' });
      }

      // Auto-generate ID Pelanggan jika kosong atau bertuliskan Otomatis
      let idPelanggan = client[6] ? String(client[6]).trim() : '';
      if (!idPelanggan || idPelanggan.toLowerCase() === 'otomatis' || idPelanggan.toUpperCase() === 'AUTO') {
        idPelanggan = generateNextCustomerId(clientsSheet);
      }
      client[6] = idPelanggan;

      const namaPelanggan = String(client[8]).trim();
      const paket = String(client[14] || '').trim();
      // TIDAK ADA nilai karangan: paket/tagihan yang kosong di web tetap kosong di sheet.
      const tagihanNormal = Number(client[42] || client[41] || 0);

      // Pastikan ID Pelanggan belum digunakan
      const existing = findRowById(clientsSheet, 7, idPelanggan); // Kolom G (index 7 di sheet)
      if (existing > 0) {
        return createJsonResponse({ status: 'error', message: 'ID Pelanggan ' + idPelanggan + ' sudah terdaftar!' });
      }

      // Tambahkan ke sheet CLIENTS
      const nextNo = clientsSheet.getLastRow();
      client[0] = nextNo;
      clientsSheet.appendRow(client);
      forceTextCells(clientsSheet, clientsSheet.getLastRow(), client); // jaga nol depan (HP/RW/dll)

      // Tambahkan ke sheet TAGIHAN
      const nextTagihanNo = tagihanSheet.getLastRow();
      const newBillingRow = [
        nextTagihanNo,
        idPelanggan,
        namaPelanggan,
        clusterNameForId(idPelanggan, client[11]), // Cluster/Wilayah dari awalan ID (sesuai data GSheet)
        paket,
        (payload.billing && payload.billing.users ? String(payload.billing.users).trim() : ''), // Jumlah User persis dari form dashboard (R = pelanggan rumahan), tanpa bawaan karangan
        tagihanNormal,
        '', '', '', '', '', '', '', '', '', '', '', '' // 12 bulan kosong
      ];
      tagihanSheet.appendRow(newBillingRow);

      // Tandai baris baru HIJAU MUDA (highlight perubahan sebelumnya dibersihkan dulu)
      const newClientRowIdx = clientsSheet.getLastRow();
      const newBillingRowIdx = tagihanSheet.getLastRow();
      clearLastHighlight(ss);
      clearPaymentHighlight(ss);
      applyHighlight(ss, [
        { sheet: SHEET_CLIENTS, range: 'A' + newClientRowIdx + ':' + colToA1(47) + newClientRowIdx, color: HL_COLOR_ADD },
        { sheet: SHEET_TAGIHAN, range: 'A' + newBillingRowIdx + ':' + colToA1(19) + newBillingRowIdx, color: HL_COLOR_ADD }
      ]);

      return createJsonResponse({
        status: 'success',
        message: 'Pelanggan baru berhasil disimpan.',
        id: idPelanggan
      });
    }

    // ==========================================
    // 2. ACTION: UPDATE CUSTOMER
    // ==========================================
    if (action === 'updateCustomer') {
      const client = payload.client;
      if (!client || !client[6]) {
        return createJsonResponse({ status: 'error', message: 'ID Pelanggan tidak boleh kosong.' });
      }

      const idPelanggan = String(client[6]).trim();
      const rowIndex = findRowById(clientsSheet, 7, idPelanggan);
      if (rowIndex < 2) {
        return createJsonResponse({ status: 'error', message: 'Data pelanggan ' + idPelanggan + ' tidak ditemukan.' });
      }

      // Update baris di sheet CLIENTS (47 kolom)
      const range = clientsSheet.getRange(rowIndex, 1, 1, 47);
      const oldValues = range.getValues()[0];
      client[0] = oldValues[0]; // Pertahankan nomor lama
      // Grup SO (Thn/Bulan/Tanggal/Hari SO) tidak pernah berubah setelah pertama
      // kali terisi — edit data lain tidak boleh menggeser Tanggal SO.
      if (hlNorm(oldValues[3]) !== '') {
        client[1] = oldValues[1];
        client[2] = oldValues[2];
        client[3] = oldValues[3];
        client[4] = oldValues[4];
      }
      range.setValues([client]);
      forceTextCells(clientsSheet, rowIndex, client); // jaga nol depan (HP/RW/dll)

      // Sinkronkan Nama & Paket ke sheet TAGIHAN jika ada
      const billingRowIndex = findRowById(tagihanSheet, 2, idPelanggan); // Kolom B di TAGIHAN
      if (billingRowIndex >= 2) {
        tagihanSheet.getRange(billingRowIndex, 3).setValue(client[8]); // Update Nama
        tagihanSheet.getRange(billingRowIndex, 4).setValue(clusterNameForId(idPelanggan, client[11])); // Sinkron Cluster/Wilayah
        tagihanSheet.getRange(billingRowIndex, 5).setValue(client[14] || ''); // Update Paket (selalu — kosong di web = kosong di sheet)
        if (payload.billing && payload.billing.users) {
          tagihanSheet.getRange(billingRowIndex, 6).setValue(String(payload.billing.users).trim()); // Update Jumlah User (R = rumahan)
        }
        const due = Number(client[42] || client[41] || 0);
        tagihanSheet.getRange(billingRowIndex, 7).setValue(due); // Update Tagihan (selalu — sumber tunggal Harga Paket Normal)
      }

      // Terapkan highlight edit: bersihkan tanda lama, warnai SATU BARIS PENUH
      // (CLIENTS 47 kolom + TAGIHAN 19 kolom) dengan biru jreng — permintaan
      // Capt 2026-11-03: edit ditandai 1 baris full, bukan hanya sel berubah.
      clearLastHighlight(ss);
      clearPaymentHighlight(ss);
      const hlItems = [
        { sheet: SHEET_CLIENTS, range: 'A' + rowIndex + ':' + colToA1(47) + rowIndex, color: HL_COLOR_EDIT }
      ];
      if (billingRowIndex >= 2) {
        hlItems.push({ sheet: SHEET_TAGIHAN, range: 'A' + billingRowIndex + ':' + colToA1(19) + billingRowIndex, color: HL_COLOR_EDIT });
      }
      applyHighlight(ss, hlItems);

      return createJsonResponse({ status: 'success', message: 'Data pelanggan berhasil diperbarui.' });
    }

    // ==========================================
    // 3. ACTION: DELETE CUSTOMER
    // ==========================================
    if (action === 'deleteCustomer') {
      const idPelanggan = String(payload.idPelanggan || '').trim();
      if (!idPelanggan) {
        return createJsonResponse({ status: 'error', message: 'ID Pelanggan tidak valid.' });
      }

      clearLastHighlight(ss); // bersihkan dulu selagi nomor baris masih akurat
      clearPaymentHighlight(ss);
      const clientRowIndex = findRowById(clientsSheet, 7, idPelanggan);
      if (clientRowIndex >= 2) {
        clientsSheet.deleteRow(clientRowIndex);
      }

      const billingRowIndex = findRowById(tagihanSheet, 2, idPelanggan);
      if (billingRowIndex >= 2) {
        tagihanSheet.deleteRow(billingRowIndex);
      }

      return createJsonResponse({ status: 'success', message: 'Pelanggan berhasil dihapus.' });
    }

    // ==========================================
    // 4. ACTION: UPDATE PAYMENT (CATAT IURAN)
    // ==========================================
    if (action === 'updatePayment') {
      const idPelanggan = String(payload.idPelanggan || '').trim();
      const monthIndex = Number(payload.monthIndex); // 0 = Januari, 11 = Desember
      const paidAmount = Number(payload.paid);

      if (!idPelanggan || isNaN(monthIndex) || monthIndex < 0 || monthIndex > 11) {
        return createJsonResponse({ status: 'error', message: 'Parameter pembayaran tidak valid.' });
      }

      const billingRowIndex = findRowById(tagihanSheet, 2, idPelanggan);
      if (billingRowIndex < 2) {
        return createJsonResponse({ status: 'error', message: 'Pelanggan tidak ditemukan di data tagihan.' });
      }

      // Kolom pembayaran Januari ada di kolom 8 (H), Februari di 9 (I), dst.
      const targetCol = 8 + monthIndex;
      const payCell = tagihanSheet.getRange(billingRowIndex, targetCol);
      const oldPaid = payCell.getValue();
      payCell.setValue(paidAmount);

      // Uang masuk ditandai KUNING JRENG pada sel bulan yang dibayar (penanda
      // pembayaran sendiri — tidak menghapus penanda data hijau/biru dari edit).
      // PENGECUALIAN: pembayaran 0 pada sel yang masih kosong hanyalah inisialisasi
      // otomatis dari dashboard (biasanya tepat setelah tambah pelanggan) — itu bukan
      // perubahan nyata, jadi tidak ditandai sama sekali.
      const isInitZero = paidAmount === 0 && hlNorm(oldPaid) === '';
      if (!isInitZero && hlNorm(oldPaid) !== String(paidAmount)) {
        clearPaymentHighlight(ss);
        applyHighlight(ss, [{ sheet: SHEET_TAGIHAN, range: colToA1(targetCol) + billingRowIndex, color: HL_COLOR_PAY }], HL_PROP_PAY);
      }

      return createJsonResponse({
        status: 'success',
        message: 'Pembayaran bulan ' + MONTH_NAMES[monthIndex] + ' untuk ' + idPelanggan + ' berhasil dicatat.'
      });
    }

    return createJsonResponse({ status: 'error', message: 'Action POST tidak dikenal: ' + action });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mencari nomor baris (1-based) berdasarkan nilai ID pada kolom tertentu
 */
function findRowById(sheet, colIndex, targetId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const data = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  const search = String(targetId).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === search) {
      return i + 1; // 1-based row index
    }
  }
  return -1;
}

/**
 * Otomatis menghasilkan ID Pelanggan berikutnya berformat JAC-XXXX
 * Menghitung nomor urut tertinggi dari kolom G (index 7)
 */
function generateNextCustomerId(clientsSheet) {
  const lastRow = clientsSheet.getLastRow();
  let maxNum = lastRow > 1 ? (lastRow - 1) : 0;
  if (lastRow > 1) {
    const idValues = clientsSheet.getRange(2, 7, lastRow - 1, 1).getValues();
    for (let i = 0; i < idValues.length; i++) {
      const val = String(idValues[i][0] || '').trim();
      const match = val.match(/(?:JAC|PW|CL|C|PMT|CKL|PRM)[-_]?(\d+)/i);
      if (match) {
        const num = parseInt(match[1], 10);
        if (!isNaN(num) && num > maxNum) maxNum = num;
      }
    }
  }
  return 'JAC-' + String(maxNum + 1).padStart(4, '0');
}

