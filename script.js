/* ==========================================================
   KEISHA KOMPUTER - SCRIPT
   Daftar isi:
     1. Pengaturan (akun, hak akses, kunci penyimpanan)
     2. Fungsi bantu
     3. Login, logout, dan perpindahan halaman
     4. Admin - Laporan Penjualan (filter periode)
     5. Ekspor Excel, Cetak/PDF, Backup & Restore
     6. Admin - Stok & Inventori
     7. Karyawan - Input Kas Harian
     8. Mulai aplikasi

   CATATAN KEAMANAN: ini login sisi-browser (tanpa server).
   Password di bawah bisa dilihat siapa pun yang membuka file ini.
   Cocok untuk belajar / satu komputer toko. Untuk sistem sungguhan,
   pindahkan pengecekan password ke server/database.
   ========================================================== */

/* ===== 1. PENGATURAN ===== */

/* GANTI password di bawah ini sebelum dipakai! */
const AKUN = {
  admin:    { password: "admin123",    role: "admin" },
  karyawan: { password: "karyawan123", role: "karyawan" },
};

/* Halaman yang boleh dibuka tiap peran (yang pertama = halaman awal) */
const AKSES = {
  admin:    ["laporan", "stok"],
  karyawan: ["kas"],
};

const KEY_SESI = "keisha_session";        // sessionStorage (hilang saat tab ditutup)
const KEY_TRANSAKSI = "keisha_transaksi"; // localStorage
const KEY_STOK = "keisha_stok";           // localStorage

const KATEGORI_KAS = {
  Pemasukan: ["Service Komputer", "Penjualan Barang", "Penjualan Sparepart", "Lainnya"],
  Pengeluaran: ["Pembelian Stok", "Operasional", "Gaji", "Lainnya"],
};

/* Kategori yang mengurangi stok barang saat dicatat */
const KATEGORI_BARANG = ["Penjualan Barang", "Penjualan Sparepart"];

/* ===== 2. FUNGSI BANTU ===== */
const $ = (id) => document.getElementById(id);

function baca(storage, kunci, cadangan) {
  try {
    const data = JSON.parse(storage.getItem(kunci));
    return data === null || data === undefined ? cadangan : data;
  } catch (e) {
    return cadangan;
  }
}

function tulis(storage, kunci, nilai) {
  try {
    storage.setItem(kunci, JSON.stringify(nilai));
    return true;
  } catch (e) {
    return false;
  }
}

const ambilTransaksi = () => { const d = baca(localStorage, KEY_TRANSAKSI, []); return Array.isArray(d) ? d : []; };
const simpanTransaksi = (d) => tulis(localStorage, KEY_TRANSAKSI, d);
const ambilStok = () => { const d = baca(localStorage, KEY_STOK, []); return Array.isArray(d) ? d : []; };
const simpanStok = (d) => tulis(localStorage, KEY_STOK, d);

function buatId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

/* Transaksi lama (dari versi sebelumnya) belum punya id: beri id sekali saja */
function siapkanData() {
  const data = ambilTransaksi();
  let berubah = false;
  data.forEach((t) => { if (!t.id) { t.id = buatId(); berubah = true; } });
  if (berubah) simpanTransaksi(data);
}

function formatRupiah(n) {
  n = Number(n) || 0;
  return (n < 0 ? "-Rp " : "Rp ") + Math.abs(n).toLocaleString("id-ID");
}

function dua(x) { return String(x).padStart(2, "0"); }

function isoTanggal(d) {
  return d.getFullYear() + "-" + dua(d.getMonth() + 1) + "-" + dua(d.getDate());
}

function hariIni() {
  return isoTanggal(new Date()); // tanggal lokal (bukan UTC) supaya tidak mundur sehari
}

function jamSekarang() {
  const d = new Date();
  return dua(d.getHours()) + ":" + dua(d.getMinutes());
}

function formatTanggal(iso) {
  const [y, m, d] = String(iso).split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
}

/* Membuat <td> dengan teks aman (tanpa innerHTML) */
function td(teks, kelas) {
  const el = document.createElement("td");
  el.textContent = teks;
  if (kelas) el.className = kelas;
  return el;
}

function barisKosong(jumlahKolom, teks) {
  const tr = document.createElement("tr");
  tr.className = "empty-row";
  const el = document.createElement("td");
  el.colSpan = jumlahKolom;
  el.textContent = teks;
  tr.appendChild(el);
  return tr;
}

function buatTombol(label, kelas, aksi) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = kelas;
  b.textContent = label;
  b.addEventListener("click", aksi);
  return b;
}

function jumlahPerJenis(jenis, daftar) {
  return daftar.filter((t) => t.jenis === jenis).reduce((n, t) => n + Number(t.nominal), 0);
}

function unduh(namaFile, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = namaFile;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/* ===== 3. LOGIN, LOGOUT, PERPINDAHAN HALAMAN ===== */
let sesi = null;

function bacaSesi() {
  const s = baca(sessionStorage, KEY_SESI, null);
  return s && AKSES[s.role] ? s : null;
}

function namaKaryawan() {
  const u = sesi && sesi.user ? sesi.user : "karyawan";
  return u.charAt(0).toUpperCase() + u.slice(1);
}

function tampilkanLogin() {
  sesi = null;
  $("aplikasi").hidden = true;
  $("halaman-login").hidden = false;
  $("login-password").value = "";
  $("login-error").hidden = true;
  $("login-username").focus();
}

function tampilkanAplikasi() {
  siapkanData();
  $("halaman-login").hidden = true;
  $("aplikasi").hidden = false;
  $("sidebar").hidden = sesi.role !== "admin";
  pindahHalaman();
}

/* Halaman yang diminta lewat alamat (#laporan, #stok, #kas), dicek hak akses */
function pindahHalaman() {
  if (!sesi) return;

  const boleh = AKSES[sesi.role];
  let halaman = location.hash.replace("#", "");
  if (boleh.indexOf(halaman) === -1) {
    halaman = boleh[0];
    try { history.replaceState(null, "", "#" + halaman); } catch (e) { location.hash = halaman; }
  }

  ["laporan", "stok", "kas"].forEach((h) => { $("view-" + h).hidden = h !== halaman; });
  document.querySelectorAll("[data-nav]").forEach((a) => {
    a.classList.toggle("active", a.dataset.nav === halaman);
  });

  if (halaman === "laporan") renderLaporan();
  if (halaman === "stok") renderStok();
  if (halaman === "kas") mulaiKas();
}

function prosesLogin(e) {
  e.preventDefault();
  const form = $("form-login");
  const user = $("login-username").value.trim().toLowerCase();
  const pass = $("login-password").value;
  const pesan = $("login-error");

  function gagal(teks) {
    pesan.textContent = teks;
    pesan.hidden = false;
    form.classList.remove("gagal");
    void form.offsetWidth; // ulangi animasi
    form.classList.add("gagal");
  }

  if (!user || !pass) return gagal("Username dan password wajib diisi.");

  const akun = Object.prototype.hasOwnProperty.call(AKUN, user) ? AKUN[user] : null;
  if (!akun || akun.password !== pass) {
    $("login-password").value = "";
    $("login-password").focus();
    return gagal("Username atau password salah.");
  }

  if (!tulis(sessionStorage, KEY_SESI, { user: user, role: akun.role })) {
    return gagal("Browser memblokir penyimpanan sesi. Aktifkan penyimpanan lalu coba lagi.");
  }

  sesi = { user: user, role: akun.role };
  $("login-username").value = "";
  try { history.replaceState(null, "", "#" + AKSES[sesi.role][0]); } catch (err) { location.hash = AKSES[sesi.role][0]; }
  tampilkanAplikasi();
}

function logout() {
  try { sessionStorage.removeItem(KEY_SESI); } catch (e) {}
  try { history.replaceState(null, "", location.pathname + location.search); } catch (e) { location.hash = ""; }
  tampilkanLogin();
}

/* ===== 4. ADMIN - LAPORAN PENJUALAN ===== */

/* Rentang tanggal sesuai pilihan periode: {dari, sampai, label} (ISO, boleh null) */
function rentangPeriode() {
  const pilihan = $("lap-periode").value;
  const sekarang = new Date();
  const hari = isoTanggal(sekarang);

  if (pilihan === "hari") {
    return { dari: hari, sampai: hari, label: "Hari ini (" + formatTanggal(hari) + ")" };
  }
  if (pilihan === "minggu") {
    const awal = isoTanggal(new Date(sekarang.getFullYear(), sekarang.getMonth(), sekarang.getDate() - 6));
    return { dari: awal, sampai: hari, label: "7 hari terakhir (" + formatTanggal(awal) + " - " + formatTanggal(hari) + ")" };
  }
  if (pilihan === "bulan") {
    const awal = isoTanggal(new Date(sekarang.getFullYear(), sekarang.getMonth(), 1));
    const akhir = isoTanggal(new Date(sekarang.getFullYear(), sekarang.getMonth() + 1, 0));
    return { dari: awal, sampai: akhir, label: "Bulan ini (" + formatTanggal(awal) + " - " + formatTanggal(akhir) + ")" };
  }
  if (pilihan === "kustom") {
    let dari = $("lap-dari").value || null;
    let sampai = $("lap-sampai").value || null;
    if (dari && sampai && dari > sampai) { const x = dari; dari = sampai; sampai = x; }
    const teksDari = dari ? formatTanggal(dari) : "awal";
    const teksSampai = sampai ? formatTanggal(sampai) : "sekarang";
    return { dari: dari, sampai: sampai, label: teksDari + " - " + teksSampai };
  }
  return { dari: null, sampai: null, label: "Semua periode" };
}

function dataLaporan() {
  const r = rentangPeriode();
  const data = ambilTransaksi().filter((t) =>
    (!r.dari || t.tanggal >= r.dari) && (!r.sampai || t.tanggal <= r.sampai));
  return { data: data, label: r.label };
}

/* Kelompokkan per tanggal, terbaru di atas */
function rekapPerHari(data) {
  const perHari = {};
  data.forEach((t) => {
    if (!perHari[t.tanggal]) perHari[t.tanggal] = [];
    perHari[t.tanggal].push(t);
  });
  return Object.keys(perHari).sort().reverse().map((tgl) => {
    const m = jumlahPerJenis("Pemasukan", perHari[tgl]);
    const k = jumlahPerJenis("Pengeluaran", perHari[tgl]);
    return { tanggal: tgl, masuk: m, keluar: k, selisih: m - k };
  });
}

/* Terbaru di atas: urut tanggal + jam, lalu urutan input */
function urutTerbaru(data) {
  return data.map((t, i) => ({ t: t, i: i }))
    .sort((a, b) => (b.t.tanggal + (b.t.jam || "")).localeCompare(a.t.tanggal + (a.t.jam || "")) || b.i - a.i)
    .map((x) => x.t);
}

function renderLaporan() {
  const lap = dataLaporan();
  const data = lap.data;

  const masuk = jumlahPerJenis("Pemasukan", data);
  const keluar = jumlahPerJenis("Pengeluaran", data);
  $("lap-omset").textContent = formatRupiah(masuk);
  $("lap-keluar").textContent = formatRupiah(keluar);
  $("lap-laba").textContent = formatRupiah(masuk - keluar);
  $("lap-cetak-judul").textContent = "Laporan Penjualan - Keisha Komputer | Periode: " + lap.label;

  const rekap = $("lap-rekap");
  rekap.innerHTML = "";
  if (data.length === 0) {
    rekap.appendChild(barisKosong(4, "Tidak ada data pada periode ini."));
  } else {
    rekapPerHari(data).forEach((r) => {
      const tr = document.createElement("tr");
      [formatTanggal(r.tanggal), formatRupiah(r.masuk), formatRupiah(r.keluar), formatRupiah(r.selisih)]
        .forEach((s) => tr.appendChild(td(s)));
      rekap.appendChild(tr);
    });
  }

  const detail = $("lap-detail");
  detail.innerHTML = "";
  if (data.length === 0) {
    detail.appendChild(barisKosong(6, "Tidak ada data pada periode ini."));
  } else {
    urutTerbaru(data).forEach((t) => {
      const tr = document.createElement("tr");
      [formatTanggal(t.tanggal), t.karyawan || "-", t.jenis, t.kategori || "-", t.keterangan || "-", formatRupiah(t.nominal)]
        .forEach((s) => tr.appendChild(td(s)));
      detail.appendChild(tr);
    });
  }
}

function aturPeriode() {
  const kustom = $("lap-periode").value === "kustom";
  $("lap-dari-wrap").hidden = !kustom;
  $("lap-sampai-wrap").hidden = !kustom;
  renderLaporan();
}

/* ===== 5. EKSPOR EXCEL, CETAK/PDF, BACKUP & RESTORE ===== */

/* ----- Penulis file .xlsx sederhana (tanpa library) -----
   File .xlsx = arsip ZIP berisi XML. Di sini dibuat ZIP tanpa kompresi. */
const TABEL_CRC = (function () {
  const t = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = TABEL_CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function buatZip(berkas) { // berkas: [{nama, isi(string)}]
  const enc = new TextEncoder();
  const bagian = [];
  const pusat = [];
  let posisi = 0;

  berkas.forEach((f) => {
    const nama = enc.encode(f.nama);
    const isi = enc.encode(f.isi);
    const crc = crc32(isi);

    const lokal = new DataView(new ArrayBuffer(30));
    lokal.setUint32(0, 0x04034b50, true);
    lokal.setUint16(4, 20, true);
    lokal.setUint16(6, 0x0800, true);   // nama file UTF-8
    lokal.setUint16(8, 0, true);        // tanpa kompresi
    lokal.setUint16(10, 0, true);
    lokal.setUint16(12, 0x21, true);    // tanggal 1980-01-01
    lokal.setUint32(14, crc, true);
    lokal.setUint32(18, isi.length, true);
    lokal.setUint32(22, isi.length, true);
    lokal.setUint16(26, nama.length, true);
    lokal.setUint16(28, 0, true);
    bagian.push(new Uint8Array(lokal.buffer), nama, isi);

    const p = new DataView(new ArrayBuffer(46));
    p.setUint32(0, 0x02014b50, true);
    p.setUint16(4, 20, true);
    p.setUint16(6, 20, true);
    p.setUint16(8, 0x0800, true);
    p.setUint16(10, 0, true);
    p.setUint16(12, 0, true);
    p.setUint16(14, 0x21, true);
    p.setUint32(16, crc, true);
    p.setUint32(20, isi.length, true);
    p.setUint32(24, isi.length, true);
    p.setUint16(28, nama.length, true);
    p.setUint32(42, posisi, true);
    pusat.push(new Uint8Array(p.buffer), nama);

    posisi += 30 + nama.length + isi.length;
  });

  const ukuranPusat = pusat.reduce((n, b) => n + b.length, 0);
  const akhir = new DataView(new ArrayBuffer(22));
  akhir.setUint32(0, 0x06054b50, true);
  akhir.setUint16(8, berkas.length, true);
  akhir.setUint16(10, berkas.length, true);
  akhir.setUint32(12, ukuranPusat, true);
  akhir.setUint32(16, posisi, true);

  return new Blob(bagian.concat(pusat, [new Uint8Array(akhir.buffer)]),
    { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function xmlAman(s) {
  return String(s).replace(/[<>&"']/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", '"': "&quot;", "'": "&apos;" }[c]))
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}

/* Jenis sel: teks, teks tebal, rupiah, rupiah tebal, tanggal */
const SEL = {
  teks: (v) => ({ t: "s", v: v, s: 0 }),
  tebal: (v) => ({ t: "s", v: v, s: 1 }),
  rupiah: (v) => ({ t: "n", v: Number(v) || 0, s: 2 }),
  tanggal: (iso) => {
    const [y, m, d] = String(iso).split("-").map(Number);
    return { t: "n", v: Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569, s: 3 };
  },
  rupiahTebal: (v) => ({ t: "n", v: Number(v) || 0, s: 4 }),
};

function hurufKolom(i) { return String.fromCharCode(65 + i); }

function xmlSheet(sheet) {
  const kolom = sheet.lebar.map((w, i) =>
    '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join("");
  const baris = sheet.baris.map((row, r) => {
    const sel = row.map((c, i) => {
      if (c === null || c === undefined) return "";
      const ref = hurufKolom(i) + (r + 1);
      if (c.t === "n") return '<c r="' + ref + '" s="' + c.s + '"><v>' + c.v + "</v></c>";
      return '<c r="' + ref + '" t="inlineStr" s="' + c.s + '"><is><t xml:space="preserve">' + xmlAman(c.v) + "</t></is></c>";
    }).join("");
    return '<row r="' + (r + 1) + '">' + sel + "</row>";
  }).join("");
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>' + kolom +
    "</cols><sheetData>" + baris + "</sheetData></worksheet>";
}

function buatXlsx(sheets) {
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

  const berkas = [
    { nama: "[Content_Types].xml", isi: HEAD +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      sheets.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join("") +
      "</Types>" },
    { nama: "_rels/.rels", isi: HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      "</Relationships>" },
    { nama: "xl/workbook.xml", isi: HEAD + "<workbook " + NS + " " + R + "><sheets>" +
      sheets.map((s, i) => '<sheet name="' + xmlAman(s.nama) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") +
      "</sheets></workbook>" },
    { nama: "xl/_rels/workbook.xml.rels", isi: HEAD +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join("") +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>" },
    { nama: "xl/styles.xml", isi: HEAD + "<styleSheet " + NS + ">" +
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="&quot;Rp &quot;#,##0"/><numFmt numFmtId="165" formatCode="dd\\ mmm\\ yyyy"/></numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="5">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>' +
      "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
  ];
  sheets.forEach((s, i) => berkas.push({ nama: "xl/worksheets/sheet" + (i + 1) + ".xml", isi: xmlSheet(s) }));
  return buatZip(berkas);
}

function eksporExcel() {
  const lap = dataLaporan();
  if (lap.data.length === 0) {
    alert("Tidak ada data pada periode ini untuk diekspor.");
    return;
  }
  const T = SEL.teks, B = SEL.tebal, Rp = SEL.rupiah;

  const rekap = rekapPerHari(lap.data);
  const totMasuk = jumlahPerJenis("Pemasukan", lap.data);
  const totKeluar = jumlahPerJenis("Pengeluaran", lap.data);

  const sheetRekap = {
    nama: "Rekap Per Hari",
    lebar: [18, 18, 18, 18],
    baris: [
      [B("Laporan Penjualan - Keisha Komputer")],
      [T("Periode: " + lap.label)],
      [],
      [B("Tanggal"), B("Pemasukan"), B("Pengeluaran"), B("Selisih")],
    ].concat(rekap.map((r) => [SEL.tanggal(r.tanggal), Rp(r.masuk), Rp(r.keluar), Rp(r.selisih)]))
     .concat([[B("TOTAL"), SEL.rupiahTebal(totMasuk), SEL.rupiahTebal(totKeluar), SEL.rupiahTebal(totMasuk - totKeluar)]]),
  };

  const sheetDetail = {
    nama: "Detail Transaksi",
    lebar: [16, 8, 14, 14, 20, 44, 18],
    baris: [
      [B("Tanggal"), B("Jam"), B("Karyawan"), B("Jenis"), B("Kategori"), B("Keterangan"), B("Nominal")],
    ].concat(urutTerbaru(lap.data).map((t) =>
      [SEL.tanggal(t.tanggal), T(t.jam || "-"), T(t.karyawan || "-"), T(t.jenis), T(t.kategori || "-"), T(t.keterangan || "-"), Rp(t.nominal)])),
  };

  unduh("laporan-keisha-" + hariIni() + ".xlsx", buatXlsx([sheetRekap, sheetDetail]));
}

function cetakPdf() {
  renderLaporan();
  window.print(); // di jendela cetak pilih "Simpan sebagai PDF"
}

/* ----- Backup & Restore ----- */
function backupData() {
  const isi = {
    app: "keisha-komputer",
    versi: 1,
    dibuat: new Date().toISOString(),
    transaksi: ambilTransaksi(),
    stok: ambilStok(),
  };
  unduh("backup-keisha-" + hariIni() + ".json", new Blob([JSON.stringify(isi, null, 2)], { type: "application/json" }));
}

function bersihkanBackup(obj) {
  if (!obj || obj.app !== "keisha-komputer" || !Array.isArray(obj.transaksi) || !Array.isArray(obj.stok)) return null;
  const pola = /^\d{4}-\d{2}-\d{2}$/;
  const transaksi = [];
  const stok = [];

  for (const t of obj.transaksi) {
    if (!t || !pola.test(t.tanggal) || ["Pemasukan", "Pengeluaran"].indexOf(t.jenis) === -1 || !isFinite(Number(t.nominal))) return null;
    const baru = {
      id: typeof t.id === "string" && t.id ? t.id : buatId(),
      tanggal: t.tanggal,
      karyawan: String(t.karyawan || ""),
      jenis: t.jenis,
      kategori: String(t.kategori || ""),
      keterangan: String(t.keterangan || ""),
      nominal: Number(t.nominal),
    };
    if (t.jam) baru.jam = String(t.jam);
    if (t.stokId !== undefined && t.stokId !== null && isFinite(Number(t.qty))) { baru.stokId = t.stokId; baru.qty = Number(t.qty); }
    transaksi.push(baru);
  }

  for (const b of obj.stok) {
    if (!b || typeof b.nama !== "string" || !b.nama.trim() || !isFinite(Number(b.stok))) return null;
    stok.push({
      id: b.id !== undefined && b.id !== null ? b.id : Date.now() + Math.random(),
      nama: b.nama,
      kategori: String(b.kategori || ""),
      stok: Math.max(0, Number(b.stok)),
      hargaBeli: Number(b.hargaBeli) || 0,
      hargaJual: Number(b.hargaJual) || 0,
    });
  }
  return { transaksi: transaksi, stok: stok };
}

function restoreData(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function () {
    let bersih = null;
    try { bersih = bersihkanBackup(JSON.parse(reader.result)); } catch (e) { bersih = null; }
    if (!bersih) {
      alert("File backup tidak valid. Pilih file .json hasil Backup Data dari website ini.");
      return;
    }
    const ok = confirm("Data saat ini akan DIGANTI dengan isi backup:\n" +
      bersih.transaksi.length + " transaksi dan " + bersih.stok.length + " barang.\n\nLanjutkan?");
    if (!ok) return;
    if (!simpanTransaksi(bersih.transaksi) || !simpanStok(bersih.stok)) {
      alert("Gagal menyimpan. Penyimpanan browser diblokir atau penuh.");
      return;
    }
    pindahHalaman();
    alert("Restore berhasil.");
  };
  reader.onerror = function () { alert("File tidak bisa dibaca."); };
  reader.readAsText(file);
}

/* ===== 6. ADMIN - STOK & INVENTORI ===== */
function renderStok() {
  const semua = ambilStok();
  const kata = $("stok-cari").value.trim().toLowerCase();
  const tampil = semua.filter((b) =>
    String(b.nama).toLowerCase().includes(kata) || String(b.kategori || "").toLowerCase().includes(kata));

  const tbody = $("stok-daftar");
  tbody.innerHTML = "";

  if (semua.length === 0) return tbody.appendChild(barisKosong(6, "Belum ada barang. Tambahkan lewat form di atas."));
  if (tampil.length === 0) return tbody.appendChild(barisKosong(6, "Barang tidak ditemukan."));

  tampil.forEach((b) => {
    const tr = document.createElement("tr");
    tr.appendChild(td(b.nama));
    tr.appendChild(td(b.kategori || "-"));

    const tdStok = document.createElement("td");
    const wrap = document.createElement("div");
    wrap.className = "stok-ctrl";
    wrap.appendChild(buatTombol("-", "btn-mini", () => ubahStok(b.id, -1)));
    const angka = document.createElement("span");
    angka.textContent = b.stok;
    if (b.stok <= 3) angka.className = "stok-rendah";
    wrap.appendChild(angka);
    wrap.appendChild(buatTombol("+", "btn-mini", () => ubahStok(b.id, 1)));
    tdStok.appendChild(wrap);
    tr.appendChild(tdStok);

    tr.appendChild(td(formatRupiah(b.hargaBeli)));
    tr.appendChild(td(formatRupiah(b.hargaJual)));

    const tdAksi = document.createElement("td");
    tdAksi.appendChild(buatTombol("Hapus", "btn-hapus", () => hapusBarang(b.id)));
    tr.appendChild(tdAksi);

    tbody.appendChild(tr);
  });
}

function ubahStok(id, delta) {
  const data = ambilStok();
  const item = data.find((b) => b.id === id);
  if (!item) return;
  item.stok = Math.max(0, item.stok + delta);
  simpanStok(data);
  renderStok();
}

function hapusBarang(id) {
  const item = ambilStok().find((b) => b.id === id);
  if (!item) return;
  if (!confirm('Hapus "' + item.nama + '" dari daftar?')) return;
  simpanStok(ambilStok().filter((b) => b.id !== id));
  renderStok();
}

function tambahBarang(e) {
  e.preventDefault();
  const data = ambilStok();
  data.push({
    id: Date.now() + Math.random(),
    nama: $("stok-nama").value.trim(),
    kategori: $("stok-kategori").value.trim(),
    stok: Number($("stok-jumlah").value) || 0,
    hargaBeli: Number($("stok-beli").value) || 0,
    hargaJual: Number($("stok-jual").value) || 0,
  });
  simpanStok(data);
  $("form-stok").reset();
  $("stok-nama").focus();
  renderStok();
}

/* ===== 7. KARYAWAN - INPUT KAS HARIAN ===== */
let sedangSimpan = false;
let ketOtomatis = ""; // teks keterangan yang diisi otomatis dari pilihan barang

function pesanKas(teks, error) {
  $("kas-pesan").textContent = teks;
  $("kas-pesan").className = "kas-pesan" + (error ? " error" : "");
}

function isiKategoriKas() {
  const pilih = $("kas-kategori");
  pilih.innerHTML = "";
  KATEGORI_KAS[$("kas-jenis").value].forEach((k) => {
    const o = document.createElement("option");
    o.value = k;
    o.textContent = k;
    pilih.appendChild(o);
  });
}

/* ----- Penjualan barang: pilih barang dari stok ----- */
function penjualanBarang() {
  return $("kas-jenis").value === "Pemasukan" && KATEGORI_BARANG.indexOf($("kas-kategori").value) !== -1;
}

function aturFieldBarang() {
  const tampil = penjualanBarang();
  $("kas-barang-wrap").hidden = !tampil;
  $("kas-jumlah-wrap").hidden = !tampil;
  if (ketOtomatis && $("kas-keterangan").value === ketOtomatis) $("kas-keterangan").value = "";
  ketOtomatis = "";
  $("kas-jumlah").value = "1";

  const pilih = $("kas-barang");
  pilih.innerHTML = "";
  if (!tampil) return;

  const kosong = document.createElement("option");
  kosong.value = "";
  kosong.textContent = "(Bukan dari stok - isi manual)";
  pilih.appendChild(kosong);

  ambilStok().forEach((b) => {
    const o = document.createElement("option");
    o.value = String(b.id);
    o.textContent = b.nama + " (stok " + b.stok + ") - " + formatRupiah(b.hargaJual);
    if (b.stok <= 0) o.disabled = true;
    pilih.appendChild(o);
  });
}

function barangTerpilih() {
  const v = $("kas-barang").value;
  if (!v) return null;
  return ambilStok().find((b) => String(b.id) === v) || null;
}

/* Isi nominal & keterangan otomatis dari barang + jumlah */
function hitungBarang() {
  const item = barangTerpilih();
  const ket = $("kas-keterangan");

  if (!item) {
    if (ket.value === ketOtomatis) ket.value = "";
    ketOtomatis = "";
    return;
  }

  const jumlah = Math.max(1, Math.floor(Number($("kas-jumlah").value) || 1));
  $("kas-nominal").value = item.hargaJual * jumlah;

  const baru = item.nama + " x" + jumlah;
  if (ket.value === "" || ket.value === ketOtomatis) ket.value = baru;
  ketOtomatis = baru;
}

function renderKas() {
  const tanggal = $("kas-tanggal").value;
  const nama = namaKaryawan();
  const data = ambilTransaksi().filter((t) => t.tanggal === tanggal && t.karyawan === nama);
  const masuk = jumlahPerJenis("Pemasukan", data);
  const keluar = jumlahPerJenis("Pengeluaran", data);

  const hariIniAtau = tanggal === hariIni() ? "Hari Ini" : formatTanggal(tanggal);
  $("kas-label-masuk").textContent = "Pemasukan " + hariIniAtau;
  $("kas-label-keluar").textContent = "Pengeluaran " + hariIniAtau;
  $("kas-masuk").textContent = formatRupiah(masuk);
  $("kas-keluar").textContent = formatRupiah(keluar);
  $("kas-selisih").textContent = formatRupiah(masuk - keluar);

  $("kas-judul-riwayat").textContent =
    tanggal === hariIni() ? "Transaksi Saya Hari Ini" : "Transaksi Saya (" + formatTanggal(tanggal) + ")";

  const tbody = $("kas-daftar");
  tbody.innerHTML = "";
  if (data.length === 0) {
    tbody.appendChild(barisKosong(6, tanggal === hariIni() ? "Belum ada transaksi hari ini." : "Belum ada transaksi pada tanggal ini."));
    return;
  }
  data.slice().reverse().forEach((t) => { // terbaru di atas
    const tr = document.createElement("tr");
    tr.appendChild(td(t.jam || "-"));
    tr.appendChild(td(t.jenis, t.jenis === "Pemasukan" ? "jenis-masuk" : "jenis-keluar"));
    tr.appendChild(td(t.kategori || "-"));
    tr.appendChild(td(t.keterangan || "-"));
    tr.appendChild(td(formatRupiah(t.nominal)));
    const tdAksi = document.createElement("td");
    tdAksi.appendChild(buatTombol("Hapus", "btn-hapus", () => hapusKas(t.id)));
    tr.appendChild(tdAksi);
    tbody.appendChild(tr);
  });
}

function mulaiKas() {
  $("kas-tanggal").value = hariIni();
  isiKategoriKas();
  aturFieldBarang();
  pesanKas("");
  renderKas();
}

function simpanKas(e) {
  e.preventDefault();
  const nominal = Number($("kas-nominal").value);
  const keterangan = $("kas-keterangan").value.trim();

  if (!$("kas-tanggal").value || !$("kas-kategori").value || !keterangan || !(nominal > 0)) {
    return pesanKas("Lengkapi semua kolom dengan benar.", true);
  }

  const transaksi = {
    id: buatId(),
    tanggal: $("kas-tanggal").value,
    jam: jamSekarang(),
    karyawan: namaKaryawan(),
    jenis: $("kas-jenis").value,
    kategori: $("kas-kategori").value,
    keterangan: keterangan,
    nominal: nominal,
  };

  /* Penjualan dari stok: kurangi stok barang */
  const stokSebelum = ambilStok();
  let stokBaru = null;
  const item = penjualanBarang() ? barangTerpilih() : null;
  if (item) {
    const qty = Number($("kas-jumlah").value);
    if (!Number.isInteger(qty) || qty < 1) return pesanKas("Jumlah barang harus bilangan bulat minimal 1.", true);
    if (qty > item.stok) return pesanKas("Stok " + item.nama + " tidak cukup (sisa " + item.stok + ").", true);
    stokBaru = stokSebelum.map((b) => (b.id === item.id ? Object.assign({}, b, { stok: b.stok - qty }) : b));
    transaksi.stokId = item.id;
    transaksi.qty = qty;
  }

  if (stokBaru && !simpanStok(stokBaru)) return pesanKas("Gagal menyimpan stok. Penyimpanan browser diblokir.", true);
  const data = ambilTransaksi();
  data.push(transaksi);
  if (!simpanTransaksi(data)) {
    if (stokBaru) simpanStok(stokSebelum); // batalkan pengurangan stok
    return pesanKas("Gagal menyimpan. Penyimpanan browser diblokir.", true);
  }

  const tanggal = $("kas-tanggal").value;
  const jenis = $("kas-jenis").value;
  sedangSimpan = true; // supaya event "reset" tidak menimpa pesan & tanggal
  $("form-kas").reset();
  $("kas-tanggal").value = tanggal; // tetap di tanggal yang sama untuk input berikutnya
  $("kas-jenis").value = jenis;
  isiKategoriKas();
  aturFieldBarang();
  renderKas();
  $("kas-keterangan").focus();
  pesanKas(item ? "Transaksi disimpan. Stok " + item.nama + " berkurang " + transaksi.qty + "." : "Transaksi berhasil disimpan.");
}

function hapusKas(id) {
  const data = ambilTransaksi();
  const t = data.find((x) => x.id === id);
  if (!t) return;
  if (!confirm('Hapus transaksi "' + t.keterangan + '" (' + formatRupiah(t.nominal) + ")?" +
      (t.stokId !== undefined ? "\nStok barang akan dikembalikan." : ""))) return;

  /* Kembalikan stok jika transaksi ini penjualan dari stok */
  if (t.stokId !== undefined) {
    const stok = ambilStok();
    const item = stok.find((b) => b.id === t.stokId);
    if (item) {
      item.stok += Number(t.qty) || 0;
      simpanStok(stok);
    }
  }

  simpanTransaksi(data.filter((x) => x.id !== id));
  renderKas();
  pesanKas("Transaksi dihapus.");
}

/* ===== 8. MULAI APLIKASI ===== */
$("form-login").addEventListener("submit", prosesLogin);
$("login-username").addEventListener("input", () => { $("login-error").hidden = true; });
$("login-password").addEventListener("input", () => { $("login-error").hidden = true; });
$("btn-logout").addEventListener("click", logout);
window.addEventListener("hashchange", pindahHalaman);

/* Laporan */
$("lap-periode").addEventListener("change", aturPeriode);
$("lap-dari").addEventListener("change", renderLaporan);
$("lap-sampai").addEventListener("change", renderLaporan);
$("btn-excel").addEventListener("click", eksporExcel);
$("btn-pdf").addEventListener("click", cetakPdf);

/* Backup & restore */
$("btn-backup").addEventListener("click", backupData);
$("btn-restore").addEventListener("click", () => $("file-restore").click());
$("file-restore").addEventListener("change", (e) => {
  restoreData(e.target.files[0]);
  e.target.value = ""; // supaya file yang sama bisa dipilih lagi
});

/* Stok */
$("form-stok").addEventListener("submit", tambahBarang);
$("stok-cari").addEventListener("input", renderStok);

/* Kas */
$("form-kas").addEventListener("submit", simpanKas);
$("form-kas").addEventListener("reset", () => {
  setTimeout(() => {
    if (sedangSimpan) { sedangSimpan = false; return; }
    $("kas-tanggal").value = hariIni();
    isiKategoriKas();
    aturFieldBarang();
    renderKas();
    pesanKas("");
  }, 0);
});
$("kas-jenis").addEventListener("change", () => { isiKategoriKas(); aturFieldBarang(); });
$("kas-kategori").addEventListener("change", aturFieldBarang);
$("kas-barang").addEventListener("change", hitungBarang);
$("kas-jumlah").addEventListener("input", hitungBarang);
$("kas-tanggal").addEventListener("change", () => { renderKas(); pesanKas(""); });

sesi = bacaSesi();
if (sesi) {
  tampilkanAplikasi();
} else {
  tampilkanLogin();
}