/* ==========================================================
   KEISHA KOMPUTER - SCRIPT
   Daftar isi:
     1. Pengaturan (Supabase, hak akses, kategori)
     2. Fungsi bantu
     2b. Data dari Supabase (muat & pembaruan otomatis)
     3. Login, logout, dan perpindahan halaman
     4. Admin - Laporan Penjualan (filter periode)
     5. Ekspor Excel, Cetak/PDF, Backup & Restore
     6. Admin - Stok & Inventori
     7. Karyawan - Input Kas Harian
     8. Mulai aplikasi

   DATABASE: Supabase (project "sistemkasir").
   - Kunci di bawah adalah kunci PUBLIK, memang boleh terlihat di sini.
     Data dilindungi Row Level Security di database:
       admin    : semua laporan, kelola stok, backup & restore
       karyawan : lihat stok, catat & hapus transaksinya sendiri
   - Akun dibuat di Dashboard Supabase > Authentication > Users dengan
     email <username>@keisha.local (contoh: admin@keisha.local).
     Di halaman login cukup ketik username-nya saja, misalnya "admin".
   ========================================================== */

/* ===== 1. PENGATURAN ===== */
const SUPABASE_URL = "https://crzwbshhdbykvhohdouk.supabase.co";
const SUPABASE_KEY = "sb_publishable_di_W3KjljZTMemZweWj6qA_dcZ20CX_"; // kunci publik
const DOMAIN_AKUN = "@keisha.local"; // username "admin" -> admin@keisha.local

/* Halaman yang boleh dibuka tiap peran (yang pertama = halaman awal) */
const AKSES = {
  admin:    ["laporan", "stok", "kas", "service"],
  karyawan: ["kas", "service"],
};

/* Data versi lama (sebelum Supabase) yang mungkin masih tersimpan di browser ini */
const KEY_TRANSAKSI_LAMA = "keisha_transaksi";
const KEY_STOK_LAMA = "keisha_stok";
const KEY_DATA_LAMA_SELESAI = "keisha_data_lama_selesai";

const KATEGORI_KAS = {
  Pemasukan: ["Service Komputer", "Penjualan Barang", "Penjualan Sparepart", "Lainnya"],
  Pengeluaran: ["Pembelian Stok", "Operasional", "Gaji", "Lainnya"],
};

/* Kategori yang mengurangi stok barang saat dicatat */
const KATEGORI_BARANG = ["Penjualan Barang", "Penjualan Sparepart"];

/* Pengeluaran yang hanya boleh dicatat admin (pemilik). Aturan yang sama dipasang di database. */
const KATEGORI_KHUSUS_ADMIN = ["Gaji", "Pembelian Stok"];

/* Kategori pengeluaran yang menambah stok barang (admin) */
const KATEGORI_PEMBELIAN = "Pembelian Stok";

/* Status service. Diambil & Batal hanya lewat tombol Lunasi / Batal */
const STATUS_AKTIF = ["Diterima", "Dikerjakan", "Selesai"];
const STATUS_SERVICE = STATUS_AKTIF.concat(["Diambil", "Batal"]);

/* Format tanggal yang disimpan: YYYY-MM-DD (tahun 4 digit) */
const POLA_TANGGAL = /^\d{4}-\d{2}-\d{2}$/;

/* Koneksi Supabase. Sesi login disimpan di sessionStorage: hilang saat tab ditutup */
const db = (function () {
  if (!window.supabase) return null; // library gagal dimuat (mis. tidak ada internet)
  let simpanan;
  try { simpanan = window.sessionStorage; } catch (e) { simpanan = undefined; }
  return window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY, {
    auth: { storage: simpanan, persistSession: true, autoRefreshToken: true },
  });
})();

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

function buatId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function formatRupiah(n) {
  n = Number(n) || 0;
  return (n < 0 ? "-Rp " : "Rp ") + Math.abs(n).toLocaleString("id-ID");
}

/* ----- Kolom uang: tampil "Rp 1.500.000" saat diketik ----- */
const MAKS_DIGIT_RUPIAH = 13; // sampai Rp 9.999.999.999.999

/* Ambil angkanya saja dari kolom uang ("Rp 1.500.000" -> 1500000). Kosong -> NaN */
function angkaRupiah(input) {
  const digit = input.value.replace(/\D/g, "");
  return digit ? Number(digit) : NaN;
}

/* Isi kolom uang dari angka (dipakai saat nominal dihitung otomatis) */
function isiRupiah(input, angka) {
  input.value = isFinite(angka) && angka >= 0 ? formatRupiah(Math.round(angka)) : "";
}

/* Format ulang saat diketik/ditempel; posisi kursor dijaga agar tidak lompat ke ujung */
function formatKetikRupiah(input) {
  const lama = input.value;
  const kursor = input.selectionStart === null ? lama.length : input.selectionStart;
  const digitDiKanan = lama.slice(kursor).replace(/\D/g, "").length;

  const digit = lama
    .replace(/,\d{0,2}\s*$/, "")   // buang sen dari teks tempelan, mis. "Rp 2.500.000,00"
    .replace(/\D/g, "")
    .replace(/^0+(?=\d)/, "")      // "05" -> "5"
    .slice(0, MAKS_DIGIT_RUPIAH);
  const baru = digit ? formatRupiah(Number(digit)) : "";
  if (baru === lama) return;
  input.value = baru;

  if (document.activeElement !== input || !baru) return;
  let posisi = baru.length;
  for (let dilewati = 0; posisi > 0 && dilewati < Math.min(digitDiKanan, digit.length); ) {
    posisi--;
    if (/\d/.test(baru[posisi])) dilewati++;
  }
  posisi = Math.max(posisi, baru.search(/\d/)); // jangan di depan tulisan "Rp "
  input.setSelectionRange(posisi, posisi);
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

/* Pesan error yang mudah dimengerti */
function teksGagal(error) {
  const pesan = String((error && error.message) || error || "");
  if (/fetch|network|load failed/i.test(pesan)) return "Tidak bisa terhubung ke server. Periksa koneksi internet.";
  return pesan || "Terjadi kesalahan.";
}

/* ===== 2b. DATA DARI SUPABASE =====
   Data disalin ke memori (dataStok, dataTransaksi) supaya tampilan tetap cepat,
   lalu diperbarui otomatis saat ada perubahan dari perangkat lain (Realtime). */
let dataStok = [];
let dataTransaksi = [];
let dataService = [];
const ambilTransaksi = () => dataTransaksi;
const ambilStok = () => dataStok;

let operasiBerjalan = 0; // jumlah simpanan yang belum selesai
let versiLokal = 0;      // naik setiap ada perubahan dari perangkat ini

/* Jalankan permintaan ke Supabase; tidak pernah melempar error */
async function kirim(permintaan) {
  versiLokal++;
  operasiBerjalan++;
  try {
    return await permintaan;
  } catch (e) {
    return { data: null, error: e };
  } finally {
    operasiBerjalan--;
  }
}

function dariBarisStok(r) {
  return {
    id: r.id,
    nama: r.nama,
    kategori: r.kategori || "",
    stok: Number(r.stok) || 0,
    hargaBeli: Number(r.harga_beli) || 0,
    hargaJual: Number(r.harga_jual) || 0,
  };
}

function dariBarisTransaksi(r) {
  const t = {
    id: r.id,
    tanggal: r.tanggal,
    jam: r.jam || "",
    karyawan: r.karyawan || "",
    jenis: r.jenis,
    kategori: r.kategori || "",
    keterangan: r.keterangan || "",
    nominal: Number(r.nominal) || 0,
    hpp: Number(r.hpp) || 0, // modal barang yang terjual (untuk hitung laba)
  };
  if (r.stok_id) { t.stokId = r.stok_id; t.qty = Number(r.qty) || 0; }
  if (r.service_id) { t.serviceId = r.service_id; t.serviceTahap = r.service_tahap || ""; } // uang dari halaman Service
  return t;
}

function dariBarisService(r) {
  return {
    id: r.id,
    nota: r.nota,
    tanggalMasuk: r.tanggal_masuk,
    pelanggan: r.pelanggan,
    noHp: r.no_hp || "",
    perangkat: r.perangkat,
    kelengkapan: r.kelengkapan || "",
    keluhan: r.keluhan || "",
    perkiraan: Number(r.perkiraan) || 0,
    status: r.status,
    biayaJasa: Number(r.biaya_jasa) || 0,
    total: Number(r.total) || 0,
    dibayar: Number(r.dibayar) || 0,
    alasanBatal: r.alasan_batal || "",
    dicatatOleh: r.dicatat_oleh || "",
    diambilAt: r.diambil_at || null,
    dibatalkanAt: r.dibatalkan_at || null,
  };
}

/* Supabase membatasi 1000 baris per permintaan: ambil bertahap sampai habis */
async function ambilSemuaBaris(tabel) {
  const UKURAN = 1000;
  const hasil = [];
  for (let dari = 0; ; dari += UKURAN) {
    const { data, error } = await db.from(tabel).select("*")
      .order("created_at", { ascending: true }).order("id", { ascending: true })
      .range(dari, dari + UKURAN - 1);
    if (error) throw error;
    hasil.push(...data);
    if (data.length < UKURAN) return hasil;
  }
}

/* Muat semua data yang boleh dilihat akun ini (diatur RLS di database).
   Hasil false = ada perubahan dari perangkat ini saat memuat, coba lagi nanti. */
async function muatData() {
  const versi = versiLokal;
  const [stok, transaksi, service] = await Promise.all(
    [ambilSemuaBaris("stok"), ambilSemuaBaris("transaksi"), ambilSemuaBaris("service")]);
  if (versi !== versiLokal || operasiBerjalan > 0) return false;
  dataStok = stok.map(dariBarisStok);
  dataTransaksi = transaksi.map(dariBarisTransaksi);
  dataService = service.map(dariBarisService);
  return true;
}

/* Tampilkan ulang halaman yang sedang dibuka (isian form yang sedang diketik tidak direset) */
function tampilkanUlang() {
  if (!sesi) return;
  if (!$("view-laporan").hidden) renderLaporan();
  if (!$("view-stok").hidden) renderStok();
  if (!$("view-kas").hidden) { renderKas(); isiPilihanBarang(); }
  if (!$("view-service").hidden) renderService();
}

let jedaMuat = null;
function jadwalMuatUlang() {
  clearTimeout(jedaMuat);
  jedaMuat = setTimeout(async () => {
    if (!sesi) return;
    if (operasiBerjalan > 0) return jadwalMuatUlang(); // tunggu simpanan yang masih berjalan
    let berhasil = false;
    try { berhasil = await muatData(); } catch (e) { return; } // gagal: tetap pakai data terakhir
    if (!berhasil) return jadwalMuatUlang();
    tampilkanUlang();
  }, 400);
}

/* Realtime: muat ulang bila stok/transaksi diubah dari perangkat lain */
let saluran = null;
function mulaiRealtime() {
  hentikanRealtime();
  saluran = db.channel("data-kasir")
    .on("postgres_changes", { event: "*", schema: "public", table: "stok" }, jadwalMuatUlang)
    .on("postgres_changes", { event: "*", schema: "public", table: "transaksi" }, jadwalMuatUlang)
    .on("postgres_changes", { event: "*", schema: "public", table: "service" }, jadwalMuatUlang)
    .subscribe((status) => { if (status === "SUBSCRIBED") jadwalMuatUlang(); }); // termasuk setelah koneksi putus
}

function hentikanRealtime() {
  if (saluran) { db.removeChannel(saluran); saluran = null; }
}

/* ===== 3. LOGIN, LOGOUT, PERPINDAHAN HALAMAN ===== */
let sesi = null; // { user, nama, role }

/* Profil akun yang sedang login (dibaca dari tabel profil, hanya baris miliknya) */
async function ambilProfil() {
  const { data, error } = await db.from("profil").select("username, nama, peran").maybeSingle();
  if (error) throw error;
  return data && AKSES[data.peran] ? { user: data.username, nama: data.nama, role: data.peran } : null;
}

function namaKaryawan() {
  return sesi && sesi.nama ? sesi.nama : "Karyawan";
}

function tampilkanLogin(pesan) {
  sesi = null;
  $("aplikasi").hidden = true;
  $("halaman-login").hidden = false;
  $("login-password").value = "";
  $("login-error").textContent = pesan || "";
  $("login-error").hidden = !pesan;
  $("login-username").focus();
}

function tampilkanAplikasi() {
  $("halaman-login").hidden = true;
  $("aplikasi").hidden = false;
  /* Menu samping untuk semua; menu khusus admin (Laporan, Stok, Backup/Restore) disembunyikan dari karyawan */
  $("sidebar").hidden = false;
  document.querySelectorAll('[data-peran="admin"]').forEach((el) => { el.hidden = sesi.role !== "admin"; });
  pindahHalaman();
  mulaiRealtime();
  if (sesi.role === "admin") setTimeout(tawarkanDataLama, 300);
}

/* Bersihkan semua data akun di memori (logout / sesi habis) */
function akhiriSesi() {
  hentikanRealtime();
  clearTimeout(jedaMuat);
  versiLokal++; // abaikan hasil muat data yang masih berjalan
  sesi = null;
  dataStok = [];
  dataTransaksi = [];
  dataService = [];
  panelService = null;
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

  ["laporan", "stok", "kas", "service"].forEach((h) => { $("view-" + h).hidden = h !== halaman; });
  document.querySelectorAll("[data-nav]").forEach((a) => {
    a.classList.toggle("active", a.dataset.nav === halaman);
  });

  if (halaman === "laporan") renderLaporan();
  if (halaman === "stok") renderStok();
  if (halaman === "kas") mulaiKas();
  if (halaman === "service") mulaiService();
}

let sedangLogin = false;

async function prosesLogin(e) {
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
  if (!db) return gagal("Gagal memuat Supabase. Periksa koneksi internet lalu muat ulang halaman.");
  if (sedangLogin) return; // tombol Masuk diklik berkali-kali
  sedangLogin = true;

  try {
    const email = user.indexOf("@") === -1 ? user + DOMAIN_AKUN : user;
    const { error } = await db.auth.signInWithPassword({ email: email, password: pass });
    if (error) {
      if (error.status === 429) return gagal("Terlalu banyak percobaan. Tunggu sebentar lalu coba lagi.");
      if (error.code === "email_not_confirmed") return gagal("Akun belum dikonfirmasi. Buat ulang akun dengan mencentang Auto Confirm User.");
      if (error.status === 400) {
        $("login-password").value = "";
        $("login-password").focus();
        return gagal("Username atau password salah.");
      }
      return gagal(teksGagal(error));
    }

    const profil = await ambilProfil();
    if (!profil) {
      await db.auth.signOut({ scope: "local" });
      return gagal("Akun ini belum punya hak akses.");
    }
    await muatData();
    sesi = profil;
  } catch (err) {
    return gagal(teksGagal(err));
  } finally {
    sedangLogin = false;
  }

  $("login-username").value = "";
  try { history.replaceState(null, "", "#" + AKSES[sesi.role][0]); } catch (err) { location.hash = AKSES[sesi.role][0]; }
  tampilkanAplikasi();
}

function logout() {
  akhiriSesi();
  try { history.replaceState(null, "", location.pathname + location.search); } catch (e) { location.hash = ""; }
  tampilkanLogin();
  if (db) db.auth.signOut({ scope: "local" }).catch(() => {}); // "local": perangkat lain tetap login
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

/* Laba = omset - modal barang yang terjual - biaya (gaji, operasional, lainnya).
   Pembelian Stok bukan biaya: uangnya berubah jadi barang, modalnya dihitung saat barang itu terjual. */
function hitungLaba(data) {
  const omset = jumlahPerJenis("Pemasukan", data);
  const modal = data.filter((t) => t.jenis === "Pemasukan").reduce((n, t) => n + (Number(t.hpp) || 0), 0);
  const biaya = data.filter((t) => t.jenis === "Pengeluaran" && t.kategori !== KATEGORI_PEMBELIAN)
    .reduce((n, t) => n + (Number(t.nominal) || 0), 0);
  return { modal: modal, biaya: biaya, laba: omset - modal - biaya };
}

function renderLaporan() {
  const lap = dataLaporan();
  const data = lap.data;

  const masuk = jumlahPerJenis("Pemasukan", data);
  const keluar = jumlahPerJenis("Pengeluaran", data);
  $("lap-omset").textContent = formatRupiah(masuk);
  $("lap-keluar").textContent = formatRupiah(keluar);
  $("lap-laba").textContent = formatRupiah(hitungLaba(data).laba);
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
  const laba = hitungLaba(lap.data);

  const sheetRekap = {
    nama: "Rekap Per Hari",
    lebar: [26, 18, 18, 18],
    baris: [
      [B("Laporan Penjualan - Keisha Komputer")],
      [T("Periode: " + lap.label)],
      [],
      [B("Tanggal"), B("Pemasukan"), B("Pengeluaran"), B("Selisih")],
    ].concat(rekap.map((r) => [SEL.tanggal(r.tanggal), Rp(r.masuk), Rp(r.keluar), Rp(r.selisih)]))
     .concat([[B("TOTAL"), SEL.rupiahTebal(totMasuk), SEL.rupiahTebal(totKeluar), SEL.rupiahTebal(totMasuk - totKeluar)]])
     .concat([
       [],
       [T("Omset"), Rp(totMasuk)],
       [T("Modal barang terjual"), Rp(laba.modal)],
       [T("Biaya (tanpa beli stok)"), Rp(laba.biaya)],
       [B("LABA BERSIH"), SEL.rupiahTebal(laba.laba)],
     ]),
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
async function backupData() {
  try { await muatData(); } catch (e) { return alert("Gagal mengambil data terbaru: " + teksGagal(e)); }
  const isi = {
    app: "keisha-komputer",
    versi: 1,
    dibuat: new Date().toISOString(),
    transaksi: ambilTransaksi(),
    stok: ambilStok(),
    service: dataService,
  };
  unduh("backup-keisha-" + hariIni() + ".json", new Blob([JSON.stringify(isi, null, 2)], { type: "application/json" }));
}

function bersihkanBackup(obj) {
  if (!obj || obj.app !== "keisha-komputer" || !Array.isArray(obj.transaksi) || !Array.isArray(obj.stok)) return null;
  const transaksi = [];
  const stok = [];

  for (const t of obj.transaksi) {
    if (!t || !POLA_TANGGAL.test(t.tanggal) || ["Pemasukan", "Pengeluaran"].indexOf(t.jenis) === -1 || !isFinite(Number(t.nominal))) return null;
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
    if (t.hpp !== undefined && t.hpp !== null && isFinite(Number(t.hpp))) baru.hpp = Number(t.hpp);
    if (t.stokId !== undefined && t.stokId !== null && isFinite(Number(t.qty))) { baru.stokId = t.stokId; baru.qty = Number(t.qty); }
    if (t.serviceId) {
      baru.serviceId = String(t.serviceId);
      if (["dp", "lunas", "batal"].indexOf(t.serviceTahap) !== -1) baru.serviceTahap = t.serviceTahap;
    }
    transaksi.push(baru);
  }

  /* Data service (backup lama belum punya bagian ini) */
  const service = [];
  if (obj.service !== undefined) {
    if (!Array.isArray(obj.service)) return null;
    for (const s of obj.service) {
      if (!s || typeof s.nota !== "string" || !s.nota || !POLA_TANGGAL.test(s.tanggalMasuk) ||
          STATUS_SERVICE.indexOf(s.status) === -1) return null;
      service.push({
        id: String(s.id || buatId()),
        nota: s.nota,
        tanggalMasuk: s.tanggalMasuk,
        pelanggan: String(s.pelanggan || "").trim() || "(Tanpa nama)",
        noHp: String(s.noHp || ""),
        perangkat: String(s.perangkat || "").trim() || "(Tidak diisi)",
        kelengkapan: String(s.kelengkapan || ""),
        keluhan: String(s.keluhan || ""),
        perkiraan: Math.max(0, Number(s.perkiraan) || 0),
        status: s.status,
        biayaJasa: Number(s.biayaJasa) || 0,
        total: Number(s.total) || 0,
        dibayar: Number(s.dibayar) || 0,
        alasanBatal: String(s.alasanBatal || ""),
        dicatatOleh: String(s.dicatatOleh || ""),
        diambilAt: s.diambilAt || null,
        dibatalkanAt: s.dibatalkanAt || null,
      });
    }
  }

  for (const b of obj.stok) {
    if (!b || !isFinite(Number(b.stok))) return null;
    /* Barang tanpa nama tidak membuat seluruh backup ditolak: beri nama pengganti */
    const nama = typeof b.nama === "string" ? b.nama.trim() : "";
    stok.push({
      id: b.id !== undefined && b.id !== null ? b.id : Date.now() + Math.random(),
      nama: nama || "(Tanpa nama)",
      kategori: String(b.kategori || ""),
      stok: Math.max(0, Number(b.stok)),
      hargaBeli: Number(b.hargaBeli) || 0,
      hargaJual: Number(b.hargaJual) || 0,
    });
  }
  return { transaksi: transaksi, stok: stok, service: service };
}

function restoreData(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async function () {
    let bersih = null;
    try { bersih = bersihkanBackup(JSON.parse(reader.result)); } catch (e) { bersih = null; }
    if (!bersih) {
      alert("File backup tidak valid. Pilih file .json hasil Backup Data dari website ini.");
      return;
    }
    const ok = confirm("Semua data di database akan DIGANTI dengan isi backup:\n" +
      bersih.transaksi.length + " transaksi, " + bersih.stok.length + " barang, dan " +
      bersih.service.length + " service.\n\nLanjutkan?");
    if (!ok) return;
    /* Dijalankan di database sebagai satu kesatuan: gagal = data lama tetap utuh */
    const { error } = await kirim(db.rpc("restore_data",
      { p_stok: bersih.stok, p_transaksi: bersih.transaksi, p_service: bersih.service }));
    if (error) {
      alert("Restore gagal: " + teksGagal(error) + "\nData di database tidak berubah.");
      return;
    }
    try { await muatData(); } catch (e) { /* tampilkan data terakhir */ }
    pindahHalaman();
    alert("Restore berhasil.");
  };
  reader.onerror = function () { alert("File tidak bisa dibaca."); };
  reader.readAsText(file);
}

/* ----- Data versi lama (localStorage) -----
   Sebelum memakai Supabase, data disimpan di browser. Bila masih ada, tawarkan
   unduh sebagai file backup agar bisa dimasukkan ke database lewat Restore Data. */
function tawarkanDataLama() {
  if (!sesi || sesi.role !== "admin") return;
  let penyimpanan;
  try { penyimpanan = window.localStorage; } catch (e) { return; }
  if (baca(penyimpanan, KEY_DATA_LAMA_SELESAI, false)) return;

  let transaksi = baca(penyimpanan, KEY_TRANSAKSI_LAMA, []);
  let stok = baca(penyimpanan, KEY_STOK_LAMA, []);
  if (!Array.isArray(transaksi)) transaksi = [];
  if (!Array.isArray(stok)) stok = [];
  if (transaksi.length === 0 && stok.length === 0) return;

  const unduhSekarang = confirm("Ditemukan data lama di browser ini yang belum ada di database:\n" +
    transaksi.length + " transaksi dan " + stok.length + " barang.\n\n" +
    "Unduh sebagai file backup? Setelah itu masukkan ke database lewat menu Restore Data.");
  if (unduhSekarang) {
    const isi = { app: "keisha-komputer", versi: 1, dibuat: new Date().toISOString(), transaksi: transaksi, stok: stok };
    unduh("backup-data-lama-keisha-" + hariIni() + ".json", new Blob([JSON.stringify(isi, null, 2)], { type: "application/json" }));
    tulis(penyimpanan, KEY_DATA_LAMA_SELESAI, true);
    return;
  }
  if (confirm("Hapus data lama dari browser ini? Data yang dihapus tidak bisa dikembalikan.\n\n" +
      "Pilih Batal untuk menyimpannya dulu (pertanyaan ini muncul lagi saat login berikutnya).")) {
    try {
      penyimpanan.removeItem(KEY_TRANSAKSI_LAMA);
      penyimpanan.removeItem(KEY_STOK_LAMA);
    } catch (e) { /* abaikan */ }
  }
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

async function ubahStok(id, delta) {
  const item = dataStok.find((b) => b.id === id);
  if (!item) return;
  item.stok = Math.max(0, item.stok + delta); // tampil langsung, database menyusul
  renderStok();
  /* Dihitung di database (stok + selisih) agar klik beruntun / perangkat lain tidak saling menimpa */
  const { error } = await kirim(db.rpc("ubah_stok", { p_id: id, p_selisih: delta }));
  if (error) alert("Gagal mengubah stok: " + teksGagal(error));
  jadwalMuatUlang();
}

const sedangDihapus = new Set(); // cegah tombol Hapus diproses dua kali

async function hapusBarang(id) {
  const item = dataStok.find((b) => b.id === id);
  if (!item || sedangDihapus.has(id)) return;
  if (!confirm('Hapus "' + item.nama + '" dari daftar?')) return;
  sedangDihapus.add(id);
  const { error } = await kirim(db.from("stok").delete().eq("id", id));
  sedangDihapus.delete(id);
  if (error) return alert("Gagal menghapus barang: " + teksGagal(error));
  dataStok = dataStok.filter((b) => b.id !== id);
  renderStok();
  jadwalMuatUlang();
}

let sedangTambahBarang = false;

async function tambahBarang(e) {
  e.preventDefault();
  const nama = $("stok-nama").value.trim();
  if (!nama) {
    /* Nama berisi spasi saja: kosongkan agar browser memunculkan peringatan "wajib diisi" */
    $("stok-nama").value = "";
    $("stok-nama").reportValidity();
    return;
  }
  if (sedangTambahBarang) return; // tombol diklik dua kali: jangan tersimpan dobel
  sedangTambahBarang = true;
  const { data, error } = await kirim(db.from("stok").insert({
    nama: nama,
    kategori: $("stok-kategori").value.trim(),
    stok: Number($("stok-jumlah").value) || 0,
    harga_beli: angkaRupiah($("stok-beli")) || 0,
    harga_jual: angkaRupiah($("stok-jual")) || 0,
  }).select().single());
  sedangTambahBarang = false;
  if (error) return alert("Gagal menyimpan barang: " + teksGagal(error));

  dataStok.push(dariBarisStok(data));
  $("form-stok").reset();
  $("stok-nama").focus();
  renderStok();
  jadwalMuatUlang();
}

/* ===== 7. KARYAWAN - INPUT KAS HARIAN ===== */
let sedangSimpan = false;
let ketOtomatis = ""; // teks keterangan yang diisi otomatis dari pilihan barang
let tanggalIkutHariIni = true; // tanggal kas ikut berganti lewat tengah malam, kecuali karyawan memilih tanggal lain

function setTanggalHariIni() {
  $("kas-tanggal").value = hariIni();
  tanggalIkutHariIni = true;
}

function pesanKas(teks, error) {
  $("kas-pesan").textContent = teks;
  $("kas-pesan").className = "kas-pesan" + (error ? " error" : "");
}

function isiKategoriKas() {
  const pilih = $("kas-kategori");
  pilih.innerHTML = "";
  const admin = sesi && sesi.role === "admin";
  KATEGORI_KAS[$("kas-jenis").value]
    .filter((k) => admin || KATEGORI_KHUSUS_ADMIN.indexOf(k) === -1) // Gaji & Pembelian Stok hanya admin
    .forEach((k) => {
      const o = document.createElement("option");
      o.value = k;
      o.textContent = k;
      pilih.appendChild(o);
    });
}

/* ----- Barang dari stok -----
   "jual": penjualan barang/sparepart, stok berkurang
   "beli": pembelian stok oleh admin, stok bertambah */
function jenisBarang() {
  const jenis = $("kas-jenis").value;
  const kategori = $("kas-kategori").value;
  if (jenis === "Pemasukan" && KATEGORI_BARANG.indexOf(kategori) !== -1) return "jual";
  if (jenis === "Pengeluaran" && kategori === KATEGORI_PEMBELIAN && sesi && sesi.role === "admin") return "beli";
  return null;
}

function aturFieldBarang() {
  const tampil = jenisBarang() !== null;
  $("kas-barang-wrap").hidden = !tampil;
  $("kas-jumlah-wrap").hidden = !tampil;
  if (ketOtomatis && $("kas-keterangan").value === ketOtomatis) $("kas-keterangan").value = "";
  ketOtomatis = "";
  $("kas-jumlah").value = "1";
  $("kas-barang").innerHTML = ""; // mulai tanpa barang terpilih
  isiPilihanBarang();
}

/* Isi daftar barang dari stok terbaru; barang yang sedang dipilih tetap dipertahankan */
function isiPilihanBarang() {
  const pilih = $("kas-barang");
  const sebelumnya = pilih.value;
  pilih.innerHTML = "";
  const mode = jenisBarang();
  if (!mode) return;

  const kosong = document.createElement("option");
  kosong.value = "";
  kosong.textContent = "(Bukan dari stok - isi manual)";
  pilih.appendChild(kosong);

  ambilStok().forEach((b) => {
    const o = document.createElement("option");
    o.value = String(b.id);
    /* Penjualan: tampilkan harga jual. Pembelian: tampilkan harga beli */
    o.textContent = b.nama + " (stok " + b.stok + ") - " + formatRupiah(mode === "beli" ? b.hargaBeli : b.hargaJual);
    if (mode === "jual" && b.stok <= 0) o.disabled = true; // barang habis tidak bisa dijual, tapi bisa dibeli
    pilih.appendChild(o);
  });

  pilih.value = sebelumnya;
  if (pilih.value !== sebelumnya) { // barang yang dipilih sudah dihapus dari stok
    pilih.value = "";
    hitungBarang();
  }
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
  const beli = jenisBarang() === "beli";
  isiRupiah($("kas-nominal"), (beli ? item.hargaBeli : item.hargaJual) * jumlah);

  const baru = (beli ? "Beli " : "") + item.nama + " x" + jumlah;
  if (ket.value === "" || ket.value === ketOtomatis) ket.value = baru;
  ketOtomatis = baru;
}

function renderKas() {
  const tanggal = $("kas-tanggal").value || hariIni(); // kolom tanggal sedang dikosongkan: anggap hari ini
  const admin = sesi && sesi.role === "admin";
  const nama = namaKaryawan();
  /* Admin melihat semua transaksi pada tanggal itu (untuk dicek & dikoreksi); karyawan hanya miliknya */
  const data = ambilTransaksi().filter((t) => t.tanggal === tanggal && (admin || t.karyawan === nama));
  const masuk = jumlahPerJenis("Pemasukan", data);
  const keluar = jumlahPerJenis("Pengeluaran", data);

  const hariIniAtau = tanggal === hariIni() ? "Hari Ini" : formatTanggal(tanggal);
  $("kas-label-masuk").textContent = "Pemasukan " + hariIniAtau;
  $("kas-label-keluar").textContent = "Pengeluaran " + hariIniAtau;
  $("kas-masuk").textContent = formatRupiah(masuk);
  $("kas-keluar").textContent = formatRupiah(keluar);
  $("kas-selisih").textContent = formatRupiah(masuk - keluar);

  const judul = admin ? "Transaksi" : "Transaksi Saya";
  $("kas-judul-riwayat").textContent =
    tanggal === hariIni() ? judul + " Hari Ini" : judul + " (" + formatTanggal(tanggal) + ")";

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
    tr.appendChild(td((t.keterangan || "-") + (admin ? " · " + (t.karyawan || "-") : ""))); // admin: tampilkan siapa pencatatnya
    tr.appendChild(td(formatRupiah(t.nominal)));
    const tdAksi = document.createElement("td");
    if (t.serviceId) {
      /* Uang dari halaman Service hanya bisa dikoreksi lewat halaman Service */
      const ket = document.createElement("span");
      ket.className = "teks-aksi";
      ket.textContent = "dari Service";
      tdAksi.appendChild(ket);
    } else {
      tdAksi.appendChild(buatTombol("Hapus", "btn-hapus", () => hapusKas(t.id)));
    }
    tr.appendChild(tdAksi);
    tbody.appendChild(tr);
  });
}

function mulaiKas() {
  setTanggalHariIni();
  /* Karyawan hanya mencatat untuk hari ini (database juga memaksa tanggal hari ini); admin bebas */
  $("kas-tanggal").readOnly = !(sesi && sesi.role === "admin");
  isiKategoriKas();
  aturFieldBarang();
  pesanKas("");
  renderKas();
}

let sedangKirimKas = false;

async function simpanKas(e) {
  e.preventDefault();
  if (sedangKirimKas) return; // tombol Simpan diklik dua kali: jangan tersimpan dobel
  cekGantiHari(); // halaman dibiarkan terbuka lewat tengah malam: pakai tanggal hari baru
  const tanggal = $("kas-tanggal").value;
  const jenis = $("kas-jenis").value;
  const kategori = $("kas-kategori").value;
  const nominal = angkaRupiah($("kas-nominal"));
  const keterangan = $("kas-keterangan").value.trim();

  if (!tanggal || !kategori || !keterangan || !(nominal > 0)) {
    return pesanKas("Lengkapi semua kolom dengan benar.", true);
  }
  /* Tahun lebih dari 4 digit (salah ketik) merusak laporan & membuat backup ditolak */
  if (!POLA_TANGGAL.test(tanggal)) return pesanKas("Tanggal tidak valid. Periksa kembali tahunnya.", true);

  /* Barang dari stok: jual (stok berkurang) / beli (stok bertambah), diproses di database bersamaan dengan transaksi */
  const mode = jenisBarang();
  const item = mode ? barangTerpilih() : null;
  let qty = null;
  if (item) {
    qty = Number($("kas-jumlah").value);
    if (!Number.isInteger(qty) || qty < 1) return pesanKas("Jumlah barang harus bilangan bulat minimal 1.", true);
    if (mode === "jual" && qty > item.stok) return pesanKas("Stok " + item.nama + " tidak cukup (sisa " + item.stok + ").", true);
  }

  sedangKirimKas = true;
  pesanKas("Menyimpan...");
  const { data, error } = await kirim(db.rpc("simpan_transaksi", {
    p_tanggal: tanggal,
    p_jam: jamSekarang(),
    p_jenis: jenis,
    p_kategori: kategori,
    p_keterangan: keterangan,
    p_nominal: nominal,
    p_stok_id: item ? item.id : null,
    p_qty: qty,
  }));
  sedangKirimKas = false;

  if (error) {
    if (item && mode === "jual" && /Stok tidak cukup/i.test(String(error.message))) {
      /* Stok sudah berubah di perangkat lain: ambil angka terbaru */
      try { await muatData(); } catch (err) { /* abaikan */ }
      isiPilihanBarang();
      const kini = dataStok.find((b) => b.id === item.id);
      return pesanKas("Stok " + item.nama + " tidak cukup (sisa " + (kini ? kini.stok : 0) + ").", true);
    }
    return pesanKas("Gagal menyimpan: " + teksGagal(error), true);
  }

  /* Perbarui data di memori tanpa menunggu muat ulang */
  const baris = Array.isArray(data) ? data[0] : data;
  if (baris) dataTransaksi.push(dariBarisTransaksi(baris));
  if (item) {
    const b = dataStok.find((x) => x.id === item.id);
    if (b) b.stok = mode === "beli" ? b.stok + qty : Math.max(0, b.stok - qty);
  }

  sedangSimpan = true; // supaya event "reset" tidak menimpa pesan & tanggal
  $("form-kas").reset();
  $("kas-tanggal").value = tanggal; // tetap di tanggal yang sama untuk input berikutnya
  $("kas-jenis").value = jenis;
  isiKategoriKas();
  aturFieldBarang();
  renderKas();
  $("kas-keterangan").focus();
  pesanKas(item
    ? "Transaksi disimpan. Stok " + item.nama + (mode === "beli" ? " bertambah " : " berkurang ") + qty + "."
    : "Transaksi berhasil disimpan.");
  jadwalMuatUlang();
}

async function hapusKas(id) {
  const t = dataTransaksi.find((x) => x.id === id);
  if (!t || sedangDihapus.has(id)) return;
  if (t.serviceId) return pesanKas("Transaksi ini dicatat dari halaman Service. Koreksi lewat halaman Service.", true);
  const pembelian = t.jenis === "Pengeluaran"; // pembelian stok: menghapusnya mengurangi stok lagi
  if (!confirm('Hapus transaksi "' + t.keterangan + '" (' + formatRupiah(t.nominal) + ")?" +
      (t.stokId !== undefined ? (pembelian ? "\nStok barang akan dikurangi lagi." : "\nStok barang akan dikembalikan.") : ""))) return;

  /* Di database: transaksi dihapus dan stok disesuaikan sekaligus */
  sedangDihapus.add(id);
  const { error } = await kirim(db.rpc("hapus_transaksi", { p_id: id }));
  sedangDihapus.delete(id);
  if (error) return pesanKas("Gagal menghapus: " + teksGagal(error), true);

  dataTransaksi = dataTransaksi.filter((x) => x.id !== id);
  if (t.stokId !== undefined) {
    const item = dataStok.find((b) => b.id === t.stokId);
    const qty = Number(t.qty) || 0;
    if (item) item.stok = pembelian ? Math.max(0, item.stok - qty) : item.stok + qty;
  }
  renderKas();
  isiPilihanBarang();
  pesanKas("Transaksi dihapus.");
  jadwalMuatUlang();
}

/* ----- Ganti hari (halaman dibiarkan terbuka lewat tengah malam) ----- */
let hariTampil = hariIni();

function cekGantiHari() {
  const sekarang = hariIni();
  const gantiHari = sekarang !== hariTampil;
  hariTampil = sekarang;
  if (!sesi) return;

  if (!$("view-kas").hidden) {
    const geser = tanggalIkutHariIni && $("kas-tanggal").value !== sekarang;
    if (geser) $("kas-tanggal").value = sekarang;
    if (geser || gantiHari) renderKas();
  }
  if (gantiHari && !$("view-laporan").hidden) renderLaporan(); // filter "Hari Ini", "Bulan Ini", dst.
}

/* ===== 7b. SERVICE =====
   Service masuk (+DP), status pengerjaan, pelunasan (+sparepart dari stok),
   pembatalan (karyawan & admin), buka lagi (admin). Semua uang tercatat
   otomatis di Input Kas oleh database. */
let panelService = null; // panel yang terbuka: { jenis: "lunas" | "batal", id, sparepart: [{ stokId, qty }] }
let sedangKirimService = false;
let sedangProsesPanel = false;
let resetSetelahSimpanService = false;

const serviceAktif = (s) => STATUS_AKTIF.indexOf(s.status) !== -1;
/* Perkiraan sisa tagihan (biaya pasti baru diketahui saat pelunasan) */
const sisaService = (s) => (serviceAktif(s) ? Math.max(0, s.perkiraan - s.dibayar) : 0);
const cariService = (id) => dataService.find((s) => s.id === id) || null;

function pesanService(id, teks, error) {
  $(id).textContent = teks;
  $(id).className = "kas-pesan" + (error ? " error" : "");
}

/* <td> berisi teks utama + baris kecil di bawahnya */
function tdDuaBaris(utama, kecil) {
  const el = document.createElement("td");
  el.textContent = utama || "-";
  if (kecil) {
    const sp = document.createElement("span");
    sp.className = "teks-kecil";
    sp.textContent = kecil;
    el.appendChild(sp);
  }
  return el;
}

function terapkanBarisService(data) {
  const baris = Array.isArray(data) ? data[0] : data;
  if (!baris) return null;
  const baru = dariBarisService(baris);
  const i = dataService.findIndex((x) => x.id === baru.id);
  if (i >= 0) dataService[i] = baru; else dataService.push(baru);
  return baru;
}

function mulaiService() {
  tutupPanelService();
  pesanService("srv-pesan", "");
  renderService();
}

function renderService() {
  const semua = dataService;
  $("srv-proses").textContent = semua.filter((s) => s.status === "Diterima" || s.status === "Dikerjakan").length;
  $("srv-siap").textContent = semua.filter((s) => s.status === "Selesai").length;
  $("srv-sisa").textContent = formatRupiah(semua.reduce((n, s) => n + sisaService(s), 0));

  const saring = $("srv-saring").value;
  const kata = $("srv-cari").value.trim().toLowerCase();
  const tampil = semua.filter((s) =>
    (saring === "semua" || (saring === "aktif" ? serviceAktif(s) : s.status === saring)) &&
    (!kata || [s.nota, s.pelanggan, s.noHp, s.perangkat, s.keluhan].some((x) => String(x).toLowerCase().includes(kata))));

  const tbody = $("srv-daftar");
  tbody.innerHTML = "";
  if (tampil.length === 0) {
    tbody.appendChild(barisKosong(7, semua.length === 0
      ? "Belum ada service. Catat lewat form Service Masuk di atas."
      : "Tidak ada service yang cocok."));
  }
  tampil.slice().reverse().forEach((s) => tbody.appendChild(barisService(s))); // terbaru di atas

  perbaruiPanelService();
}

function barisService(s) {
  const tr = document.createElement("tr");
  tr.appendChild(td(s.nota));
  tr.appendChild(td(formatTanggal(s.tanggalMasuk)));
  tr.appendChild(tdDuaBaris(s.pelanggan, s.noHp));
  tr.appendChild(tdDuaBaris(s.perangkat,
    [s.keluhan, s.kelengkapan ? "Kelengkapan: " + s.kelengkapan : ""].filter(Boolean).join(" · ")));

  /* Status: bisa diubah langsung selama service masih berjalan */
  const tdStatus = document.createElement("td");
  if (serviceAktif(s)) {
    const pilih = document.createElement("select");
    pilih.className = "select-status status-" + s.status.toLowerCase();
    pilih.setAttribute("aria-label", "Status " + s.nota);
    STATUS_AKTIF.forEach((st) => {
      const o = document.createElement("option");
      o.value = st;
      o.textContent = st;
      pilih.appendChild(o);
    });
    pilih.value = s.status;
    pilih.addEventListener("change", () => ubahStatusService(s.id, pilih.value));
    tdStatus.appendChild(pilih);
  } else {
    const lencana = document.createElement("span");
    lencana.className = "status status-" + s.status.toLowerCase();
    lencana.textContent = s.status;
    tdStatus.appendChild(lencana);
  }
  tr.appendChild(tdStatus);

  /* Pembayaran */
  if (serviceAktif(s)) {
    tr.appendChild(tdDuaBaris(s.dibayar > 0 ? "DP " + formatRupiah(s.dibayar) : "Belum bayar",
      s.perkiraan > 0 ? "Perkiraan " + formatRupiah(s.perkiraan) : ""));
  } else if (s.status === "Diambil") {
    tr.appendChild(tdDuaBaris("Lunas " + formatRupiah(s.total),
      s.diambilAt ? "Diambil " + formatTanggal(isoTanggal(new Date(s.diambilAt))) : ""));
  } else {
    tr.appendChild(tdDuaBaris(s.total > 0 ? "Biaya cek " + formatRupiah(s.total) : "Tanpa biaya", s.alasanBatal));
  }

  /* Aksi */
  const tdAksi = document.createElement("td");
  const wadah = document.createElement("div");
  wadah.className = "aksi-service";
  if (serviceAktif(s)) {
    wadah.appendChild(buatTombol("Lunasi", "btn-aksi", () => bukaPelunasan(s.id)));
    wadah.appendChild(buatTombol("Batal", "btn-hapus", () => bukaPembatalan(s.id)));
  } else if (sesi && sesi.role === "admin") {
    wadah.appendChild(buatTombol("Buka Lagi", "btn-aksi btn-aksi-abu", () => bukaLagiService(s.id)));
  } else {
    const ket = document.createElement("span");
    ket.className = "teks-aksi";
    ket.textContent = "-";
    wadah.appendChild(ket);
  }
  tdAksi.appendChild(wadah);
  tr.appendChild(tdAksi);
  return tr;
}

async function ubahStatusService(id, status) {
  const s = cariService(id);
  if (!s || s.status === status) return;
  const lama = s.status;
  s.status = status; // tampil langsung, database menyusul
  renderService();
  const { error } = await kirim(db.rpc("service_ubah_status", { p_id: id, p_status: status }));
  if (error) {
    s.status = lama;
    renderService();
    alert("Gagal mengubah status " + s.nota + ": " + teksGagal(error));
  }
  jadwalMuatUlang();
}

async function simpanServiceMasuk(e) {
  e.preventDefault();
  if (sedangKirimService) return; // tombol diklik dua kali: jangan tersimpan dobel
  const isi = {
    p_pelanggan: $("srv-pelanggan").value.trim(),
    p_no_hp: $("srv-hp").value.trim(),
    p_perangkat: $("srv-perangkat").value.trim(),
    p_kelengkapan: $("srv-kelengkapan").value.trim(),
    p_keluhan: $("srv-keluhan").value.trim(),
    p_perkiraan: angkaRupiah($("srv-perkiraan")) || 0,
    p_dp: angkaRupiah($("srv-dp")) || 0,
  };
  if (!isi.p_pelanggan || !isi.p_perangkat || !isi.p_keluhan) {
    return pesanService("srv-pesan", "Nama pelanggan, perangkat, dan keluhan wajib diisi.", true);
  }
  if (isi.p_perkiraan > 0 && isi.p_dp > isi.p_perkiraan &&
      !confirm("DP (" + formatRupiah(isi.p_dp) + ") lebih besar dari perkiraan biaya (" +
               formatRupiah(isi.p_perkiraan) + "). Tetap simpan?")) return;

  sedangKirimService = true;
  pesanService("srv-pesan", "Menyimpan...");
  const { data, error } = await kirim(db.rpc("service_terima", isi));
  sedangKirimService = false;
  if (error) return pesanService("srv-pesan", "Gagal menyimpan: " + teksGagal(error), true);

  const s = terapkanBarisService(data);
  resetSetelahSimpanService = true; // supaya event "reset" tidak menghapus pesan
  $("form-service").reset();
  $("srv-saring").value = "aktif";
  $("srv-cari").value = "";
  renderService();
  $("srv-pelanggan").focus();
  pesanService("srv-pesan", "Service " + s.nota + " disimpan." +
    (isi.p_dp > 0 ? " DP " + formatRupiah(isi.p_dp) + " tercatat di kas." : "") +
    " Tulis nomor " + s.nota + " di nota pelanggan.");
  jadwalMuatUlang(); // ambil transaksi DP dari database
}

/* ----- Panel pelunasan & pembatalan ----- */
function tutupPanelService() {
  panelService = null;
  $("srv-panel-lunas").hidden = true;
  $("srv-panel-batal").hidden = true;
}

function tampilkanPanel(id, fokus) {
  $(id).hidden = false;
  $(id).scrollIntoView({ block: "center", behavior: "smooth" });
  $(fokus).focus({ preventScroll: true });
}

/* Isi ulang panel yang sedang terbuka (juga saat data berubah dari perangkat lain) */
function perbaruiPanelService() {
  if (!panelService) return;
  const s = cariService(panelService.id);
  if (!s || !serviceAktif(s)) {
    tutupPanelService();
    pesanService("srv-pesan", "Service itu sudah diproses dari perangkat lain.", true);
    return;
  }
  if (panelService.jenis === "lunas") renderPanelLunas(s); else renderPanelBatal(s);
}

function barisRingkasan(wadah, label, nilai, kelas) {
  const div = document.createElement("div");
  if (kelas) div.className = kelas;
  const a = document.createElement("span");
  a.className = "label";
  a.textContent = label;
  const b = document.createElement("span");
  b.textContent = nilai;
  div.appendChild(a);
  div.appendChild(b);
  wadah.appendChild(div);
}

function bukaPelunasan(id) {
  const s = cariService(id);
  if (!s || !serviceAktif(s)) return;
  panelService = { jenis: "lunas", id: id, sparepart: [] };
  $("srv-panel-batal").hidden = true;
  $("form-lunas").reset();
  $("lunas-judul").textContent = "Pelunasan " + s.nota + " - " + s.pelanggan + " (" + s.perangkat + ")";
  pesanService("lunas-pesan", "");
  perbaruiPanelService();
  tampilkanPanel("srv-panel-lunas", "lunas-jasa");
}

function isiPilihanSparepart() {
  const pilih = $("lunas-barang");
  const sebelumnya = pilih.value;
  pilih.innerHTML = "";
  const kosong = document.createElement("option");
  kosong.value = "";
  kosong.textContent = dataStok.length ? "(Pilih sparepart)" : "(Stok masih kosong)";
  pilih.appendChild(kosong);
  dataStok.forEach((b) => {
    const o = document.createElement("option");
    o.value = String(b.id);
    o.textContent = b.nama + " (stok " + b.stok + ") - " + formatRupiah(b.hargaJual);
    if (b.stok <= 0) o.disabled = true;
    pilih.appendChild(o);
  });
  pilih.value = sebelumnya;
  if (pilih.value !== sebelumnya) pilih.value = "";
}

function tambahSparepart() {
  if (!panelService || panelService.jenis !== "lunas") return;
  const b = dataStok.find((x) => String(x.id) === $("lunas-barang").value);
  const qty = Number($("lunas-qty").value);
  if (!b) return pesanService("lunas-pesan", "Pilih sparepart dulu.", true);
  if (!Number.isInteger(qty) || qty < 1) return pesanService("lunas-pesan", "Jumlah sparepart minimal 1.", true);
  const ada = panelService.sparepart.find((x) => x.stokId === b.id);
  const jumlah = qty + (ada ? ada.qty : 0);
  if (jumlah > b.stok) return pesanService("lunas-pesan", "Stok " + b.nama + " tidak cukup (sisa " + b.stok + ").", true);
  if (ada) ada.qty = jumlah; else panelService.sparepart.push({ stokId: b.id, qty: qty });
  $("lunas-barang").value = "";
  $("lunas-qty").value = "1";
  pesanService("lunas-pesan", "");
  perbaruiPanelService();
}

function hapusSparepart(stokId) {
  if (!panelService || panelService.jenis !== "lunas") return;
  panelService.sparepart = panelService.sparepart.filter((x) => x.stokId !== stokId);
  perbaruiPanelService();
}

function hitungPelunasan(s) {
  const jasa = angkaRupiah($("lunas-jasa")) || 0;
  const sparepart = panelService.sparepart.reduce((n, x) => {
    const b = dataStok.find((y) => y.id === x.stokId);
    return n + (b ? b.hargaJual * x.qty : 0);
  }, 0);
  const total = jasa + sparepart;
  return { jasa: jasa, sparepart: sparepart, total: total, sisa: total - s.dibayar };
}

function renderPanelLunas(s) {
  isiPilihanSparepart();
  /* barang yang sudah dihapus dari stok dikeluarkan dari daftar */
  panelService.sparepart = panelService.sparepart.filter((x) => dataStok.some((b) => b.id === x.stokId));

  const tbody = $("lunas-sparepart");
  tbody.innerHTML = "";
  panelService.sparepart.forEach((x) => {
    const b = dataStok.find((y) => y.id === x.stokId);
    const tr = document.createElement("tr");
    tr.appendChild(td(b.nama));
    tr.appendChild(td(String(x.qty)));
    tr.appendChild(td(formatRupiah(b.hargaJual * x.qty)));
    const tdAksi = document.createElement("td");
    tdAksi.appendChild(buatTombol("Hapus", "btn-hapus", () => hapusSparepart(x.stokId)));
    tr.appendChild(tdAksi);
    tbody.appendChild(tr);
  });
  $("lunas-tabel").hidden = panelService.sparepart.length === 0;

  const h = hitungPelunasan(s);
  const r = $("lunas-ringkasan");
  r.innerHTML = "";
  if (s.perkiraan > 0) barisRingkasan(r, "Perkiraan awal", formatRupiah(s.perkiraan), "catatan");
  barisRingkasan(r, "Biaya jasa", formatRupiah(h.jasa));
  barisRingkasan(r, "Sparepart", formatRupiah(h.sparepart));
  barisRingkasan(r, "Total biaya", formatRupiah(h.total), "total");
  barisRingkasan(r, "Sudah dibayar (DP)", formatRupiah(s.dibayar));
  if (h.sisa >= 0) barisRingkasan(r, "Diterima dari pelanggan", formatRupiah(h.sisa), "total");
  else barisRingkasan(r, "Kembalikan ke pelanggan", formatRupiah(-h.sisa), "total");
}

async function simpanPelunasan(e) {
  e.preventDefault();
  if (!panelService || panelService.jenis !== "lunas" || sedangProsesPanel) return;
  const s = cariService(panelService.id);
  if (!s) return;
  const jasa = angkaRupiah($("lunas-jasa"));
  if (!(jasa >= 0)) return pesanService("lunas-pesan", "Isi biaya jasa (ketik 0 kalau gratis).", true);
  const h = hitungPelunasan(s);
  if (!confirm("Simpan pelunasan " + s.nota + " - " + s.pelanggan + "?\n\n" +
      "Total biaya: " + formatRupiah(h.total) + "\n" +
      (h.sisa >= 0 ? "Terima dari pelanggan: " + formatRupiah(h.sisa) : "Kembalikan ke pelanggan: " + formatRupiah(-h.sisa)) +
      "\n\nBarang ditandai sudah diambil.")) return;

  sedangProsesPanel = true;
  pesanService("lunas-pesan", "Menyimpan...");
  const { data, error } = await kirim(db.rpc("service_lunasi", {
    p_id: s.id,
    p_biaya_jasa: jasa,
    p_sparepart: panelService.sparepart.map((x) => ({ stok_id: x.stokId, qty: x.qty })),
  }));
  sedangProsesPanel = false;
  if (error) {
    if (/tidak cukup/i.test(String(error.message))) {
      try { await muatData(); } catch (err) { /* abaikan */ }
      perbaruiPanelService();
    }
    return pesanService("lunas-pesan", "Gagal: " + teksGagal(error), true);
  }

  terapkanBarisService(data);
  tutupPanelService();
  try { await muatData(); } catch (err) { /* tampilkan data terakhir */ }
  renderService();
  pesanService("srv-pesan", s.nota + " lunas dan sudah diserahkan. " +
    (h.sisa >= 0 ? "Diterima " + formatRupiah(h.sisa) : "Dikembalikan " + formatRupiah(-h.sisa)) + " (tercatat di kas).");
  jadwalMuatUlang();
}

function bukaPembatalan(id) {
  const s = cariService(id);
  if (!s || !serviceAktif(s)) return;
  panelService = { jenis: "batal", id: id };
  $("srv-panel-lunas").hidden = true;
  $("form-batal").reset();
  $("batal-judul").textContent = "Batalkan " + s.nota + " - " + s.pelanggan + " (" + s.perangkat + ")";
  pesanService("batal-pesan", "");
  perbaruiPanelService();
  tampilkanPanel("srv-panel-batal", "batal-cek");
}

function renderPanelBatal(s) {
  const cek = angkaRupiah($("batal-cek")) || 0;
  const r = $("batal-ringkasan");
  r.innerHTML = "";
  barisRingkasan(r, "Sudah dibayar (DP)", formatRupiah(s.dibayar));
  barisRingkasan(r, "Biaya cek", formatRupiah(cek));
  if (s.dibayar > cek) barisRingkasan(r, "Kembalikan ke pelanggan", formatRupiah(s.dibayar - cek), "total");
  else if (cek > s.dibayar) barisRingkasan(r, "Diterima dari pelanggan", formatRupiah(cek - s.dibayar), "total");
  else barisRingkasan(r, "Uang keluar/masuk", formatRupiah(0), "total");
}

async function simpanPembatalan(e) {
  e.preventDefault();
  if (!panelService || panelService.jenis !== "batal" || sedangProsesPanel) return;
  const s = cariService(panelService.id);
  if (!s) return;
  const cek = angkaRupiah($("batal-cek")) || 0;
  const selisih = s.dibayar - cek;
  if (!confirm("Batalkan " + s.nota + " - " + s.pelanggan + "?\n\n" +
      (selisih > 0 ? "Kembalikan ke pelanggan: " + formatRupiah(selisih)
        : selisih < 0 ? "Terima biaya cek: " + formatRupiah(-selisih)
        : "Tidak ada uang yang perlu dikembalikan.") +
      (sesi && sesi.role === "admin" ? "" : "\n\nSetelah dibatalkan, hanya admin yang bisa membukanya lagi."))) return;

  sedangProsesPanel = true;
  pesanService("batal-pesan", "Menyimpan...");
  const { data, error } = await kirim(db.rpc("service_batal", {
    p_id: s.id, p_biaya_cek: cek, p_alasan: $("batal-alasan").value.trim(),
  }));
  sedangProsesPanel = false;
  if (error) return pesanService("batal-pesan", "Gagal: " + teksGagal(error), true);

  terapkanBarisService(data);
  tutupPanelService();
  try { await muatData(); } catch (err) { /* tampilkan data terakhir */ }
  renderService();
  pesanService("srv-pesan", s.nota + " dibatalkan." +
    (selisih > 0 ? " DP " + formatRupiah(selisih) + " dikembalikan (tercatat di kas)."
      : selisih < 0 ? " Biaya cek " + formatRupiah(-selisih) + " tercatat di kas." : ""));
  jadwalMuatUlang();
}

/* Admin: batalkan pelunasan/pembatalan untuk koreksi salah input */
async function bukaLagiService(id) {
  const s = cariService(id);
  if (!s || serviceAktif(s) || sedangDihapus.has(id)) return;
  if (!confirm("Buka lagi " + s.nota + " - " + s.pelanggan + "?\n\n" +
      (s.status === "Diambil"
        ? "Pelunasan dibatalkan: uang pelunasan dihapus dari kas dan stok sparepart dikembalikan. Status kembali ke Selesai."
        : "Pembatalan dibatalkan: uang pengembalian DP / biaya cek dihapus dari kas. Status kembali ke Diterima."))) return;
  sedangDihapus.add(id);
  const { data, error } = await kirim(db.rpc("service_buka_lagi", { p_id: id }));
  sedangDihapus.delete(id);
  if (error) return alert("Gagal membuka lagi " + s.nota + ": " + teksGagal(error));
  terapkanBarisService(data);
  try { await muatData(); } catch (err) { /* tampilkan data terakhir */ }
  renderService();
  pesanService("srv-pesan", s.nota + " dibuka lagi.");
  jadwalMuatUlang();
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

/* Kolom uang (Nominal, Harga Beli, Harga Jual): format Rupiah otomatis */
document.querySelectorAll(".input-rupiah").forEach((el) => {
  el.addEventListener("input", () => formatKetikRupiah(el));
});

/* Service */
$("form-service").addEventListener("submit", simpanServiceMasuk);
$("form-service").addEventListener("reset", () => {
  setTimeout(() => {
    if (resetSetelahSimpanService) { resetSetelahSimpanService = false; return; }
    pesanService("srv-pesan", "");
  }, 0);
});
$("srv-cari").addEventListener("input", renderService);
$("srv-saring").addEventListener("change", renderService);
$("form-lunas").addEventListener("submit", simpanPelunasan);
$("lunas-tambah").addEventListener("click", tambahSparepart);
$("lunas-qty").addEventListener("keydown", (e) => { // Enter di kolom jumlah = Tambah, bukan simpan
  if (e.key === "Enter") { e.preventDefault(); tambahSparepart(); }
});
$("lunas-jasa").addEventListener("input", perbaruiPanelService);
$("lunas-tutup").addEventListener("click", tutupPanelService);
$("form-batal").addEventListener("submit", simpanPembatalan);
$("batal-cek").addEventListener("input", perbaruiPanelService);
$("batal-tutup").addEventListener("click", tutupPanelService);

/* Stok */
$("form-stok").addEventListener("submit", tambahBarang);
$("stok-cari").addEventListener("input", renderStok);

/* Kas */
$("form-kas").addEventListener("submit", simpanKas);
$("form-kas").addEventListener("reset", () => {
  setTimeout(() => {
    if (sedangSimpan) { sedangSimpan = false; return; }
    setTanggalHariIni();
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
$("kas-tanggal").addEventListener("change", () => {
  tanggalIkutHariIni = $("kas-tanggal").value === hariIni();
  renderKas();
  pesanKas("");
});
$("kas-tanggal").addEventListener("blur", () => {
  if ($("kas-tanggal").value) return;
  setTanggalHariIni(); // tanggal dikosongkan lalu ditinggal: kembali ke hari ini
  renderKas();
});

/* Lewat tengah malam; kembali ke tab ini: ambil data terbaru */
setInterval(cekGantiHari, 30000);
document.addEventListener("visibilitychange", () => {
  if (document.hidden) return;
  cekGantiHari();
  if (sesi) jadwalMuatUlang();
});

/* Sesi login habis / dicabut di Supabase: kembali ke halaman login */
if (db) {
  db.auth.onAuthStateChange((event) => {
    if (event !== "SIGNED_OUT" || !sesi) return;
    setTimeout(() => {
      akhiriSesi();
      tampilkanLogin("Sesi login berakhir. Silakan masuk lagi.");
    }, 0);
  });
}

/* Buka website: lanjutkan sesi login yang masih ada di tab ini */
async function mulai() {
  if (!db) return tampilkanLogin("Gagal memuat Supabase. Periksa koneksi internet lalu muat ulang halaman.");
  try {
    const { data } = await db.auth.getSession();
    if (!data.session) return tampilkanLogin();
    const profil = await ambilProfil();
    if (!profil) {
      await db.auth.signOut({ scope: "local" });
      return tampilkanLogin();
    }
    await muatData();
    sesi = profil;
    tampilkanAplikasi();
  } catch (e) {
    tampilkanLogin(teksGagal(e));
  }
}

mulai();