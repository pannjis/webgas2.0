function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('SiGAS PRO - Sistem Agen Gas')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- SETUP DATABASE OTOMATIS ---
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = [
    {name: 'USERS', header: ['Username', 'Password', 'Role', 'Nama']},
    {name: 'PRODUK', header: ['ID', 'Nama_Produk', 'Harga_Jual', 'Harga_Beli', 'Stok_Isi', 'Stok_Kosong', 'SKU', 'Kode', 'Link_Gambar']},
    {name: 'PELANGGAN', header: ['ID', 'Nama', 'NoHP', 'Alamat']},
    {name: 'SUPPLIER', header: ['ID', 'Nama_Supplier', 'NoHP', 'Alamat']},
    {name: 'TRANSAKSI', header: ['ID_Trans', 'Waktu', 'Pelanggan', 'Produk', 'Qty', 'Total', 'Tipe', 'Kasir', 'Metode_Bayar', 'Jatuh_Tempo', 'Status']},
    {name: 'PEMBELIAN', header: ['ID_Beli', 'Waktu', 'Supplier', 'Produk', 'Qty', 'Total', 'Metode']},
    {name: 'KEUANGAN', header: ['ID', 'Tanggal', 'Jenis', 'Kategori', 'Nominal', 'Keterangan']},
    {name: 'KATEGORI', header: ['Nama_Kategori']},
    {name: 'KARYAWAN', header: ['ID', 'Nama', 'NoHP', 'Gaji_Pokok', 'Bonus_Per_Pcs', 'Status']}, 
    {name: 'KASBON', header: ['ID_Kasbon', 'Tanggal', 'Nama_Karyawan', 'Nominal', 'Keterangan', 'Status_Lunas']}
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.header);
      // Data Dummy Awal
      if(s.name === 'USERS') sheet.appendRow(['admin', 'admin123', 'Admin', 'Super Admin']);
      if(s.name === 'KATEGORI') {
        sheet.appendRow(['Listrik & Air']);
        sheet.appendRow(['Operasional Toko']);
        sheet.appendRow(['Konsumsi']);
      }
    }
  });
}

// --- HELPER DATA ---
function getData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

// --- LOGIN ---
function loginUser(username, password) {
  const data = getData('USERS');
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { status: 'success', role: data[i][2], nama: data[i][3] };
    }
  }
  return { status: 'failed' };
}

// --- DASHBOARD ---
function getDashboardStats() {
  const keu = getData('KEUANGAN');
  let income = 0, expense = 0;
  
  keu.forEach(r => {
    if(r[2] === 'Pemasukan') income += Number(r[4]);
    if(r[2] === 'Pengeluaran') expense += Number(r[4]);
  });
  
  return { income, expense, net: income - expense };
}

// [UPDATE] Fungsi Tambah Produk (Versi Debugging)
// [UPDATE] Fungsi Tambah Produk (Upload ke Folder Khusus)
function tambahProduk(form) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUK');
  
  // ID Folder Google Drive Anda
  const FOLDER_ID = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 
  
  let imageUrl = '';

  // PROSES UPLOAD
  if (form.gambar && form.gambar.data) {
    try {
      const decoded = Utilities.base64Decode(form.gambar.data);
      const blob = Utilities.newBlob(decoded, form.gambar.mimeType, form.gambar.fileName);
      
      // 1. Ambil Folder Tujuan
      const folder = DriveApp.getFolderById(FOLDER_ID);
      
      // 2. Simpan File di Folder Tersebut
      const file = folder.createFile(blob); 
      
      // 3. Set Permission (Coba Publik -> Domain -> Private)
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e1) {
        try {
           file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e2) {
           console.log("Gagal set permission: " + e1.message); 
        }
      }

      // 4. Ambil Link
      // Ganti format link jadi Thumbnail (agar tidak crash/broken di browser)
      imageUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

    } catch (e) {
      // Tampilkan error detail jika gagal
      throw new Error("Gagal Upload: " + e.message); 
    }
  } else {
    // Jika manual link
    imageUrl = (typeof form.gambar === 'string') ? form.gambar : '';
  }

  // Simpan ke Spreadsheet
  sheet.appendRow([
    'P-' + Date.now(), 
    form.nama, 
    form.hargaJual, 
    form.hargaBeli, 
    form.stokIsi, 
    form.stokKosong,
    form.sku,     
    form.kode,    
    imageUrl 
  ]);
}

function hapusProduk(nama) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUK');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == nama) { sheet.deleteRow(i + 1); break; }
  }
}

// --- MODIFIKASI: TRANSAKSI & KASIR ---

// 1. Simpan Transaksi (BULK / BANYAK ITEM SEKALIGUS)
function simpanTransaksiBulk(dataTransaksi) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  const prodData = prodSheet.getDataRange().getValues();
  const idTrxMaster = 'TRX-' + Date.now(); 
  const waktu = new Date();
  let totalBelanja = 0;
  let summaryProduk = [];

  // [BAGIAN INI YANG TADI HILANG]
  // Kita tentukan statusnya SEKALI saja di sini
  let statusTrx = (dataTransaksi.metode === 'Hutang') ? 'Belum Lunas' : 'Lunas';

  // Loop setiap item di keranjang
  dataTransaksi.items.forEach(item => {
    let itemFound = false;
    
    // Update Stok
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1] == item.produkNama) {
        let curIsi = Number(prodData[i][4]);
        let curKosong = Number(prodData[i][5]);
        
        // Validasi Stok
        if (curIsi < item.qty) throw new Error(`Stok ${item.produkNama} Habis! Sisa: ${curIsi}`);

        // Update logic
        let newIsi = curIsi - item.qty;
        let newKosong = curKosong;
        
        if (item.tipe === 'Tukar (Refill)') {
           newKosong = curKosong + Number(item.qty); 
        }
        
        prodSheet.getRange(i + 1, 5).setValue(newIsi);
        prodSheet.getRange(i + 1, 6).setValue(newKosong);
        itemFound = true;
        break;
      }
    }
    
    if(!itemFound) throw new Error(`Produk ${item.produkNama} tidak ditemukan di database.`);

    // Catat ke Sheet TRANSAKSI
    // Sekarang variabel 'statusTrx' sudah dikenali karena sudah dibuat di atas loop
    trxSheet.appendRow([
      idTrxMaster, 
      waktu, 
      dataTransaksi.pelanggan, 
      item.produkNama, 
      item.qty, 
      item.total, 
      item.tipe, 
      dataTransaksi.kasir, 
      dataTransaksi.metode, 
      dataTransaksi.jatuhTempo, 
      statusTrx 
    ]);

    totalBelanja += Number(item.total);
    summaryProduk.push(`${item.produkNama} (${item.qty})`);
  });

  // LOGIKA KEUANGAN (Hanya catat jika BUKAN Hutang)
  if (dataTransaksi.metode !== 'Hutang') {
      keuSheet.appendRow([
        'FIN-' + idTrxMaster, waktu, 'Pemasukan', 'Penjualan Gas', 
        totalBelanja, `Penjualan: ${summaryProduk.join(', ')} (${dataTransaksi.metode})`
      ]);
  }
  
  return "Transaksi Berhasil Disimpan!";
}

// --- FITUR PIUTANG (BACKEND) ---

// --- FITUR PIUTANG (VERSI SMART COLUMN) ---

function getDataPiutang() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  // Ambil semua data TERMASUK Header (Judul Kolom)
  const allData = sheet.getDataRange().getValues();
  
  // Cek jika data kosong
  if (allData.length < 2) return [];

  const headers = allData[0]; // Baris pertama adalah Header
  
  // [PENTING] Cari urutan kolom secara otomatis berdasarkan Namanya
  // Ini mencegah error jika kolom tergeser
  const idxStatus = headers.indexOf('Status'); 
  const idxJatuhTempo = headers.indexOf('Jatuh_Tempo');
  const idxMetode = headers.indexOf('Metode_Bayar'); 
  
  // Jika kolom Status tidak ditemukan, hentikan (daripada error)
  if (idxStatus === -1) return [];

  let grouped = {};

  // Loop mulai dari baris ke-2 (Index 1) karena baris 0 adalah Header
  for (let i = 1; i < allData.length; i++) {
    let row = allData[i];
    let status = row[idxStatus]; // Ambil status dari kolom yang ditemukan tadi

    // Logika Filter: Hanya ambil yang "Belum Lunas"
    if (status === 'Belum Lunas') {
       let id = row[0]; // ID selalu di kolom pertama
       
       if(!grouped[id]) {
          grouped[id] = {
             id: id,
             waktu: row[1],
             pelanggan: row[2],
             total: 0,
             jatuhTempo: (idxJatuhTempo !== -1) ? row[idxJatuhTempo] : '' // Ambil tgl jika kolom ada
          };
       }
       // Jumlahkan total (Pastikan angka)
       grouped[id].total += Number(row[5]); 
    }
  }

  // Kembalikan hasil dalam bentuk Array
  return Object.values(grouped).map(x => [x.id, x.waktu, x.pelanggan, x.total, x.jatuhTempo]);
}

// 2. Proses Pelunasan
function lunasiHutang(idTrx, totalBayar, namaPelanggan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTrx = ss.getSheetByName('TRANSAKSI');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
  
  const dataTrx = sheetTrx.getDataRange().getValues();
  
  // A. Update Status di TRANSAKSI jadi 'Lunas'
  for(let i=1; i<dataTrx.length; i++) {
     if(dataTrx[i][0] == idTrx) {
        // Kolom K (Index 11, karena start dari 1 di sheet) -> Kolom ke-11
        sheetTrx.getRange(i+1, 11).setValue('Lunas'); 
     }
  }

  // B. Masukkan Uang ke KEUANGAN (Karena baru terima duit sekarang)
  sheetKeu.appendRow([
      'LUNAS-' + Date.now(), 
      new Date(), 
      'Pemasukan', 
      'Pelunasan Piutang', 
      totalBayar, 
      `Pelunasan Bon: ${namaPelanggan} (${idTrx})`
  ]);

  return "Hutang Berhasil Dilunasi & Masuk Kas!";
}

function getJumlahJatuhTempo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TRANSAKSI');
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  let count = 0;
  let uniqueIDs = []; // Supaya tidak double hitung item dalam 1 struk

  // Loop data transaksi
  for (let i = 1; i < data.length; i++) {
    let idTrx = data[i][0];
    let status = data[i][10]; // Kolom K (Status)
    let tglTempo = new Date(data[i][9]); // Kolom J (Jatuh Tempo)

    // Logika: Status Belum Lunas DAN Tanggal Tempo < Hari Ini (Sudah lewat)
    if (status === 'Belum Lunas' && tglTempo <= today && !uniqueIDs.includes(idTrx)) {
       count++;
       uniqueIDs.push(idTrx);
    }
  }
  return count;
}

// 2. Ambil Riwayat Transaksi
// --- Code.gs ---

function getRiwayatTransaksi() {
  const data = getData('TRANSAKSI'); // Ambil semua data
  
  // Objek penampung untuk pengelompokan
  let grouped = {};

  data.forEach(row => {
    let id = row[0];
    
    // Konversi Tanggal agar aman dikirim ke browser
    let waktuStr = row[1];
    if (row[1] instanceof Date) {
       waktuStr = row[1].toISOString();
    }

    // Jika ID belum ada di penampung, buat baru
    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2],
        kasir: row[7],
        totalBayar: 0,  // Nanti dijumlahkan
        items: []       // Array untuk menyimpan detail barang
      };
    }

    // Tambahkan detail item ke transaksi tersebut
    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: row[6],
      status: row[8]
    });

    // Akumulasi Total Bayar (Hanya jika status bukan Retur Full, opsional)
    grouped[id].totalBayar += Number(row[5]);
  });

  // Ubah Object menjadi Array dan urutkan dari yang terbaru (Descending)
  const result = Object.values(grouped).sort((a, b) => {
      return new Date(b.waktu) - new Date(a.waktu);
  });

  // Ambil 50 transaksi terakhir saja agar ringan
  return result.slice(0, 50);
}

// --- Code.gs ---

// 1. GET RIWAYAT PEMBELIAN (Grouping per ID)
function getRiwayatPembelian() {
  const data = getData('PEMBELIAN');
  let grouped = {};

  data.forEach(row => {
    let id = row[0];
    let waktuStr = row[1] instanceof Date ? row[1].toISOString() : row[1];

    if (!grouped[id]) {
      grouped[id] = {
        id: id,
        waktu: waktuStr,
        pelanggan: row[2], // Di sheet PEMBELIAN kolom ini adalah Supplier
        totalBayar: 0,
        items: []
      };
    }

    // Sheet PEMBELIAN: ID, Waktu, Supplier, Produk, Qty, Total, Metode
    grouped[id].items.push({
      produk: row[3],
      qty: row[4],
      hargaTotal: row[5],
      tipe: 'Stok Masuk', // Default tipe
      status: 'Sukses' 
    });
    
    grouped[id].totalBayar += Number(row[5]);
  });

  return Object.values(grouped).sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 50);
}

// 2. FUNGSI RETUR BARU (Support Partial & Jenis Transaksi)
function prosesReturBaru(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  // Tentukan Sheet Target berdasarkan jenis
  const targetSheetName = payload.jenis === 'JUAL' ? 'TRANSAKSI' : 'PEMBELIAN';
  const trxSheet = ss.getSheetByName(targetSheetName);
  const trxData = trxSheet.getDataRange().getValues();
  const prodData = prodSheet.getDataRange().getValues();

  let totalRefund = 0;
  let logItem = [];

  // Loop item yang diretur
  payload.items.forEach(returItem => {
    if(returItem.qtyRetur > 0) {
      
      // A. UPDATE STOK PRODUK
      for (let i = 1; i < prodData.length; i++) {
        if (prodData[i][1] == returItem.produk) {
           let curIsi = Number(prodData[i][4]);
           let curKosong = Number(prodData[i][5]);
           
           if(payload.jenis === 'JUAL') {
              // Retur Penjualan: Stok Isi KEMBALI (+), Stok Kosong BERKURANG (karena sebelumnya tukar)
              prodSheet.getRange(i+1, 5).setValue(curIsi + returItem.qtyRetur);
              // Cek jika itu refill, tabung kosong dikembalikan ke pelanggan (stok kita berkurang)
              if(returItem.tipe && returItem.tipe.includes('Refill')) {
                 prodSheet.getRange(i+1, 6).setValue(curKosong - returItem.qtyRetur);
              }
           } else {
              // Retur Pembelian: Stok Isi BERKURANG (-) (Balikin ke supplier)
              prodSheet.getRange(i+1, 5).setValue(curIsi - returItem.qtyRetur);
              // Jika beli tukar tabung, stok kosong kita bertambah lagi (dibalikin supplier)
               // (Sederhananya kita kurangi stok isi saja dulu untuk keamanan)
           }
           break;
        }
      }

      // B. UPDATE STATUS TRANSAKSI (Tandai Retur)
      // Cari baris transaksi spesifik
      for(let i=1; i<trxData.length; i++) {
         if(trxData[i][0] == payload.idTrx && trxData[i][3] == returItem.produk) {
             // Opsional: Bisa update kolom qty atau tambah catatan "Retur Partial"
             // Disini kita biarkan record asli, tapi catat di Keuangan sebagai pengurang
         }
      }
      
      totalRefund += (returItem.hargaSatuan * returItem.qtyRetur);
      logItem.push(`${returItem.produk} (x${returItem.qtyRetur})`);
    }
  });

  // C. CATAT DI KEUANGAN (Balance)
  if(totalRefund > 0) {
     if(payload.jenis === 'JUAL') {
        // Retur Jual = Uang Keluar (Refund ke Pelanggan)
        keuSheet.appendRow(['RET-' + Date.now(), new Date(), 'Pengeluaran', 'Retur Penjualan', totalRefund, `Retur TRX: ${payload.idTrx}. ${payload.alasan}`]);
     } else {
        // Retur Beli = Uang Masuk (Refund dari Supplier)
        keuSheet.appendRow(['RET-' + Date.now(), new Date(), 'Pemasukan', 'Retur Pembelian', totalRefund, `Retur BELI: ${payload.idTrx}. ${payload.alasan}`]);
     }
  }

  return "Retur Berhasil Diproses!";
}

function simpanPelanggan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
  
  // EDIT MODE
  if(form.id) { 
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        // Update: Nama, Perusahaan, HP, Alamat
        sheet.getRange(i+1, 2, 1, 4).setValues([[form.nama, form.pt, form.hp, form.alamat]]);
        return "Data Pelanggan Diupdate";
      }
    }
  }
  
  // BARU MODE
  sheet.appendRow(['CUST-' + Date.now(), form.nama, form.pt, form.hp, form.alamat]);
  return "Pelanggan Baru Disimpan";
}

function hapusPelanggan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PELANGGAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { 
      sheet.deleteRow(i+1); 
      return "Pelanggan Dihapus";
    }
  }
}

// Fungsi bantu untuk mengambil List Pelanggan di Kasir
function getListPelanggan() {
  return getData('PELANGGAN'); // <--- WAJIB ADA 'return'
}

// 3. Hapus / Retur Transaksi
function prosesRetur(idTrx, produkNama, qty, tipe, mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  const trxSheet = ss.getSheetByName('TRANSAKSI');
  const keuSheet = ss.getSheetByName('KEUANGAN');
  
  // A. KEMBALIKAN STOK
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == produkNama) {
       let curIsi = Number(prodData[i][4]);
       let curKosong = Number(prodData[i][5]);
       
       // Logic Retur: Kembalikan Stok Isi, Kurangi Stok Kosong (jika refill)
       prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(qty));
       
       if(tipe === 'Tukar (Refill)') {
          prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(qty));
       }
       break;
    }
  }

  // B. UPDATE STATUS TRANSAKSI & KEUANGAN
  // Cari baris transaksi
  const trxData = trxSheet.getDataRange().getValues();
  let nominalRefund = 0;

  for(let i=1; i<trxData.length; i++) {
    // Mencocokkan ID, Produk, dan memastikan belum diretur
    if(trxData[i][0] == idTrx && trxData[i][3] == produkNama && trxData[i][8] != 'Retur') {
       if(mode === 'FULL') {
         trxSheet.deleteRow(i+1); // Hapus baris permanen jika mau bersih
         // Atau tandai: trxSheet.getRange(i+1, 9).setValue('Retur');
       } else {
         trxSheet.getRange(i+1, 9).setValue('Retur Item');
       }
       nominalRefund = trxData[i][5]; // Ambil total harga item tsb
       break;
    }
  }

  // C. CATAT PENGELUARAN REFUND DI KEUANGAN (Agar Balance)
  keuSheet.appendRow([
      'REFUND-' + Date.now(), new Date(), 
      'Pengeluaran', 'Retur Penjualan', 
      nominalRefund, `Retur: ${produkNama} (${idTrx})`
  ]);

  return "Berhasil Retur/Hapus";
}

// --- TAMBAHAN: SIMPAN PEMBELIAN BULK (KERANJANG) ---
function simpanPembelianBulk(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBeli = ss.getSheetByName('PEMBELIAN');
  const sheetProd = ss.getSheetByName('PRODUK');
  const sheetKeu = ss.getSheetByName('KEUANGAN');
  
  const idBeliMaster = 'BELI-' + Date.now();
  const waktu = new Date();
  const prodData = sheetProd.getDataRange().getValues();
  
  let summaryItem = [];

  // Loop setiap item di keranjang beli
  data.items.forEach(item => {
    // 1. Catat di Sheet PEMBELIAN
    // Format: ID, Waktu, Supplier, Produk, Qty, Total, Metode
    sheetBeli.appendRow([
      idBeliMaster, 
      waktu, 
      data.supplier, 
      item.produk, 
      item.qty, 
      item.total, 
      'Tunai'
    ]);

    // 2. Update Stok di Sheet PRODUK
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1] == item.produk) {
        let curIsi = Number(prodData[i][4]);
        let curKosong = Number(prodData[i][5]);
        
        // Stok Isi Bertambah (+)
        sheetProd.getRange(i + 1, 5).setValue(curIsi + Number(item.qty));
        
        // Jika Tukar Tabung, Stok Kosong Berkurang (-)
        if(item.isTukar) {
           sheetProd.getRange(i + 1, 6).setValue(curKosong - Number(item.qty));
        }
        break;
      }
    }
    summaryItem.push(`${item.produk} (x${item.qty})`);
  });

  // 3. Catat di KEUANGAN (Satu baris total pengeluaran)
  sheetKeu.appendRow([
    'OUT-' + Date.now(), 
    waktu, 
    'Pengeluaran', 
    'Pembelian Stok', 
    data.grandTotal, 
    `Beli Stok: ${summaryItem.join(', ')}`
  ]);

  return "Stok Berhasil Ditambahkan!";
}

// --- PEMBELIAN (BELI) ---
function tambahSupplier(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SUPPLIER').appendRow(['SUP-' + Date.now(), form.nama, form.hp, form.alamat]);
}

function simpanPembelian(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodSheet = ss.getSheetByName('PRODUK');
  
  // 1. Catat Beli
  ss.getSheetByName('PEMBELIAN').appendRow(['BELI-' + Date.now(), new Date(), data.supplier, data.produk, data.qty, data.total, data.metode]);
  
  // 2. Update Stok
  const prodData = prodSheet.getDataRange().getValues();
  for (let i = 1; i < prodData.length; i++) {
    if (prodData[i][1] == data.produk) {
      let curIsi = Number(prodData[i][4]);
      let curKosong = Number(prodData[i][5]);
      
      prodSheet.getRange(i + 1, 5).setValue(curIsi + Number(data.qty)); // Stok Isi Nambah
      if(data.isTukar) {
        prodSheet.getRange(i + 1, 6).setValue(curKosong - Number(data.qty)); // Stok Kosong Berkurang
      }
      break;
    }
  }
  
  // 3. Catat Pengeluaran
  ss.getSheetByName('KEUANGAN').appendRow(['OUT-' + Date.now(), new Date(), 'Pengeluaran', 'Pembelian Stok', data.total, `Beli ${data.produk}`]);
}

// --- KEUANGAN ---
function getKategori() {
  return getData('KATEGORI').map(r => r[0]);
}

function tambahKategori(nama) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KATEGORI').appendRow([nama]);
}

function simpanKeuangan(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KEUANGAN')
    .appendRow(['MANUAL-' + Date.now(), new Date(), form.jenis, form.kategori, form.nominal, form.keterangan]);
}

// --- SDM: KARYAWAN ---
function simpanKaryawan(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KARYAWAN');
  
  if(form.id) { // Edit Mode
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
      if(data[i][0] == form.id) {
        sheet.getRange(i+1, 2, 1, 4).setValues([[form.nama, form.hp, form.gaji, form.bonus]]);
        return "Data Updated";
      }
    }
  } 
  // New Mode
  sheet.appendRow(['KRY-' + Date.now(), form.nama, form.hp, form.gaji, form.bonus, 'Aktif']);
  return "Karyawan Baru Disimpan";
}

function hapusKaryawan(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KARYAWAN');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][0] == id) { sheet.deleteRow(i+1); return; }
  }
}

// --- SDM: KASBON ---
function simpanKasbon(form) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KASBON')
    .appendRow(['KSB-' + Date.now(), new Date(), form.nama, form.nominal, form.ket, 'Belum Lunas']);
  return "Kasbon Dicatat";
}

// --- SDM: PAYROLL LOGIC ---
function getDataPayroll() {
  const karyawan = getData('KARYAWAN');
  const kasbonData = getData('KASBON');
  
  let result = karyawan.map(k => {
    let nama = k[1];
    let gaji = Number(k[3]);
    let bonusSet = Number(k[4]);
    
    // Hitung Kasbon Belum Lunas
    let totalKasbon = 0;
    kasbonData.forEach(ksb => {
      if(ksb[2] === nama && ksb[5] === 'Belum Lunas') {
        totalKasbon += Number(ksb[3]);
      }
    });
    
    // Bonus Sementara (Dummy: 0), nanti bisa dikembangkan hitung jumlah penjualan kasir
    let totalBonus = 0; 

    return {
      id: k[0],
      nama: nama,
      gaji: gaji,
      bonus: totalBonus,
      kasbon: totalKasbon,
      total: gaji + totalBonus - totalKasbon
    };
  });
  return result;
}

function prosesPayrollFinal(listGaji) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keuSheet = ss.getSheetByName('KEUANGAN');
  const kasbonSheet = ss.getSheetByName('KASBON');
  const kasbonData = kasbonSheet.getDataRange().getValues();
  
  let totalKeluar = 0;
  
  listGaji.forEach(g => {
    totalKeluar += Number(g.total);
    // Lunaskan Kasbon
    if(g.kasbon > 0) {
      for(let i=1; i<kasbonData.length; i++) {
        if(kasbonData[i][2] == g.nama && kasbonData[i][5] == 'Belum Lunas') {
          kasbonSheet.getRange(i+1, 6).setValue('Lunas (Potong Gaji)');
        }
      }
    }
  });
  
  keuSheet.appendRow(['PAY-' + Date.now(), new Date(), 'Pengeluaran', 'Gaji Karyawan', totalKeluar, 'Payroll Periode Ini']);
  return "Gaji Dicairkan & Kasbon Terpotong.";
}

function TES_BIKIN_FILE() {
  // ID Folder Anda
  const id = '15hiLtvusofF2OJpXVq8lJkePbmqVIuPM'; 
  
  const folder = DriveApp.getFolderById(id);
  
  // Kita coba bikin file teks kosong beneran untuk mancing izin "Write"
  folder.createFile('Tes_Izin.txt', 'Halo, ini tes izin upload.');
  
  console.log("Sukses! Izin Upload sudah aktif.");
}
