function syncCustomerFromOrder() {
  const customerFile = SpreadsheetApp.openById("1Gl1aqwmy3EtgByzFZ3kcXMdiHYFYNAO22j_IL4QqXWY");
  const orderFile = SpreadsheetApp.openById("17vOP4wK4_CxyjmlXH3oRA633KLjsR_FNptW7ez7gK-s");
  const produkFile = SpreadsheetApp.openById("1SfLQhBGdmOGr7JoXBR2p_vPJhZh7hz4iwjEtT2iSex0");

  const customerSheet = customerFile.getSheetByName("Db_Customer");
  const orderSheet = orderFile.getSheetByName("Db_Order");
  const produkSheet = produkFile.getSheetByName("Db_Produk");

  const customerData = customerSheet.getDataRange().getValues();
  const orderData = orderSheet.getDataRange().getValues();
  const produkData = produkSheet.getDataRange().getValues();

  const headerCustomer = customerData[0];
  const headerOrder = orderData[0];
  const headerProduk = produkData[0];

  const idx = name => headerCustomer.indexOf(name);
  const today = new Date();

  // Index kolom di Db_Customer
  const idxMap = {
    idCustomer: idx("ID_Customer"),
    jumlahOrder: idx("Jumlah_Order"),
    produkTerakhir: idx("Produk_Terakhir"),
    topik: idx("Topik"),
    lastTrans: idx("Last_Transaction"),
    keyFirst: idx("Keyname_First_Order"),
    keyLast: idx("Keyname_Terakhir_Order"),
    funnel: idx("Status_Funnel"),
    lt: idx("Status_LT")
  };

  // ================================
  // 1. MAPPING PRODUK (AMAN UNTUK KOLUMN BARU)
  // ================================
  const headerProdukLower = headerProduk.map(h => String(h).toLowerCase());

  const idxIdProduk = headerProdukLower.indexOf("id_produk");
  const idxTopik = headerProdukLower.indexOf("topik"); // otomatis cocokkan “topik”, bukan “Topik”

  const produkMap = {};
  for (let i = 1; i < produkData.length; i++) {
    const row = produkData[i];
    const idProduk = row[idxIdProduk];
    const topikProduk = row[idxTopik] || "";
    produkMap[idProduk] = topikProduk;
  }

  // ================================
  // 2. GROUP ORDER BY ID CUSTOMER
  // ================================
  const orderMap = {};
  for (let i = 1; i < orderData.length; i++) {
    const row = orderData[i];
    const [idOrder, keyname, idCustomer, idProduk, tsOrder, tglClosing] = row;

    const tglClosingDate = new Date(tglClosing);
    if (!idCustomer || isNaN(tglClosingDate)) continue;

    if (!orderMap[idCustomer]) orderMap[idCustomer] = [];
    orderMap[idCustomer].push({
      tglClosing: tglClosingDate,
      keyname,
      idProduk
    });
  }

  // Fungsi bantu hitung selisih bulan
  function getMonthDiff(d1, d2) {
    const years = d2.getFullYear() - d1.getFullYear();
    const months = d2.getMonth() - d1.getMonth();
    return years * 12 + months;
  }

  // ================================
  // 3. PROSES CUSTOMER
  // ================================
  for (let i = 1; i < customerData.length; i++) {
    const row = customerData[i];
    const idCustomer = row[idxMap.idCustomer];
    const orders = orderMap[idCustomer] || [];

    // Jika tidak ada order → Prospek
    if (orders.length === 0) {
      row[idxMap.funnel] = "Prospek";
      row[idxMap.jumlahOrder] = 0;
      row[idxMap.produkTerakhir] = "";
      row[idxMap.topik] = "";
      row[idxMap.lastTrans] = "";
      row[idxMap.keyFirst] = "";
      row[idxMap.keyLast] = "";
      row[idxMap.lt] = "";
      continue;
    }

    // Hitung transaksi
    const tglUnik = [...new Set(orders.map(o => o.tglClosing.toDateString()))];
    const orderCount = orders.length;
    const transCount = tglUnik.length;

    // Sorting transaksi berdasar tanggal
    orders.sort((a, b) => a.tglClosing - b.tglClosing);
    const first = orders[0];
    const last = orders[orders.length - 1];

    // Ambil topik dari produkMap
    const produkTerakhir = last.idProduk;
    const topik = produkMap[produkTerakhir] || "";

    // Update data customer
    row[idxMap.funnel] = transCount === 1 ? "Buyer" : `RO${transCount - 1}`;
    row[idxMap.jumlahOrder] = orderCount;
    row[idxMap.produkTerakhir] = produkTerakhir;
    row[idxMap.topik] = topik;
    row[idxMap.lastTrans] = last.tglClosing;
    row[idxMap.keyFirst] = first.keyname;
    row[idxMap.keyLast] = last.keyname;

    // Status_LT
    const selisihBulan = getMonthDiff(last.tglClosing, today);
    if (selisihBulan < 12) {
      row[idxMap.lt] = `LT${selisihBulan + 1}`;
    } else {
      row[idxMap.lt] = `LTY${Math.floor(selisihBulan / 12)}`;
    }
  }

  // ================================
  // 4. SIMPAN HASIL KE SHEET
  // ================================
  customerSheet
    .getRange(2, 1, customerData.length - 1, customerData[0].length)
    .setValues(customerData.slice(1));
}
