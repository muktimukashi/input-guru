/* Excel support uses vendored SheetJS CE 0.20.3 (see vendor/SheetJS-LICENSE.txt). */
const excelHeaders = ["Nama Barang", "Jumlah", "Satuan/Unit", "Kondisi", "Keterangan"];
const excelMaxRows = 1000;

function excelStatus(message, error = false) {
  const status = document.getElementById("excelStatus");
  status.textContent = message;
  status.style.color = error ? "#b91c1c" : "#166534";
}

function createInventoryTemplate() {
  const book = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet([excelHeaders]);
  sheet["!cols"] = [32, 14, 18, 24, 50].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(book, sheet, "Inventaris");
  const instructions = [
    ["PETUNJUK PENGISIAN INVENTARIS"],
    ["1. Isi sheet Inventaris mulai baris 2. Jangan ubah nama sheet atau judul kolom."],
    ["2. Satu file untuk satu ruangan. Pilih unit dan ruangan serta isi PIC di formulir web."],
    ["3. Nama Barang, Jumlah, Satuan/Unit, dan Kondisi wajib diisi. Keterangan opsional."],
    ["4. Jumlah berupa bilangan bulat minimal 1. Gunakan nilai langsung, bukan rumus."],
    ["5. Nama barang harus unik, termasuk terhadap barang yang sudah ada di daftar web."],
    ["6. Maksimal 1.000 barang dan ukuran file 5 MB. Baris kosong dilewati."],
    ["7. Upload menambahkan barang ke daftar. Periksa hasil lalu klik Simpan Pemeriksaan."],
    ["8. Jika ada kesalahan, seluruh upload ditolak. Perbaiki baris yang disebutkan dan upload ulang."],
    ["Satuan yang disarankan", itemUnitRecommendations.join(", ")],
    ["Kondisi yang disarankan", conditionRecommendations.join(", ")],
    ["Satuan dan kondisi lain boleh diisi seperti pada formulir manual."],
    [], ["CONTOH (salin ke sheet Inventaris hanya jika sesuai data Anda)"],
    excelHeaders,
    ["Meja Guru", 1, "Unit", "Baik", ""],
    ["Kursi Siswa", 20, "Buah", "Rusak Ringan", "Perlu perbaikan sandaran"]
  ];
  const guide = XLSX.utils.aoa_to_sheet(instructions);
  guide["!cols"] = [100, 60, 18, 24, 40].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(book, guide, "Petunjuk");
  return book;
}

function downloadInventoryTemplate() {
  try {
    if (!window.XLSX) throw new Error("Fitur Excel belum berhasil dimuat. Muat ulang halaman lalu coba lagi.");
    XLSX.writeFile(createInventoryTemplate(), "Template-Inventaris.xlsx");
    excelStatus("Template diunduh. Isi sheet Inventaris; contoh dan aturan tersedia di sheet Petunjuk.");
  } catch (error) {
    excelStatus(error.message || "Template gagal diunduh. Coba lagi.", true);
  }
}

function validateInventoryWorkbook(book, existingAssets) {
  const sheet = book.Sheets.Inventaris;
  if (!sheet || !sheet["!ref"]) throw new Error("Sheet Inventaris tidak ditemukan atau kosong. Gunakan template yang disediakan.");
  const range = XLSX.utils.decode_range(sheet["!fullref"] || sheet["!ref"]);
  if (range.e.r > 10000 || range.e.c > 50) throw new Error("Area sheet terlalu besar. Salin data ke template baru (maksimal 1.000 barang).");
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: true, range: 0 });
  if (!excelHeaders.every((header, index) => String(rows[0]?.[index] ?? "").trim() === header)) {
    throw new Error("Judul kolom tidak sesuai template. Gunakan: " + excelHeaders.join(", ") + ".");
  }
  const names = new Set(existingAssets.map(asset => asset.name.trim().toLowerCase()));
  const assets = [];
  const errors = [];
  let filledRows = 0;
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const hasFormula = Array.from({ length: range.e.c + 1 }, (_, column) => sheet[XLSX.utils.encode_cell({ r: rowIndex, c: column })]).some(cell => cell?.f);
    if (!hasFormula && row.every(value => String(value).trim() === "")) continue;
    filledRows++;
    if (filledRows > excelMaxRows) throw new Error("Maksimal 1.000 barang dalam satu file.");
    const [name, quantity, itemUnit, condition, note] = excelHeaders.map((_, index) => String(row[index] ?? "").trim());
    const issues = [];
    if (hasFormula) issues.push("gunakan nilai langsung, bukan rumus");
    if (row.slice(excelHeaders.length).some(value => String(value).trim())) issues.push("ada data di luar kolom template");
    if (!name) issues.push("Nama Barang wajib diisi");
    const qty = Number(quantity);
    if (!/^\d+$/.test(quantity) || !Number.isSafeInteger(qty) || qty < 1 || qty > 2147483647) issues.push("Jumlah harus bilangan bulat 1–2.147.483.647");
    if (!itemUnit) issues.push("Satuan/Unit wajib diisi");
    if (!condition) issues.push("Kondisi wajib diisi");
    const key = name.toLowerCase();
    if (name && names.has(key)) issues.push(`Nama Barang "${name}" duplikat di file atau daftar`);
    if (name) names.add(key);
    if (issues.length) {
      errors.push(`Baris ${rowIndex + 1}: ${issues.join("; ")}.`);
      continue;
    }
    assets.push({ id: null, code: `INV-${Date.now()}-${rowIndex}`, name, registeredQty: qty, actualQty: qty, itemUnit, exists: true, condition, note });
  }
  if (errors.length) throw new Error("Upload dibatalkan. Perbaiki data berikut:\n" + errors.slice(0, 30).join("\n") + (errors.length > 30 ? `\nDan ${errors.length - 30} baris lainnya. Perbaiki lalu upload ulang.` : ""));
  if (!assets.length) throw new Error("Sheet Inventaris belum berisi barang. Isi mulai baris 2.");
  return assets;
}

async function importInventoryExcel(input) {
  const file = input.files[0];
  if (!file) return;
  input.disabled = true;
  try {
    if (!window.XLSX) throw new Error("Fitur Excel belum berhasil dimuat. Muat ulang halaman lalu coba lagi.");
    changeRoom();
    if (!selectedRoomId) throw new Error("Pilih unit dan isi nama ruangan sebelum upload.");
    if (!/\.xlsx$/i.test(file.name)) throw new Error("Gunakan file Excel .xlsx sesuai template.");
    if (file.size > 5 * 1024 * 1024) throw new Error("Ukuran file maksimal 5 MB.");
    const targetRoomId = selectedRoomId;
    const targetUnit = document.getElementById("unitSelect").value;
    const targetName = document.getElementById("roomSelect").value.trim();
    excelStatus("Membaca dan memeriksa file Excel...");
    const buffer = await file.arrayBuffer();
    if (selectedRoomId !== targetRoomId || document.getElementById("unitSelect").value !== targetUnit || document.getElementById("roomSelect").value.trim() !== targetName) throw new Error("Ruangan berubah saat membaca file. Upload ulang untuk ruangan yang dipilih.");
    // An XLSX workbook is a ZIP container; reject other formats renamed to .xlsx.
    const signature = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
    if (signature.length < 4 || signature[0] !== 80 || signature[1] !== 75 || signature[2] !== 3 || signature[3] !== 4) throw new Error("File bukan Excel .xlsx yang valid. Simpan ulang menggunakan template.");
    let book;
    try {
      book = XLSX.read(buffer, { type: "array", cellFormula: true, sheetRows: 10002 });
    } catch {
      throw new Error("File Excel tidak dapat dibaca. Pastikan file tidak rusak atau dilindungi kata sandi.");
    }
    const existing = inventoryByRoom[targetRoomId] || [];
    const assets = validateInventoryWorkbook(book, existing);
    inventoryByRoom[targetRoomId] = [...existing, ...assets];
    window.inventoryDraft?.edited();
    renderAssets();
    excelStatus(`${assets.length} barang berhasil ditambahkan ke ${targetUnit} / ${targetName}. Periksa daftar lalu klik Simpan Pemeriksaan untuk menyimpan.`);
  } catch (error) {
    excelStatus(error.message || "Upload gagal. Periksa file dan coba lagi.", true);
  } finally {
    input.disabled = false;
    input.value = "";
  }
}
