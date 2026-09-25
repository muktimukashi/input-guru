const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));

function setup() {
  const elements = new Map();
  const context = vm.createContext({
    XLSX, window: { XLSX }, Uint8Array,
    document: { getElementById(id) {
      if (!elements.has(id)) elements.set(id, { value: '', style: {}, textContent: '', disabled: false });
      return elements.get(id);
    } },
  });
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const inline = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1];
  vm.runInContext(inline.replace('loadRooms();', ''), context);
  vm.runInContext(fs.readFileSync(path.join(root, 'excel-import.js'), 'utf8'), context);
  vm.runInContext(`selectedRoomId = 'room1'; inventoryByRoom.room1 = [];
    changeRoom = () => {}; renderAssets = () => {};
    document.getElementById('unitSelect').value = 'SD';
    document.getElementById('roomSelect').value = 'Kelas 1A';`, context);
  return context;
}

function workbook(rows) {
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([
    ['Nama Barang', 'Jumlah', 'Satuan/Unit', 'Kondisi', 'Keterangan'], ...rows
  ]), 'Inventaris');
  return book;
}

test('template round trip: blank inventory, instructions and valid data', () => {
  const ctx = setup();
  const template = ctx.createInventoryTemplate();
  const decoded = XLSX.read(XLSX.write(template, { type: 'buffer', bookType: 'xlsx' }));
  assert.deepEqual(decoded.SheetNames, ['Inventaris', 'Petunjuk']);
  assert.throws(() => ctx.validateInventoryWorkbook(decoded, []), /belum berisi/);
  XLSX.utils.sheet_add_aoa(decoded.Sheets.Inventaris, [['Meja', 2, 'Unit', 'Baik', '']], { origin: 'A2' });
  const assets = ctx.validateInventoryWorkbook(decoded, []);
  assert.equal(assets.length, 1);
  assert.equal(assets[0].actualQty, 2);
  assert.equal(assets[0].note, '');
});

test('required fields, invalid quantities and duplicates report original row numbers', () => {
  const ctx = setup();
  for (const qty of [0, -1, 1.5, 'abc', '', 2147483648, true]) {
    assert.throws(() => ctx.validateInventoryWorkbook(workbook([['Meja', qty, 'Unit', 'Baik']]), []), /Baris 2:.*Jumlah/);
  }
  assert.throws(() => ctx.validateInventoryWorkbook(workbook([[], ['', 1, '', '']]), []), /Baris 3:.*Nama Barang.*Satuan\/Unit.*Kondisi/);
  assert.throws(() => ctx.validateInventoryWorkbook(workbook([['Meja', 1, 'Unit', 'Baik'], [' MEJA ', 2, 'Unit', 'Baik']]), []), /Baris 3:.*duplikat/);
  assert.throws(() => ctx.validateInventoryWorkbook(workbook([[' meja ', 1, 'Unit', 'Baik']]), [{ name: 'Meja' }]), /duplikat/);
});

test('reject malformed template, formulas, extra data and excessive rows', () => {
  const ctx = setup();
  assert.throws(() => ctx.validateInventoryWorkbook({ Sheets: {} }, []), /Sheet Inventaris/);
  const book = workbook([['Meja', 1, 'Unit', 'Baik']]);
  book.Sheets.Inventaris.A1.v = 'Nama';
  assert.throws(() => ctx.validateInventoryWorkbook(book, []), /Judul kolom/);
  book.Sheets.Inventaris.A1.v = 'Nama Barang';
  book.Sheets.Inventaris.B2.f = '1+1';
  assert.throws(() => ctx.validateInventoryWorkbook(book, []), /bukan rumus/);
  assert.throws(() => ctx.validateInventoryWorkbook(workbook([['Meja', 1, 'Unit', 'Baik', '', 'extra']]), []), /di luar kolom/);
  assert.throws(() => ctx.validateInventoryWorkbook(workbook(Array.from({ length: 1001 }, (_, i) => [`Barang ${i}`, 1, 'Unit', 'Baik'])), []), /Maksimal 1.000/);
});

test('upload appends only valid data and repeat upload preserves existing list', async () => {
  const ctx = setup();
  const buffer = XLSX.write(workbook([['Meja', 2, 'Unit', 'Baik']]), { type: 'array', bookType: 'xlsx' });
  const input = { files: [{ name: 'inventaris.xlsx', size: buffer.byteLength, arrayBuffer: async () => buffer }], value: 'file' };
  await ctx.importInventoryExcel(input);
  assert.equal(vm.runInContext('inventoryByRoom.room1.length', ctx), 1);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /berhasil ditambahkan/);
  await ctx.importInventoryExcel(input);
  assert.equal(vm.runInContext('inventoryByRoom.room1.length', ctx), 1);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /duplikat/);
  assert.equal(input.disabled, false);
  assert.equal(input.value, '');
});

test('invalid batch is atomic; changed room and invalid files are rejected', async () => {
  const ctx = setup();
  const buffer = XLSX.write(workbook([['Meja', 1, 'Unit', 'Baik'], ['Kursi', 0, 'Unit', 'Baik']]), { type: 'array', bookType: 'xlsx' });
  const input = { files: [{ name: 'inventaris.xlsx', size: buffer.byteLength, arrayBuffer: async () => buffer }] };
  await ctx.importInventoryExcel(input);
  assert.equal(vm.runInContext('inventoryByRoom.room1.length', ctx), 0);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /Baris 3/);
  input.files[0].arrayBuffer = async () => {
    ctx.document.getElementById('roomSelect').value = 'Kelas 2A';
    return buffer;
  };
  await ctx.importInventoryExcel(input);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /Ruangan berubah/);
  input.files[0].size = 6 * 1024 * 1024;
  await ctx.importInventoryExcel(input);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /maksimal 5 MB/);
  input.files[0].size = 3;
  input.files[0].arrayBuffer = async () => new Uint8Array([1, 2, 3]).buffer;
  await ctx.importInventoryExcel(input);
  assert.match(ctx.document.getElementById('excelStatus').textContent, /bukan Excel/);
});
