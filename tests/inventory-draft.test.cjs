const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const key = 'inventory-draft-v1';

function setup(storage = new Map(), failStorage = false, db = null) {
  const elements = new Map();
  const events = {};
  const document = {
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, { value: '', textContent: '', style: {} });
      return elements.get(id);
    },
    addEventListener(name, fn) { events[name] = fn; }
  };
  const context = vm.createContext({ window: db ? {supabase: {createClient: () => db}} : {}, document, console: {log() {}, error() {}}, alert() {},
    localStorage: {
      getItem: name => storage.get(name) || null,
      setItem(name, value) { if (failStorage) throw Error('quota'); storage.set(name, value); },
      removeItem(name) { if (failStorage) throw Error('blocked'); storage.delete(name); }
    }
  });
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  vm.runInContext([...html.matchAll(/<script>([\s\S]*?)<\/script>/g)][0][1], context);
  vm.runInContext(`updateSuggestions = () => {}; loadRooms = async () => {};
    renderAssets = () => window.inventoryDraft?.save();`, context);
  vm.runInContext(fs.readFileSync(path.join(root, 'inventory-draft.js'), 'utf8'), context);
  return { context, storage, events, document, run: code => vm.runInContext(code, context) };
}

function fill(app) {
  app.document.getElementById('teacherName').value = 'Budi';
  app.document.getElementById('unitSelect').value = 'SD';
  app.document.getElementById('roomSelect').value = 'Kelas 1A';
  app.run(`changeRoom(); inventoryByRoom[selectedRoomId] = [{name:'Meja', itemUnit:'Unit', actualQty:2, condition:'Baik', note:''}];`);
  app.document.getElementById('manualAssetName').value = 'Kursi belum ditambahkan';
  app.context.window.inventoryDraft.save();
}

test('refresh restores identity, inventory and unfinished manual fields', () => {
  const first = setup();
  fill(first);
  const second = setup(first.storage);
  assert.equal(second.document.getElementById('teacherName').value, 'Budi');
  assert.equal(second.document.getElementById('roomSelect').value, 'Kelas 1A');
  assert.equal(second.run('inventoryByRoom[selectedRoomId][0].actualQty'), 2);
  assert.equal(second.document.getElementById('manualAssetName').value, 'Kursi belum ditambahkan');
  assert.match(second.document.getElementById('draftStatus').textContent, /dipulihkan/);
});

test('known room IDs survive refresh without clearing their inventory', () => {
  const app = setup();
  app.run(`rooms = [{id:12, name:'Kelas 1A', unit:'SD', code:'R12'}]`);
  fill(app);
  const restored = setup(app.storage);
  assert.equal(restored.run('selectedRoomId'), 12);
  assert.equal(restored.run('inventoryByRoom[12].length'), 1);
});

test('typing table values is saved before blur or onchange', () => {
  const app = setup();
  fill(app);
  const controls = ['Meja baru', '3', 'Buah', 'Rusak Ringan', 'Sedang diketik'].map(value => ({value}));
  app.events.input({ target: {
    matches: () => true,
    closest: () => ({ sectionRowIndex: 0, querySelectorAll: () => controls })
  }});
  const restored = setup(app.storage);
  assert.equal(restored.run('inventoryByRoom[selectedRoomId][0].note'), 'Sedang diketik');
  assert.equal(restored.run('inventoryByRoom[selectedRoomId][0].actualQty'), '3');
});

test('successful save clears current draft while preserving another room', () => {
  const app = setup();
  fill(app);
  app.run(`inventoryByRoom.other = [{name:'Kursi', itemUnit:'Unit', condition:'Baik'}]`);
  app.context.window.inventoryDraft.complete();
  const draft = JSON.parse(app.storage.get(key));
  assert.deepEqual(Object.keys(draft.inventories), ['other']);
  assert.equal(draft.form.roomSelect, '');
  app.context.window.inventoryDraft.edited();
  app.context.window.inventoryDraft.save();
  assert.ok(JSON.parse(app.storage.get(key)).inventories['manual:SD:Kelas 1A']);
});

test('last completed draft is removed and offline submission retains draft', async () => {
  const app = setup();
  fill(app);
  await app.run('submitInventory()');
  assert.ok(app.storage.has(key));
  assert.equal(app.document.getElementById('submitButton').textContent, 'Simpan Pemeriksaan');
  app.context.window.inventoryDraft.complete();
  assert.equal(app.storage.has(key), false);
});

test('storage errors and corrupt drafts produce visible messages', () => {
  const blocked = setup(new Map(), true);
  fill(blocked);
  assert.match(blocked.document.getElementById('draftStatus').textContent, /gagal disimpan/);
  const corrupt = setup(new Map([[key, '{broken']]));
  assert.match(corrupt.document.getElementById('draftStatus').textContent, /tidak dapat dipulihkan/);
});

test('server failures retain draft; only successful header and detail save clear it', async () => {
  for (const fail of ['header', 'details', null]) {
    const db = { from(table) {
      return { insert() {
        if (table === 'inventory_check_items') return Promise.resolve({error: fail === 'details' ? new Error('failed') : null});
        return {select() {return {single: async () => ({data: {id: 1}, error: fail === 'header' ? new Error('failed') : null})};}};
      }};
    }};
    const app = setup(new Map(), false, db);
    app.run(`rooms = [{id:12, name:'Kelas 1A', unit:'SD', code:'R12'}]`);
    fill(app);
    await app.run('submitInventory()');
    assert.equal(app.storage.has(key), fail !== null);
  }
});
