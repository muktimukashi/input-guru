/* Local drafts stay in this browser; only submitInventory sends data to the server. */
(() => {
  const key = 'inventory-draft-v1';
  const fields = ['teacherName', 'unitSelect', 'roomSelect'];
  const pendingFields = ['manualAssetName', 'manualAssetQty', 'manualAssetUnit', 'manualAssetCondition', 'manualAssetNote'];
  const completed = new Set();
  let restoring = false;
  const element = id => document.getElementById(id);
  const values = ids => Object.fromEntries(ids.map(id => [id, element(id)?.value || '']));
  const status = (message, error = false) => {
    element('draftStatus').textContent = message;
    element('draftStatus').style.color = error ? '#b91c1c' : '#64748b';
  };

  function save() {
    if (restoring) return;
    try {
      const inventories = Object.fromEntries(Object.entries(inventoryByRoom).filter(([id]) => !completed.has(id)));
      const currentCompleted = completed.has(String(selectedRoomId));
      const form = currentCompleted ? { teacherName: '', unitSelect: '', roomSelect: '' } : values(fields);
      const pending = currentCompleted ? {} : values(pendingFields);
      if (!Object.values(form).some(Boolean) && !Object.values(inventories).some(items => items.length)) {
        localStorage.removeItem(key);
        return;
      }
      const draft = {
        version: 1, form, pending, inventories,
        selectedRoomId: currentCompleted ? null : selectedRoomId,
        rooms: rooms.filter(room => Object.hasOwn(inventories, room.id)),
        savedAt: new Date().toISOString()
      };
      localStorage.setItem(key, JSON.stringify(draft));
      status('Draft tersimpan otomatis di browser ini. Belum dikirim ke sistem.');
    } catch {
      status('Draft gagal disimpan di browser. Jangan tutup atau refresh halaman sebelum menyimpan pemeriksaan.', true);
    }
  }

  function restore() {
    restoring = true;
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return;
      const draft = JSON.parse(raw);
      if (draft.version !== 1 || !draft.form || !draft.inventories || typeof draft.inventories !== 'object' || Array.isArray(draft.inventories) || !Array.isArray(draft.rooms)) throw new Error('Invalid draft');
      for (const items of Object.values(draft.inventories)) {
        if (!Array.isArray(items) || items.some(item => !item || typeof item.name !== 'string' || typeof item.itemUnit !== 'string' || typeof item.condition !== 'string')) throw new Error('Invalid inventory');
      }
      for (const id of fields) if (typeof draft.form[id] !== 'string') throw new Error('Invalid form');
      inventoryByRoom = Object.assign(Object.create(null), draft.inventories);
      Object.keys(inventoryByRoom).forEach(id => manuallyEnteredRooms.add(id));
      rooms = draft.rooms;
      fields.forEach(id => { element(id).value = draft.form[id]; });
      changeRoom();
      pendingFields.forEach(id => {
        if (element(id) && typeof draft.pending?.[id] === 'string') element(id).value = draft.pending[id];
      });
      status('Draft sebelumnya dipulihkan. Lanjutkan pengisian lalu klik Simpan Pemeriksaan.');
    } catch {
      status('Draft sebelumnya tidak dapat dipulihkan. Periksa isian sebelum melanjutkan.', true);
    } finally {
      restoring = false;
    }
  }

  function onEdit(event) {
    const target = event.target;
    if (!target.matches('input:not([type="file"]), select, textarea')) return;
    completed.delete(String(selectedRoomId));
    // Table edits use onchange for rendering; capture the current keystroke too.
    const row = target.closest('.asset-table tbody tr');
    if (row) {
      const asset = inventoryByRoom[selectedRoomId]?.[row.sectionRowIndex];
      if (asset) {
        const controls = row.querySelectorAll('input, textarea');
        asset.name = controls[0].value;
        asset.actualQty = controls[1].value;
        asset.registeredQty = controls[1].value;
        asset.itemUnit = controls[2].value;
        asset.condition = controls[3].value;
        asset.note = controls[4].value;
      }
    }
    save();
  }

  window.inventoryDraft = {
    save,
    edited() { completed.delete(String(selectedRoomId)); },
    complete() {
      completed.add(String(selectedRoomId));
      save();
      // Do not replace a storage failure warning with a success message.
      if (!element('draftStatus').textContent.startsWith('Draft gagal')) status('Pemeriksaan berhasil disimpan. Draft ruangan ini sudah dihapus.');
    }
  };
  restore();
  document.addEventListener('input', onEdit);
  document.addEventListener('change', onEdit);
  loadRooms().catch(() => status('Daftar ruangan belum dapat dimuat. Isian draft tetap tersedia.', true));
})();
