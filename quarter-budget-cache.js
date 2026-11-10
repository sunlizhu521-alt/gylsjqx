(function () {
  'use strict';
  const database = 'gylsjqx-quarter-budget-files';
  const roles = new Set(['negotiated', 'ledger', 'budget', 'evidence', 'existing']);
  let queue = Promise.resolve();

  function transaction(mode, action) {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(database, 1);
      let blocked = false;
      request.onupgradeneeded = () => request.result.createObjectStore('files');
      request.onerror = () => reject(request.error || new Error('无法打开文件缓存'));
      request.onblocked = () => { blocked = true; reject(new Error('文件缓存被其他窗口占用')); };
      request.onsuccess = () => {
        const db = request.result;
        if (blocked) { db.close(); return; }
        db.onversionchange = () => db.close();
        try {
          const tx = db.transaction('files', mode);
          let value;
          const operation = action(tx.objectStore('files'));
          if (operation) operation.onsuccess = () => { value = operation.result; };
          tx.oncomplete = () => { db.close(); resolve(value); };
          tx.onerror = tx.onabort = () => { db.close(); reject(tx.error || new Error('文件缓存操作未完成')); };
        } catch (error) { db.close(); reject(error); }
      };
    });
  }
  function run(mode, action) {
    const next = queue.then(() => transaction(mode, action));
    queue = next.catch(() => {});
    return next;
  }
  function checkRole(role) {
    if (!roles.has(role)) throw new Error('未知文件槽位');
  }
  function state(store, settings) {
    store.put({kind:'settings', value:settings}, '@settings');
  }
  window.QuarterBudgetCache = {
    database,
    async load() {
      const entries = await run('readonly', store => store.getAll());
      const files = {};
      let settings = null;
      for (const entry of entries || []) {
        if (entry?.kind === 'file' && roles.has(entry.role)) files[entry.role] = entry.value;
        else if (entry?.kind === 'settings') settings = entry.value;
      }
      return {files, settings};
    },
    saveFile(role, file, settings) {
      checkRole(role);
      return run('readwrite', store => {
        store.put({kind:'file', role, value:file}, role);
        state(store, settings);
      });
    },
    removeFile(role, settings) {
      checkRole(role);
      return run('readwrite', store => { store.delete(role); state(store, settings); });
    },
    saveState(settings) { return run('readwrite', store => state(store, settings)); },
    clear() { return run('readwrite', store => store.clear()); }
  };
})();
