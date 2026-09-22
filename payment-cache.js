(function () {
  'use strict';
  const database = 'gylsjqx-payment-files';
  let queue = Promise.resolve();
  function transaction(mode, action) {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(database, 1);
      request.onupgradeneeded = () => request.result.createObjectStore('files');
      request.onerror = () => reject(request.error);
      request.onblocked = () => reject(new Error('文件缓存暂时被其他窗口占用'));
      request.onsuccess = () => {
        const db = request.result;
        let result;
        try {
          const tx = db.transaction('files', mode);
          const operation = action(tx.objectStore('files'));
          if (operation) operation.onsuccess = () => { result = operation.result; };
          tx.oncomplete = () => { db.close(); resolve(result); };
          tx.onerror = tx.onabort = () => { db.close(); reject(tx.error || new Error('缓存操作未完成')); };
        } catch (e) { db.close(); reject(e); }
      };
    });
  }
  function run(mode, action) {
    const next = queue.then(() => transaction(mode, action));
    queue = next.catch(() => {});
    return next;
  }
  window.PaymentCache = {
    get: key => run('readonly', store => store.get(key)),
    put: (key, value) => run('readwrite', store => store.put(value, key)),
    remove: key => run('readwrite', store => store.delete(key)),
    clear: () => run('readwrite', store => store.clear()),
  };
})();
