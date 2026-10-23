/* Separate from payment-file storage; update both snapshots in one transaction. */
(()=>{'use strict';
function run(current){return new Promise((resolve,reject)=>{
 const request=indexedDB.open('gylsjqx-fulfillment-history',1);
 request.onupgradeneeded=()=>request.result.createObjectStore('snapshots');
 request.onerror=()=>reject(request.error);
 request.onblocked=()=>reject(Error('履约历史缓存被其他窗口占用，请关闭其他窗口后重试'));
 request.onsuccess=()=>{
  const db=request.result,tx=db.transaction('snapshots',current?'readwrite':'readonly'),store=tx.objectStore('snapshots');
  let value,error;
  const read=store.get('history');
  read.onsuccess=()=>{try{const history=read.result?{...read.result,current:{...read.result.current,result:FulfillmentCore.currentScope(read.result.current.result)},previous:read.result.previous?{...read.result.previous,result:FulfillmentCore.currentScope(read.result.previous.result)}:null}:null;value=current?FulfillmentCore.advance(history,current):history;if(current)store.put(value,'history');}catch(e){error=e;tx.abort();}};
  tx.oncomplete=()=>{db.close();resolve(value);};
  tx.onabort=tx.onerror=()=>{db.close();reject(error||tx.error||Error('历史保存失败，请检查浏览器存储空间'));};
 };
});}
window.FulfillmentCache={get:()=>run(),save:current=>run(current)};
})();
