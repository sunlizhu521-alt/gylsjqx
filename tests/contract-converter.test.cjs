const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function harness() {
  const timers = new Map(), frames = [], status = {};
  let timerId = 0;
  const timer = (fn, delay, repeat = false) => { const id = ++timerId; timers.set(id, { fn, delay, repeat }); return id; };
  const context = vm.createContext({
    ContractCore: {}, window: {}, location: { origin: 'https://example.test' },
    File, Blob, ArrayBuffer, Uint8Array, TextDecoder,
    setTimeout: (fn, delay) => timer(fn, delay), clearTimeout: id => timers.delete(id),
    setInterval: (fn, delay) => timer(fn, delay, true), clearInterval: id => timers.delete(id),
    document: {
      addEventListener() {}, getElementById: () => status,
      body: { append: frame => frames.push(frame) },
      createElement() {
        const events = new Map();
        const frame = {
          dataset: {}, messages: [], removed: false, setAttribute() {},
          contentWindow: { postMessage(message) { frame.messages.push(message); } },
          addEventListener: (name, fn) => events.set(name, fn),
          removeEventListener: name => events.delete(name),
          emit: name => events.get(name)?.(), remove() { this.removed = true; },
        };
        return frame;
      },
    },
  });
  const source = fs.readFileSync(require.resolve('../contract-tool.js'), 'utf8').replace(/\}\)\(\);\s*$/, `
    window.testApi = { ensureConverterFrame, handleConverterMessage, convertContractToPdf,
      pending: () => converterRequests.size };
  })();`);
  vm.runInContext(source, context);
  const api = context.window.testApi;
  return {
    ...api, frames, timers,
    tick(delay) {
      for (const [id, item] of [...timers]) if (item.delay === delay && timers.has(id)) {
        if (!item.repeat) timers.delete(id); item.fn();
      }
    },
    receive(frame, type, id, payload = {}, origin = 'https://edit.chaxus.com') {
      api.handleConverterMessage({ source: frame.contentWindow, origin, data: { type, id, payload } });
    },
  };
}

const flush = () => new Promise(resolve => setImmediate(resolve));

test('转换组件漏发 ready 时，主动状态响应仍可完成完整转换', async () => {
  const h = harness();
  const result = h.convertContractToPdf(new Blob(['synthetic']), '合同.xlsx');
  const frame = h.frames[0];
  frame.emit('load');
  const probe = frame.messages[0];
  assert.equal(probe.type, 'document:get-state');
  h.receive(frame, 'document:state', probe.id, { hasDocument: false });
  await flush();
  const open = frame.messages.find(item => item.type === 'document:open-buffer');
  assert(open);
  h.receive(frame, 'document:opened', open.id);
  await flush();
  const save = frame.messages.find(item => item.type === 'document:save');
  assert.equal(save.payload.targetExt, 'PDF');
  h.receive(frame, 'document:saved', save.id, { file: new Blob(['%PDF-1.7\n' + 'x'.repeat(110)]) });
  const pdf = await result;
  assert.equal(pdf.name, '合同.pdf');
  assert.equal(h.pending(), 0);
  assert.equal(h.timers.size, 0);
  assert.equal(h.frames.length, 1);
});

test('组件先触发 load 后才监听消息，轮询可重试；其他窗口和来源不能冒充就绪', async () => {
  const h = harness(), ready = h.ensureConverterFrame(), frame = h.frames[0];
  assert.equal(h.ensureConverterFrame(), ready);
  frame.emit('load'); h.tick(1000);
  assert.equal(frame.messages.length, 2);
  const id = frame.messages[0].id;
  h.receive(frame, 'document:state', id, {}, 'https://wrong.test');
  h.receive({ contentWindow: {} }, 'document:ready');
  h.receive(frame, 'document:state', 'wrong-id');
  h.receive(frame, 42);
  assert.equal(frame.dataset.ready, undefined);
  h.receive(frame, 'document:state', id);
  await ready;
  assert.equal(frame.dataset.ready, 'true');
  assert.equal(h.timers.size, 0);
});

test('原有 ready 通知仍可直接使用，重复生成复用已就绪组件', async () => {
  const h = harness(), ready = h.ensureConverterFrame(), frame = h.frames[0];
  h.receive(frame, 'document:ready'); await ready;
  await h.ensureConverterFrame();
  assert.equal(h.frames.length, 1);
  assert.equal(h.timers.size, 0);
});

test('加载超时移除失效组件，重试忽略旧窗口消息并建立新组件', async () => {
  const h = harness(), failed = assert.rejects(h.ensureConverterFrame(), /加载超时/), old = h.frames[0];
  h.tick(60000); await failed;
  assert(old.removed);
  assert.equal(h.timers.size, 0);
  const ready = h.ensureConverterFrame(), frame = h.frames[1];
  h.receive(old, 'document:ready'); assert.equal(frame.dataset.ready, undefined);
  h.receive(frame, 'document:ready'); await ready;
  assert.equal(frame.dataset.ready, 'true');
});

test('组件加载失败后可重试，不复用已拒绝的等待', async () => {
  const h = harness(), failed = assert.rejects(h.ensureConverterFrame(), /加载失败/);
  h.frames[0].emit('error'); await failed;
  assert(h.frames[0].removed);
  const ready = h.ensureConverterFrame();
  h.receive(h.frames[1], 'document:ready'); await ready;
  assert.equal(h.timers.size, 0);
});

for (const phase of ['open', 'save']) test(`${phase} 超时恢复组件，不再重复等待四分钟`, async () => {
  const h = harness(), failed = assert.rejects(h.convertContractToPdf(new Blob(['synthetic']), '合同.xlsx'), /超时/);
  const frame = h.frames[0]; h.receive(frame, 'document:ready'); await flush();
  if (phase === 'save') {
    h.receive(frame, 'document:opened', frame.messages[0].id); await flush();
  }
  h.tick(240000); await failed;
  assert(frame.removed);
  assert.equal(h.pending(), 0);
  assert.equal(h.timers.size, 0);
  assert.equal(frame.messages.filter(item => item.type === 'document:save').length, phase === 'save' ? 1 : 0);
  const ready = h.ensureConverterFrame();
  h.receive(h.frames[1], 'document:ready'); await ready;
});
