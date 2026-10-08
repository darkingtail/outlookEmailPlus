// Dependency-free DOM/async behavior tests for the actual inline customer-page script.
// Run: node tests/customer_mail_detail_behavior.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor(tag = 'div') {
    this.tagName = tag;
    this.children = [];
    this.listeners = {};
    this.attributes = {};
    this.dataset = {};
    this.value = '';
    this._text = '';
    this._html = '';
  }
  append(...nodes) { this.children.push(...nodes); }
  replaceChildren(...nodes) { this.children = nodes; this._text = ''; this._html = ''; }
  set textContent(value) { this.replaceChildren(); this._text = String(value); }
  get textContent() { return this._text + this.children.map(node => node.textContent).join(''); }
  set innerHTML(value) {
    this.replaceChildren();
    this._html = value;
    if (value.includes('detail-empty-description')) {
      const description = new Element();
      description.className = 'detail-empty-description';
      this.append(description);
    }
  }
  get innerHTML() { return this._html; }
  setAttribute(key, value) { this.attributes[key] = value; }
  addEventListener(name, handler) { (this.listeners[name] ||= []).push(handler); }
  dispatch(name) { for (const fn of this.listeners[name] || []) fn({ target: this }); }
  querySelector(selector) {
    for (const node of this.children) {
      if (node.className === selector.slice(1)) return node;
      const found = node.querySelector(selector);
      if (found) return found;
    }
    return null;
  }
}

const root = process.argv[2] || path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'templates/customer_mail.html'), 'utf8');
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(match => match[1]).join('\n');
assert.ok(script.includes('/api/customer/messages/detail'), 'page must call the customer detail API');
assert.deepEqual([...html.matchAll(/data-view="([^"]+)"/g)].map(match => match[1]), ['all', 'junk']);
assert.ok(!html.includes('verification_code'), 'customer page must not extract or display verification codes');

function setup() {
  const elements = new Map();
  const node = id => {
    if (!elements.has(id)) elements.set(id, new Element());
    return elements.get(id);
  };
  const views = ['all', 'junk'].map(view => {
    const el = new Element('button');
    el.dataset.view = view;
    return el;
  });
  const counts = views.map(view => node(`count-${view.dataset.view}`));
  const requests = [];
  const document = {
    getElementById: node,
    createElement: tag => new Element(tag),
    querySelector(selector) {
      const view = selector.match(/data-count="(\w+)"/)[1];
      return node(`count-${view}`);
    },
    querySelectorAll: selector => selector === '[data-count]' ? counts : views,
  };
  const storage = new Proxy({}, { get() { throw new Error('must not use browser persistent storage'); } });
  const context = vm.createContext({
    document, console, AbortController, localStorage: storage, sessionStorage: storage,
    fetch(url, options) {
      let resolve;
      const promise = new Promise(done => { resolve = done; });
      requests.push({
        url, options, payload: JSON.parse(options.body),
        respond(data, ok = true) { resolve({ ok, json: async () => data }); },
      });
      return promise; // Deliberately ignore abort to exercise stale-response guards.
    },
  });
  vm.runInContext(script, context, { filename: 'customer_mail.inline.js' });
  node('email').value = 'demo@customer.test';
  node('secret').value = 'fixture-customer-secret';
  return { node, requests, run: code => vm.runInContext(code, context), views };
}

const settle = () => new Promise(resolve => setImmediate(resolve));
const summary = id => ({
  message_id: id, subject: `Subject ${id}`, from: 'sender@test',
  date: '2026-10-08T00:00:00Z', preview: `PREVIEW-${id}`,
});
const body = value => ({
  success: true,
  message: { subject: 'Complete subject', from: 'sender@test', to: 'demo@customer.test', cc: 'cc@test', body: value },
});
async function listReady(t) {
  const pending = t.run('query()');
  t.requests.at(-1).respond({ success: true, messages: [summary('a'), summary('b')] });
  await pending;
  return t.requests.at(-1);
}

(async () => {
  let passed = 0;
  async function check(name, fn) { await fn(); passed++; console.log(`PASS ${name}`); }

  await check('refresh button fits its desktop grid track without being clipped', async () => {
    const formCss = html.match(/\.connection-form\s*\{([^}]+)\}/)[1];
    const buttonCss = html.match(/\.icon-button\s*\{([^}]+)\}/)[1];
    const trackWidth = Number(formCss.match(/grid-template-columns:[^;]*\s(\d+)px\s*;/)[1]);
    const buttonWidth = Number(buttonCss.match(/width:\s*(\d+)px/)[1]);
    assert.ok(trackWidth >= buttonWidth, `${buttonWidth}px refresh button overflows ${trackWidth}px grid track`);
  });

  await check('default query is all mail, with only all and junk views', async () => {
    const t = setup();
    assert.equal(t.run('view'), 'all');
    const pending = t.run('query()');
    assert.equal(t.requests.at(-1).payload.view, 'all');
    t.requests.at(-1).respond({ success: true, messages: [] });
    await pending;
    assert.equal(t.node('list-title').textContent, '全部邮件 · 最多 50 封');
  });

  await check('full body uses POST credentials, plaintext and no persistent storage', async () => {
    const t = setup();
    const request = await listReady(t);
    assert.equal(request.url, '/api/customer/messages/detail');
    assert.equal(request.options.method, 'POST');
    assert.equal(request.options.cache, 'no-store');
    assert.deepEqual(request.payload, {
      email: 'demo@customer.test', secret: 'fixture-customer-secret', message_id: 'a',
    });
    assert.ok(t.node('detail').textContent.includes('正在读取完整正文'));
    assert.ok(!t.node('detail').textContent.includes('PREVIEW'));
    const full = '<img src=x onerror=bad()> Full paragraph\n' + 'end '.repeat(600);
    request.respond(body(full));
    await settle();
    assert.equal(t.node('detail').querySelector('.detail-content').textContent, full);
    assert.equal(t.node('detail').querySelector('.detail-content').innerHTML, '');
    assert.ok(t.node('detail').textContent.includes('收件人：demo@customer.test'));
    assert.ok(t.node('detail').textContent.includes('抄送：cc@test'));
    assert.equal(t.node('detail').attributes['aria-busy'], 'false');
  });

  await check('late detail cannot overwrite the newly selected message', async () => {
    const t = setup();
    const old = await listReady(t);
    t.node('messages').children[1].dispatch('click');
    const latest = t.requests.at(-1);
    assert.equal(latest.payload.message_id, 'b');
    assert.equal(old.options.signal.aborted, true);
    latest.respond(body('NEW-BODY'));
    await settle();
    old.respond(body('OLD-BODY'));
    await settle();
    assert.ok(t.node('detail').textContent.includes('NEW-BODY'));
    assert.ok(!t.node('detail').textContent.includes('OLD-BODY'));
  });

  await check('editing credentials clears displayed data and ignores in-flight detail', async () => {
    const t = setup();
    const old = await listReady(t);
    t.node('secret').value = 'new-key';
    t.node('secret').dispatch('input');
    assert.equal(old.options.signal.aborted, true);
    old.respond(body('OLD-ACCOUNT-CONTENT'));
    await settle();
    assert.ok(!t.node('detail').textContent.includes('OLD-ACCOUNT-CONTENT'));
    assert.equal(t.run('messages.length'), 0);
    assert.equal(t.run('credentials'), null);
  });

  await check('late list result and finally cannot affect a newer query', async () => {
    const t = setup();
    const first = t.run('query()');
    const old = t.requests.at(-1);
    const second = t.run('query()');
    const latest = t.requests.at(-1);
    old.respond({ success: true, messages: [summary('old')] });
    await first;
    assert.equal(t.node('submit').disabled, true);
    assert.equal(t.requests.length, 2);
    latest.respond({ success: true, messages: [summary('new')] });
    await second;
    assert.equal(t.requests.at(-1).payload.message_id, 'new');
    assert.equal(t.node('submit').disabled, false);
  });

  await check('switching folder cancels detail and queries the selected folder', async () => {
    const t = setup();
    const old = await listReady(t);
    t.run('selectView("junk")');
    assert.equal(old.options.signal.aborted, true);
    assert.equal(t.requests.at(-1).payload.view, 'junk');
    old.respond(body('PREVIOUS-FOLDER'));
    await settle();
    assert.ok(!t.node('detail').textContent.includes('PREVIOUS-FOLDER'));
  });

  await check('detail errors are safe text, with retry and no preview fallback', async () => {
    const t = setup();
    const request = await listReady(t);
    request.respond({ success: false, message: '<img onerror=bad> expired reference' }, false);
    await settle();
    assert.equal(t.node('detail').querySelector('.detail-content').innerHTML, '');
    assert.ok(!t.node('detail').textContent.includes('PREVIEW'));
    const retry = t.node('detail').children.at(-1);
    assert.equal(retry.textContent, '重新读取');
    retry.dispatch('click');
    t.requests.at(-1).respond(body('RETRIED-BODY'));
    await settle();
    assert.ok(t.node('detail').textContent.includes('RETRIED-BODY'));
  });

  await check('blank body is an empty email, never its preview', async () => {
    const t = setup();
    const request = await listReady(t);
    request.respond(body(''));
    await settle();
    assert.equal(t.node('detail').querySelector('.detail-content').textContent, '这封邮件没有正文。');
    assert.ok(!t.node('detail').textContent.includes('PREVIEW'));
  });

  await check('changing credentials during list fetch cannot restore previous results', async () => {
    const t = setup();
    const pending = t.run('query()');
    const old = t.requests.at(-1);
    t.node('email').value = 'another@customer.test';
    t.node('email').dispatch('input');
    old.respond({ success: true, messages: [summary('private-old')] });
    await pending;
    assert.equal(t.run('messages.length'), 0);
    assert.equal(t.requests.length, 1);
    assert.equal(t.node('submit').disabled, false);
  });
  console.log(`UI_RESULT=${passed} passed`);
})().catch(error => { console.error(error); process.exitCode = 1; });
