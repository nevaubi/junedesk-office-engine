import assert from 'node:assert/strict';
import { test } from 'node:test';
import { EnginePool } from './engine-pool.mjs';

test('large native saves send only control JSON to platform and never forward engine token to storage',async()=>{
  const original=globalThis.fetch, calls=[];
  globalThis.fetch=async(url,init)=>{
    calls.push({url,init});
    if(calls.length===1)return Response.json({url:'https://synthetic-storage.test/file',checksum:'synthetic-checksum'});
    if(calls.length===2)return new Response(null,{status:200});
    return Response.json({version:9});
  };
  try {
    const saved=await EnginePool.prototype.persist.call({platformUrl:'https://synthetic-platform.test'},
      {docId:'test-doc',version:8,token:'synthetic-token',kind:{mime:'test/type'}},Buffer.alloc(3*1024*1024+1), 'operation');
    assert.equal(saved.version,9);assert.equal(calls.length,3);
    const [prepare,upload,commit]=calls;
    assert.equal(prepare.init.method,'POST');assert.ok(prepare.init.body.length<300);
    assert.equal(upload.init.headers.authorization,undefined);assert.equal(upload.init.headers['x-amz-checksum-sha256'],'synthetic-checksum');
    assert.equal(upload.init.body.byteLength,3*1024*1024+1);
    assert.equal(commit.init.method,'PUT');assert.equal(commit.init.body,prepare.init.body);
    assert.equal(commit.init.headers.authorization,'Bearer synthetic-token');
  } finally {globalThis.fetch=original;}
});

test('failed staged native upload never sends a revision commit',async()=>{
  const original=globalThis.fetch;let calls=0;
  globalThis.fetch=async()=>++calls===1?Response.json({url:'https://synthetic-storage.test/file',checksum:'synthetic'}):new Response(null,{status:503});
  try {
    await assert.rejects(EnginePool.prototype.persist.call({platformUrl:'https://synthetic-platform.test'},
      {docId:'test-doc',version:8,token:'synthetic-token'},Buffer.alloc(3*1024*1024+1),'operation'),/upload failed/);
    assert.equal(calls,2);
  } finally {globalThis.fetch=original;}
});

for (const failure of ['network', 'gateway']) test(`lost successful staged commit response (${failure}) reuses the operation without reuploading`, async () => {
  const original = globalThis.fetch, calls = [];
  let committed = false, revisions = 0;
  globalThis.fetch = async (url, init) => {
    calls.push({ url, init });
    if (init.method === 'POST') return Response.json({ url: 'https://synthetic-storage.test/file', checksum: 'synthetic' });
    if (url.includes('synthetic-storage')) return new Response(null, { status: 200 });
    if (!committed) {
      committed = true; revisions++;
      if (failure === 'network') throw new TypeError('fetch failed after commit');
      return new Response('gateway lost response', { status: 502 });
    }
    return Response.json({ version: 9, replayed: true });
  };
  try {
    const saved = await EnginePool.prototype.persist.call({ platformUrl: 'https://synthetic-platform.test' },
      { docId: 'test-doc', version: 8, token: 'synthetic-token' }, Buffer.alloc(3 * 1024 * 1024 + 1), 'same-operation');
    assert.equal(saved.version, 9); assert.equal(saved.replayed, true); assert.equal(revisions, 1);
    assert.equal(calls.length, 4); assert.equal(calls[2].init.body, calls[3].init.body);
    assert.equal(JSON.parse(calls[3].init.body).operationId, 'same-operation');
    assert.equal(calls.filter(c => c.url.includes('synthetic-storage')).length, 1);
  } finally { globalThis.fetch = original; }
});

test('small native save retries the identical binary operation only once', async () => {
  const original = globalThis.fetch, calls = [], bytes = Buffer.from('synthetic-native-bytes');
  globalThis.fetch = async (url, init) => {
    calls.push(init);
    if (calls.length === 1) throw new TypeError('connection reset after commit');
    return Response.json({ version: 4, replayed: true });
  };
  try {
    const saved = await EnginePool.prototype.persist.call({ platformUrl: 'https://synthetic-platform.test' },
      { docId: 'test-doc', version: 3, token: 'synthetic-token', kind: { mime: 'test/type' } }, bytes, 'same-operation');
    assert.equal(saved.version, 4); assert.equal(calls.length, 2);
    for (const call of calls) {
      assert.equal(call.body, bytes); assert.equal(call.headers['idempotency-key'], 'same-operation');
      assert.equal(call.headers['if-match'], '3');
    }
  } finally { globalThis.fetch = original; }
});

test('native save never retries permission, conflict, validation, or malformed confirmation responses', async () => {
  const original = globalThis.fetch;
  try {
    for (const status of [401, 403, 409, 413, 422, 200]) {
      let calls = 0;
      globalThis.fetch = async () => { calls++; return Response.json({ error: 'synthetic failure' }, { status }); };
      await assert.rejects(EnginePool.prototype.persist.call({ platformUrl: 'https://synthetic-platform.test' },
        { docId: 'test-doc', version: 3, token: 'synthetic-token', kind: { mime: 'test/type' } }, Buffer.from('tiny'), 'same-operation'));
      assert.equal(calls, 1, `HTTP ${status} must not repeat`);
    }
    for (const failure of ['network', 'gateway']) {
      let calls = 0;
      globalThis.fetch = async () => { calls++; if (failure === 'network') throw new TypeError('offline'); return new Response(null, { status: 502 }); };
      await assert.rejects(EnginePool.prototype.persist.call({ platformUrl: 'https://synthetic-platform.test' },
        { docId: 'test-doc', version: 3, token: 'synthetic-token', kind: { mime: 'test/type' } }, Buffer.from('tiny'), 'same-operation'));
      assert.equal(calls, 2, 'transient retry must be bounded');
    }
  } finally { globalThis.fetch = original; }
});
