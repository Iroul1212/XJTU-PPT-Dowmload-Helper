// ==UserScript==
// @name         XJTU PPT Helper (multi-site)
// @namespace    https://XJTUPPTHelper.com/
// @version      1.1
// @description  在学习空间/课堂平台上列出 PPT 直链与预览，支持一键下载、复制CSV/文本，优化UI交互与多余调试信息（紧凑版）。
// @author       Monika & Noan Cliffe (Enhanced)
// @match        https://lms.xjtu.edu.cn/*
// @match        http://lms.xjtu.edu.cn/*
// @match        https://ispace.xjtu.edu.cn/*
// @match        http://ispace.xjtu.edu.cn/*
// @match        https://v-ispace.xjtu.edu.cn:*/*
// @match        http://v-ispace.xjtu.edu.cn:*/*
// @match        https://class.xjtu.edu.cn/*
// @match        http://class.xjtu.edu.cn/*
// @match        https://v-class.xjtu.edu.cn:*/*
// @match        http://v-class.xjtu.edu.cn:*/*
// @run-at       document-end
// @license      GPL
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // ---------- utils ----------
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const qs = (s, r = document) => r.querySelector(s);
    const json = (x) => { try { return JSON.parse(x); } catch { return null; } };
    const csvEscape = (s='') => `"${String(s).replace(/"/g,'""')}"`;

    // ---------- style ----------
    const css = `
    :root {
        --xjtu-bg: rgba(30, 30, 34, 0.9);
        --xjtu-border: rgba(255, 255, 255, 0.2);
        --xjtu-hover: rgba(255, 255, 255, 0.1);
        --xjtu-text: #eaeaea;
        --xjtu-muted: #aaa;
        --xjtu-accent: #61afef;
    }
    #xjtu-ppt-helper {
        position: fixed; top: 30px; right: 30px; z-index: 2147483647 !important; pointer-events: auto !important;
        background: var(--xjtu-bg); color: var(--xjtu-text);
        border: 1px solid var(--xjtu-border); border-radius: 8px;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3);
        backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px);
        width: 450px; max-width: 70vw; max-height: 65vh; min-width: 260px; min-height: 100px;
        display: flex; flex-direction: column; overflow: hidden; resize: both;
        font: 12px/1.4 system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
        transition: width 0.3s cubic-bezier(0.25, 0.8, 0.25, 1), height 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
    }
    #xjtu-ppt-helper.collapsed {
        width: 200px !important; min-width: 200px !important; min-height: 34px !important;
        height: 34px !important; resize: none; overflow: hidden;
    }
    #xjtu-ppt-helper .hd {
        display: flex; align-items: center; justify-content: space-between;
        padding: 6px 12px; border-bottom: 1px solid var(--xjtu-border); cursor: move;
        user-select: none; -webkit-user-select: none; background: rgba(0,0,0,0.2);
    }
    #xjtu-ppt-helper .title { font-size: 13px; font-weight: 600; letter-spacing: 0.5px; }
    #xjtu-ppt-helper .actions a {
        color: var(--xjtu-muted); text-decoration: none; margin-left: 10px; font-weight: bold; transition: color 0.2s;
        display: inline-block; width: 14px; text-align: center; font-size: 13px;
    }
    #xjtu-ppt-helper .actions a:hover { color: #fff; }
    #xjtu-ppt-helper .toolbar {
        padding: 6px 12px; border-bottom: 1px dashed var(--xjtu-border); background: rgba(255,255,255,0.02);
        display: flex; flex-wrap: wrap; gap: 6px; align-items: center;
    }
    #xjtu-ppt-helper .toolbar .btn, #xjtu-ppt-helper .links .btn {
        display: inline-block; background: var(--xjtu-hover); padding: 3px 8px;
        border-radius: 4px; border: 1px solid transparent; transition: all 0.2s; font-size: 11.5px;
    }
    #xjtu-ppt-helper .toolbar .btn:hover, #xjtu-ppt-helper .links .btn:hover {
        background: rgba(255,255,255,0.2); border-color: rgba(255,255,255,0.3);
    }
    #xjtu-ppt-helper .bd { padding: 0 12px 6px 12px; overflow-y: auto; flex-grow: 1; }
    #xjtu-ppt-helper .bd::-webkit-scrollbar { width: 4px; }
    #xjtu-ppt-helper .bd::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.3); border-radius: 2px; }
    #xjtu-ppt-helper .row {
        padding: 6px 1;
        border-bottom: 1.5px solid rgba(255,255,255,1);
    }
    #xjtu-ppt-helper .row:last-child { border-bottom: none; }
    #xjtu-ppt-helper .name { font-weight: 500; font-size: 12.5px; margin-bottom: 4px; color: #fff; word-break: break-all; }
    #xjtu-ppt-helper .links { display: flex; gap: 6px; }
    #xjtu-ppt-helper a { color: var(--xjtu-accent); text-decoration: none; }
    #xjtu-ppt-helper .muted { color: var(--xjtu-muted); font-size: 11px; }
    `;
    const style = document.createElement('style'); style.textContent = css; document.documentElement.append(style);

    // ---------- state ----------
    let currentActivityId = null;
    let isPanelClosedByUser = false;
    /** @type {{name:string, uploadId:number|null, refId:number|null, urlDirect:string|null, urlRef:string|null, urlAliyun:string|null}[]} */
    let items = [];

    // ---------- panel ----------
    function ensurePanel() {
        if (isPanelClosedByUser) return null;
        let p = qs('#xjtu-ppt-helper');
        if (p) return p;
        p = document.createElement('div');
        p.id = 'xjtu-ppt-helper';
        p.innerHTML = `
      <div class="hd" id="xjtu-ppt-handle">
        <div class="title">📄 XJTU PPT Helper</div>
        <div class="actions">
          <a href="javascript:void 0" id="xjtu-ppt-collapse" title="折叠/展开">—</a>
          <a href="javascript:void 0" id="xjtu-ppt-close" title="关闭">✕</a>
        </div>
      </div>
      <div class="toolbar">
        <a class="btn" href="javascript:void 0" id="xjtu-download-all">📥 一键下载</a>
        <a class="btn" href="javascript:void 0" id="xjtu-copy-csv">📋 复制 CSV</a>
        <a class="btn" href="javascript:void 0" id="xjtu-copy-text">📋 复制文本</a>
        <span class="muted" id="xjtu-status" style="width: 100%; display: block; margin-top: 2px;"></span>
      </div>
      <div class="bd">
        <div id="xjtu-ppt-body">正在截获底层网络请求或检索活动数据...</div>
      </div>
    `;
        document.body.appendChild(p);

        p.querySelector('#xjtu-ppt-close').onclick = () => {
            isPanelClosedByUser = true;
            p.remove();
        };
        p.querySelector('#xjtu-ppt-collapse').onclick = (e) => {
            const isCollapsed = p.classList.toggle('collapsed');
            e.target.textContent = isCollapsed ? '✚' : '—';
        };

        makeDraggable(p, p.querySelector('#xjtu-ppt-handle'));

        p.querySelector('#xjtu-copy-csv').onclick = () => copyAll('csv');
        p.querySelector('#xjtu-copy-text').onclick = () => copyAll('text');
        p.querySelector('#xjtu-download-all').onclick = downloadAll;

        return p;
    }

    function makeDraggable(el, handle) {
        let startX=0, startY=0, startLeft=0, startTop=0, dragging=false;

        const pt = (e) => e.touches?.[0] ? {x:e.touches[0].clientX, y:e.touches[0].clientY} : {x:e.clientX, y:e.clientY};

        const onDown = (e) => {
            if(e.target.tagName.toLowerCase() === 'a') return;
            e.preventDefault();
            const p = pt(e);
            startX = p.x; startY = p.y;
            startLeft = el.offsetLeft; startTop = el.offsetTop;
            dragging = true;
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
            document.addEventListener('touchmove', onMove, {passive:false});
            document.addEventListener('touchend', onUp);
            el.style.right = 'auto'; el.style.bottom = 'auto';
        };
        const onMove = (e) => {
            if (!dragging) return;
            e.preventDefault();
            const p = pt(e);
            const dx = p.x - startX, dy = p.y - startY;
            el.style.left = Math.max(0, startLeft + dx) + 'px';
            el.style.top  = Math.max(0, startTop  + dy) + 'px';
        };
        const onUp = () => {
            dragging = false;
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            document.removeEventListener('touchmove', onMove);
            document.removeEventListener('touchend', onUp);
        };

        handle.addEventListener('mousedown', onDown);
        handle.addEventListener('touchstart', onDown, {passive:false});
    }

    function setStatus(msg) {
        const el = qs('#xjtu-status');
        if (el) el.textContent = msg;
    }

    function renderList() {
        if (isPanelClosedByUser) return;
        const panel = ensurePanel();
        if (!panel) return;

        const body = panel.querySelector('#xjtu-ppt-body');
        if (!items.length) {
            body.innerHTML = '<span class="muted">未在此页面探测到课件附件。</span>';
            return;
        }
        body.innerHTML = '';
        for (const it of items) {
            const row = document.createElement('div');
            row.className = 'row';
            row.innerHTML = `
        <div class="name">${it.name} ${it.urlDirect ? '' : '<span class="muted">（缺失底层数据，将使用引用链接）</span>'}</div>
        <div class="links">
          ${it.urlDirect ? `<a class="btn" href="${it.urlDirect}" target="_blank" title="同域直链下载">⬇️ 下载</a>` : ''}
          ${(!it.urlDirect && it.urlRef) ? `<a class="btn" href="${it.urlRef}" target="_blank" title="引用下载（备用）">🧩 引用下载</a>` : ''}
          ${it.urlAliyun ? `<a class="btn" href="${it.urlAliyun}" target="_blank" title="阿里云 WebOffice 预览">🖥 网页预览</a>` : ''}
        </div>
      `;
            body.appendChild(row);
        }
        setStatus(`已解析 ${items.length} 个附件源`);
    }

    async function downloadAll() {
        if (!items.length) { setStatus('当前上下文无解析文件'); return; }
        setStatus(`启动批量下载流水线，共 ${items.length} 个文件...`);

        for (let i = 0; i < items.length; i++) {
            const it = items[i];
            const targetUrl = it.urlDirect || it.urlRef;
            if (!targetUrl) continue;

            setStatus(`[${i+1}/${items.length}] 正在下发请求: ${it.name}`);

            const a = document.createElement('a');
            a.href = targetUrl;
            a.download = it.name || 'download';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            await sleep(600);
        }
        setStatus('批量下载指令下发完成');
    }

    // ---------- data adapters ----------
    function upsertFromUploadReferences(activityId, refs=[]) {
        const map = new Map(items.map(it => [String(it.uploadId||'r:'+it.refId), it]));
        for (const ref of refs) {
            const name = ref.name || ref.upload?.name || '未命名';
            const uploadId = ref.upload_id || ref.upload?.id || null;
            const refId = ref.id ?? null;
            const itKey = String(uploadId ?? `r:${refId ?? Math.random()}`);
            const urlDirect = uploadId ? `/api/uploads/${uploadId}/blob` : null;
            const urlRef = refId ? `/api/uploads/reference/${refId}/blob` : null;
            const urlAliyun = uploadId ? `/api/uploads/${uploadId}/preview/aliyun?preview=true&refer_id=${activityId}&refer_type=learning_activity` : null;

            const base = map.get(itKey) || {};
            map.set(itKey, { ...base, name, uploadId, refId, urlDirect, urlRef, urlAliyun });
        }
        items = Array.from(map.values());
        renderList();
    }

    function upsertFromUploads(activityId, uploads=[]) {
        const map = new Map(items.map(it => [String(it.uploadId||'r:'+it.refId), it]));
        for (const u of uploads) {
            const name = u.name || '未命名';
            const uploadId = u.id || null;
            const refId = u.reference_id || null;
            const itKey = String(uploadId ?? `r:${refId ?? Math.random()}`);
            const urlDirect = uploadId ? `/api/uploads/${uploadId}/blob` : null;
            const urlRef = refId ? `/api/uploads/reference/${refId}/blob` : null;
            const urlAliyun = uploadId ? `/api/uploads/${uploadId}/preview/aliyun?preview=true&refer_id=${activityId}&refer_type=learning_activity` : null;
            const base = map.get(itKey) || {};
            map.set(itKey, { ...base, name, uploadId, refId, urlDirect, urlRef, urlAliyun });
        }
        items = Array.from(map.values());
        renderList();
    }

    async function copyAll(kind) {
        if (!items.length) { setStatus('当前缓冲无数据可复制'); return; }
        let text = '';
        if (kind === 'csv') {
            const header = ['name','upload_id','ref_id','direct_url','ref_url','aliyun_preview'];
            text += header.map(csvEscape).join(',') + '\n';
            for (const it of items) {
                text += [
                    csvEscape(it.name), csvEscape(it.uploadId ?? ''), csvEscape(it.refId ?? ''),
                    csvEscape(it.urlDirect ?? ''), csvEscape(it.urlRef ?? ''), csvEscape(it.urlAliyun ?? '')
                ].join(',') + '\n';
            }
        } else {
            text = items.map(it => `${it.name}\t${it.urlDirect || it.urlRef || ''}`).join('\n');
        }
        try {
            if (navigator.clipboard?.writeText) { await navigator.clipboard.writeText(text); }
            else {
                const ta = document.createElement('textarea'); ta.value = text; ta.style.position='fixed'; ta.style.opacity='0';
                document.body.appendChild(ta); ta.focus(); ta.select(); document.execCommand('copy'); document.body.removeChild(ta);
            }
            setStatus(`导出指令执行完毕: 已写入${kind === 'csv' ? 'CSV' : '纯文本'}剪贴板`);
        } catch (e) { setStatus('复制异常：' + e.message); }
    }

    function getActivityIdHeuristics() {
        const h = location.hash || ''; const m = h.match(/#\/(\d+)/); if (m) return m[1];
        const p = location.pathname.match(/(\d+)(?:\/?$)/); return p ? p[1] : null;
    }

    async function tryFetchCommonEndpoints() {
        if (isPanelClosedByUser) return;
        currentActivityId = getActivityIdHeuristics();
        if (!currentActivityId) return;

        try {
            const r1 = await fetch(`/api/activities/${currentActivityId}/upload_references`, { credentials: 'same-origin' });
            if (r1.ok) {
                const d = await r1.json().catch(()=>null);
                if (d?.references?.length) { upsertFromUploadReferences(currentActivityId, d.references); setStatus(`API: upload_references 主动探测就绪`); return; }
            }
        } catch {}
        try {
            const r2 = await fetch(`/api/activities/${currentActivityId}`, { credentials: 'same-origin' });
            if (r2.ok) {
                const d = await r2.json().catch(()=>null);
                if (d?.uploads?.length) { upsertFromUploads(currentActivityId, d.uploads); setStatus(`API: activities.uploads 主动探测就绪`); return; }
            }
        } catch {}
    }

    function hookNetwork() {
        const _fetch = window.fetch;
        window.fetch = async (...args) => {
            const res = await _fetch(...args);
            try {
                const url = String(args[0]?.url || args[0] || '');
                if (/\/api\/activities\/\d+\/upload_references/.test(url)) {
                    const cloned = res.clone(); const d = await cloned.json().catch(()=>null);
                    if (d?.references) {
                        const aid = (url.match(/\/api\/activities\/(\d+)\/upload_references/)||[])[1] || getActivityIdHeuristics();
                        upsertFromUploadReferences(aid, d.references); setStatus('HOOK: upload_references 捕获');
                    }
                } else if (/\/api\/activities\/\d+(\?|$)/.test(url)) {
                    const cloned = res.clone(); const d = await cloned.json().catch(()=>null);
                    if (d?.uploads) {
                        const aid = (url.match(/\/api\/activities\/(\d+)/)||[])[1] || getActivityIdHeuristics();
                        upsertFromUploads(aid, d.uploads); setStatus('HOOK: activities.uploads 捕获');
                    }
                }
            } catch {}
            return res;
        };

        const _open = XMLHttpRequest.prototype.open;
        const _send = XMLHttpRequest.prototype.send;
        XMLHttpRequest.prototype.open = function(m, u, ...rest){ this._xjtu_url = String(u || ''); return _open.call(this, m, u, ...rest); };
        XMLHttpRequest.prototype.send = function(...rest){
            this.addEventListener('load', function(){
                try {
                    const url = this._xjtu_url || this.responseURL || '';
                    if (this.status >= 200 && this.status < 300 && this.responseType === '' || this.responseType === 'text') {
                        if (/\/api\/activities\/\d+\/upload_references/.test(url)) {
                            const d = json(this.responseText);
                            if (d?.references) {
                                const aid = (url.match(/\/api\/activities\/(\d+)\/upload_references/)||[])[1] || getActivityIdHeuristics();
                                upsertFromUploadReferences(aid, d.references); setStatus('HOOK: upload_references (XHR) 捕获');
                            }
                        } else if (/\/api\/activities\/\d+(\?|$)/.test(url)) {
                            const d = json(this.responseText);
                            if (d?.uploads) {
                                const aid = (url.match(/\/api\/activities\/(\d+)/)||[])[1] || getActivityIdHeuristics();
                                upsertFromUploads(aid, d.uploads); setStatus('HOOK: activities.uploads (XHR) 捕获');
                            }
                        }
                    }
                } catch {}
            });
            return _send.apply(this, rest);
        };
    }

    async function boot() {
        ensurePanel();
        hookNetwork();
        await tryFetchCommonEndpoints();
        window.addEventListener('hashchange', () => setTimeout(tryFetchCommonEndpoints, 200));
    }

    boot();
})();