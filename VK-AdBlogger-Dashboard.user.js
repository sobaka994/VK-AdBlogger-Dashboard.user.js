// ==UserScript==
// @name         VK AdBlogger — Таблица, память, аналитика, Excel
// @namespace    vk-adblogger-local-dashboard
// @version      15.0
// @description  Локальная память, продолжение загрузки, аналитика по публикациям и Excel
// @match        https://adblogger.vk.ru/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @require      https://cdn.jsdelivr.net/npm/chart.js@4.4.8/dist/chart.umd.min.js
// @run-at       document-idle
// @grant        none
// @noframes
// ==/UserScript==

(async function () {
    'use strict';

    if (document.getElementById('vkab15')) return;

    const CFG = {
        list: '/api/v1/author/soco/placements',
        detail: '/api/v1/author/ao/placement/',
        limit: 50,
        pause: 400,
        timeout: 30000,
        priceDivisor: 100,
        rateDivisor: 1000,
        rewardDivisor: 1,
        db: 'vk-adblogger-dashboard-v15',
        lastProfile: 'vk-adblogger-dashboard-v15-profile'
    };

    const CORE = [
        'clickCount', 'totalOrderCount', 'buyoutCount',
        'inProgressOrderCount', 'predictedReward', 'factReward'
    ];

    const FIELDS = [...CORE, 'goodsSold', 'cpm100'];

    const METRICS = [
        ['clickCount', 'Клики'],
        ['totalOrderCount', 'Заказы всего'],
        ['inProgressOrderCount', 'Заказы в обработке'],
        ['buyoutCount', 'Выкуплено'],
        ['other', 'Остальные заказы — расчёт'],
        ['buyoutRate', 'Доля выкупа, %'],
        ['conversion', 'Заказы / клики, %'],
        ['epc', 'Факт на клик, ₽'],
        ['predictedReward', 'Прогноз, ₽'],
        ['factReward', 'Факт, ₽']
    ];

    const LABEL = Object.fromEntries(METRICS);
    const numericSort = new Set([
        ...METRICS.map(x => x[0]), 'price', 'rate', 'number'
    ]);

    let db, S, profile = '';
    let busy = false;
    let controller = null;
    let confirmedProfile = '';
    let storageError = '';
    let saveChain = Promise.resolve();
    let charts = [];
    let chartTimer;
    let message = 'Загрузка локальной памяти…';

    const host = document.createElement('div');
    host.id = 'vkab15';
    document.body.appendChild(host);
    const root = host.attachShadow({ mode: 'open' });
    const $ = sel => root.querySelector(sel);

    const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;',
        '"': '&quot;', "'": '&#39;'
    })[c]);

    const nf = new Intl.NumberFormat('ru-RU', {
        maximumFractionDigits: 2
    });

    function n(v) {
        if (v === null || v === undefined || v === '') return null;
        if (typeof v !== 'number' && typeof v !== 'string') return null;
        const x = Number(typeof v === 'string'
            ? v.replace(/[\s\u00a0\u202f]/g, '').replace(',', '.')
            : v);
        return Number.isFinite(x) ? x : null;
    }

    const fmt = v => n(v) === null ? '—' : nf.format(n(v));
    const rub = v => n(v) === null ? '—' : fmt(v) + ' ₽';
    const pct = v => n(v) === null ? '—' : fmt(v) + '%';

    function date(v) {
        if (!v || String(v).startsWith('0001-')) return null;
        const d = new Date(v);
        return Number.isFinite(d.getTime()) ? d : null;
    }

    function localDay(d) {
        return [
            d.getFullYear(),
            String(d.getMonth() + 1).padStart(2, '0'),
            String(d.getDate()).padStart(2, '0')
        ].join('-');
    }

    const dateText = v => date(v)?.toLocaleDateString('ru-RU') || '—';
    const timeText = v => date(v)?.toLocaleString('ru-RU') || '—';

    function url(v) {
        if (!v) return '';
        try {
            const u = new URL(v, location.origin);
            return /^https?:$/.test(u.protocol) ? u.href : '';
        } catch {
            return '';
        }
    }

    function link(href, text) {
        const target = url(href);
        return target
            ? `<a href="${esc(target)}" target="_blank"
                 rel="noopener noreferrer">${esc(text)}</a>`
            : esc(text);
    }

    function orderUrl(o) {
        return location.origin + '/app/author-placements/ad-products/' +
            encodeURIComponent(o.uuid);
    }

    function status(o) {
        const code = String(o.status?.code || '');
        const short = code.replace(/^SummarizedPlacementStatusCode/, '');
        const names = {
            AOInWork: 'В работе',
            InWork: 'В работе',
            AOInModeration: 'На модерации',
            InModeration: 'На модерации',
            Moderation: 'На модерации',
            OnModeration: 'На модерации',
            PendingModeration: 'Ожидает модерации',
            WaitingForModeration: 'Ожидает модерации',
            Scheduled: 'Запланирована',
            InProgress: 'Опубликована',
            Succeeded: 'Завершена',
            Canceled: 'Отменена',
            Cancelled: 'Отменена',
            Rejected: 'Отклонена',
            Draft: 'Черновик',
            Accepted: 'Принята'
        };
        if (names[short]) return names[short];
        if (/Canceled|Cancelled/.test(code)) return 'Отменена';
        if (/Rejected/.test(code)) return 'Отклонена';
        return code ? 'Другой статус' : 'Не указан';
    }

    function category(o) {
        const code = String(o.status?.code || '');
        if (/Canceled|Cancelled|Rejected/.test(code)) return 'canceled';
        const d = date(o.publishDateTime);
        if (!d || d.getTime() > Date.now()) return 'unpublished';
        if (url(o.publishedUrl) ||
            /(?:InProgress|Succeeded)$/.test(code)) return 'published';
        return 'unknown';
    }

    const categoryName = k => ({
        published: 'Опубликованные',
        unpublished: 'Без публикации',
        canceled: 'Отклоненные',
        unknown: 'Проверить статус'
    })[k] || k;

    function val(o, key) {
        if (key === 'name') return o.product?.name || '';
        if (key === 'community') return o.community?.title || '';
        if (key === 'status') return status(o);
        if (key === 'number') return n(o.number);
        if (key === 'published') return date(o.publishDateTime)?.getTime() ?? null;
        if (key === 'price') {
            const x = n(o.product?.price100);
            return x === null ? null : x / CFG.priceDivisor;
        }
        if (key === 'rate') {
            const x = n(o.product?.rate1000);
            return x === null ? null : x / CFG.rateDivisor;
        }
        if (key === 'refusals') return null;

        const stats = o.stats || {};
        const total = n(stats.totalOrderCount);
        const bought = n(stats.buyoutCount);
        const processing = n(stats.inProgressOrderCount);
        const clicks = n(stats.clickCount);

        if (key === 'other') {
            if ([total, bought, processing].some(x => x === null)) return null;
            const x = total - bought - processing;
            return x >= 0 ? x : null;
        }
        if (key === 'buyoutRate') {
            return total > 0 && bought !== null && bought >= 0 && bought <= total
                ? bought / total * 100 : null;
        }
        if (key === 'conversion') {
            return clicks > 0 && total !== null
                ? total / clicks * 100 : null;
        }
        if (key === 'epc') {
            const reward = val(o, 'factReward');
            return clicks > 0 && reward !== null ? reward / clicks : null;
        }
        const x = n(stats[key]);
        return x === null ? null
            : ['predictedReward', 'factReward'].includes(key)
                ? x / CFG.rewardDivisor : x;
    }

    function metricText(key, v) {
        if (['buyoutRate', 'conversion'].includes(key)) return pct(v);
        if (['epc', 'factReward', 'predictedReward'].includes(key)) return rub(v);
        return fmt(v);
    }

    function allRows() {
        return Object.values(S.rows);
    }

    function completeStats(o) {
        return CORE.every(k => n(o.stats?.[k]) !== null);
    }

    function blankState() {
        return {
            version: 15,
            rows: {},
            scan: {
                offset: 0, total: null, done: false,
                complete: false, seen: [], started: null, updated: null
            },
            queue: null,
            checked: {},
            saved: null,
            ui: {
                query: '', category: '', sort: 'published',
                dir: -1, hideZero: false, page: 1
            },
            analytics: {
                period: 'all', from: '', to: '', group: 'month',
                metric: 'predictedReward', community: ''
            }
        };
    }

    // ---------- IndexedDB ----------

    function openDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(CFG.db, 1);
            req.onupgradeneeded = () => {
                if (!req.result.objectStoreNames.contains('profiles')) {
                    req.result.createObjectStore('profiles');
                }
            };
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    function readProfile(name) {
        return new Promise((resolve, reject) => {
            const tx = db.transaction('profiles', 'readonly');
            const req = tx.objectStore('profiles').get(name);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        });
    }

    function writeProfile(name, data) {
        return new Promise((resolve, reject) => {
            const tx = db.transaction('profiles', 'readwrite');
            tx.objectStore('profiles').put(data, name);
            tx.oncomplete = () => resolve();
            tx.onabort = () => reject(tx.error || new Error('Запись отменена'));
            tx.onerror = () => reject(tx.error || new Error('Ошибка IndexedDB'));
        });
    }

    function profileNames() {
        return new Promise((resolve, reject) => {
            const req = db.transaction('profiles', 'readonly')
                .objectStore('profiles').getAllKeys();
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }

    function save() {
        S.saved = new Date().toISOString();
        const snapshot = structuredClone(S);
        const name = profile;

        // Снимки записываются строго по порядку.
        const task = saveChain.then(() => writeProfile(name, snapshot));
        saveChain = task.catch(() => {});

        return task.then(() => {
            storageError = '';
            $('#memory').textContent =
                `Память: ${name} · сохранено ${timeText(snapshot.saved)}`;
        }).catch(error => {
            storageError =
                'Не удалось сохранить данные. Не закрывайте страницу. ' +
                (error?.message || error);
            $('#memory').textContent = storageError;
            $('#memory').classList.add('error');
            controller?.abort();
            throw new Error(storageError);
        });
    }

    function saveUI() {
        save().catch(error => {
            message = error.message;
            render();
        });
    }

    async function refreshProfileMenu() {
        const names = await profileNames();
        if (!names.includes(profile)) names.push(profile);
        $('#profile').innerHTML = names.sort().map(name =>
            `<option value="${esc(name)}">${esc(name)}</option>`
        ).join('');
        $('#profile').value = profile;
    }

    async function switchProfile(name) {
        await saveChain;
        profile = name;
        S = await readProfile(name) || blankState();
        if (S.version !== 15) throw new Error('Несовместимый формат памяти.');
        confirmedProfile = '';
        localStorage.setItem(CFG.lastProfile, name);
        message = allRows().length
            ? 'Данные восстановлены из памяти. Сетевые запросы не запускались.'
            : 'Профиль пуст. Нажмите «Загрузить всё».';
        await refreshProfileMenu();
        syncControls();
        render();
        scheduleCharts();
    }

    // В память не сохраняем вложения, подписанные URL, ERIR-токены,
    // cookies или заголовки авторизации.
    function ingest(raw, source) {
        if (!raw?.uuid) return;

        const old = S.rows[raw.uuid] || {};
        const now = new Date().toISOString();
        const stats = { ...(old.stats || {}) };
        const metricDates = { ...(old.metricDates || {}) };

        for (const k of FIELDS) {
            if (Object.prototype.hasOwnProperty.call(raw.stats || {}, k)) {
                stats[k] = n(raw.stats[k]);
                metricDates[k] = now;
            }
        }

        function part(key, fields) {
            const result = { ...(old[key] || {}) };
            for (const f of fields) {
                if (Object.prototype.hasOwnProperty.call(raw[key] || {}, f)) {
                    result[f] = raw[key][f];
                }
            }
            return result;
        }

        const o = {
            ...old,
            uuid: raw.uuid,
            stats,
            metricDates,
            product: part('product', ['name', 'externalURL', 'price100', 'rate1000']),
            community: part('community', ['communityID', 'title', 'link']),
            status: part('status', ['code', 'expiresAt']),
            campaignPeriod: part('campaignPeriod', ['from', 'to']),
            index: old.index ?? Object.keys(S.rows).length,
            listUpdated: source === 'Список' ? now : old.listUpdated || null,
            detailUpdated: source === 'Карточка' ? now : old.detailUpdated || null,
            source,
            error: source === 'Карточка' ? '' : old.error || ''
        };

        for (const k of [
            'number', 'format', 'publishDateTime', 'publishedUrl',
            'message', 'statusComment', 'fixedIncomeStatus', 'sum100'
        ]) {
            if (Object.prototype.hasOwnProperty.call(raw, k)) o[k] = raw[k];
        }

        S.rows[o.uuid] = o;
    }

    // ---------- Интерфейс ----------

    root.innerHTML = `
    <style>
        :host{all:initial;font:13px -apple-system,BlinkMacSystemFont,"Segoe UI",Arial;
              color:#213047}
        *{box-sizing:border-box}
        [hidden]{display:none!important}
        button,input,select{font:inherit;color:inherit}
        button,input,select{
            border:1px solid #dce3ee;border-radius:8px;
            padding:8px 10px;background:#fff
        }
        button{cursor:pointer}button:hover{background:#eef5ff}
        button:disabled{opacity:.45;cursor:default}
        a{color:#087be5;text-decoration:none}a:hover{text-decoration:underline}
        .primary{background:#087bf0;color:#fff;border-color:#087bf0}
        .primary:hover{background:#0066d0}
        .green{background:#eaf7ef;color:#207849}
        .grow{flex:1}.muted{color:#7b8798;font-size:11px}
        .error{color:#b63f39!important}
        #launch{position:fixed;right:18px;bottom:18px;z-index:2147483645;
                background:#087bf0;color:white;box-shadow:0 3px 18px #0003}
        .panel{position:fixed;inset:12px;background:#fff;z-index:2147483646;
               border:1px solid #dce3ee;border-radius:14px;
               box-shadow:0 10px 45px #0003;display:flex;flex-direction:column;
               overflow:hidden}
        #analytics{inset:22px;z-index:2147483647;background:#f5f7fb}
        .bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;
             padding:10px 14px;border-bottom:1px solid #e8edf4}
        .head{background:#f6f8fc}.head strong{font-size:17px}
        #forecast{font-size:16px;font-weight:700}
        #notice{padding:9px 14px;line-height:1.7;border-bottom:1px solid #e8edf4}
        #memory{font-size:11px;color:#78879b}
        progress{width:100%;height:6px;accent-color:#087bf0}
        .scroll{overflow:auto;flex:1;min-height:0}
        table{width:100%;border-collapse:separate;border-spacing:0;min-width:1120px}
        th{position:sticky;top:0;z-index:2;background:#f2f5fa;
           text-align:left;color:#60718a;font-size:12px;padding:10px}
        th button{background:transparent;border:0;padding:0;color:inherit}
        td{padding:10px;border-bottom:1px solid #edf1f6;vertical-align:top;
           line-height:1.5}
        tbody tr:hover{background:#f5f9ff}
        .product{max-width:300px;font-weight:600;display:-webkit-box;
                 -webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
        .metric{display:flex;justify-content:space-between;gap:14px;white-space:nowrap}
        .metric b{font-variant-numeric:tabular-nums;font-weight:600}
        .badge{display:inline-block;padding:3px 7px;border-radius:6px;
               background:#edf2f8;font-size:11px;max-width:155px}
        .published{background:#e8f4ff;color:#176caa}
        .canceled{background:#fff0ef;color:#ac4540}
        .money{text-align:right;white-space:nowrap;font-weight:600}
        .footer{padding:8px 14px;border-top:1px solid #e8edf4;
                display:flex;align-items:center;gap:10px;flex-wrap:wrap}
        #dash{padding:16px}
        .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));
              gap:12px;margin:12px 0}
        .card{background:#fff;border:1px solid #e4eaf3;border-radius:12px;padding:16px}
        .card .value{font-size:24px;font-weight:700;margin:8px 0;color:#183759}
        .grid{display:grid;grid-template-columns:2fr 1fr;gap:14px}
        .wide{grid-column:1/-1}
        .chart{height:310px;position:relative;width:100%}
        .card h3{font-size:14px;margin:0 0 14px}
        .note{background:#eaf2ff;color:#36587f;padding:12px;border-radius:9px;
              line-height:1.55}
        #quality{margin:12px 0;line-height:1.6;color:#60718a}
        #analysisTable table{min-width:850px}
        #analysisTable th{position:static}
        label{display:inline-flex;align-items:center;gap:5px}
        input[type=checkbox]{accent-color:#087bf0}
        @media(max-width:1000px){.grid{grid-template-columns:1fr}.wide{grid-column:auto}}
    </style>

    <button id="launch">📊 VK Аналитика</button>

    <section id="panel" class="panel" hidden>
        <div class="bar head">
            <strong>Заявки · VK AdBlogger</strong>
            <span class="grow"></span><span id="forecast">Прогноз: —</span>
            <button id="showAnalytics">Аналитика</button>
            <button id="close">К сайту ✕</button>
        </div>
        <div class="bar">
            <span>Память:</span><select id="profile"></select>
            <button id="newProfile">Новый профиль</button>
            <button id="clear">Очистить профиль</button>
            <span class="grow"></span>
            <button id="excel" class="green">Excel</button>
        </div>
        <div class="bar">
            <button id="load" class="primary">Загрузить всё / продолжить</button>
            <button id="rescan">Обновить список с начала</button>
            <button id="details">Уточнить показатели</button>
            <button id="resume" hidden>Продолжить проверку</button>
            <button id="retry">Повторить ошибки</button>
            <button id="refresh">Обновить все карточки</button>
            <button id="stop" disabled>Остановить</button>
        </div>
        <div class="bar">
            <input id="search" placeholder="Товар, номер, сообщество">
            <select id="category">
                <option value="">Все категории</option>
                <option value="published">Опубликованные</option>
                <option value="unpublished">Без публикации</option>
                <option value="canceled">Отклонённые / отменённые</option>
                <option value="unknown">Проверить статус</option>
            </select>
            <select id="sort">
                <option value="published">Дата публикации</option>
                <option value="name">Товар</option>
                <option value="number">Номер заявки</option>
                <option value="price">Цена</option>
                <option value="rate">Процент вознаграждения</option>
                <option value="community">Сообщество</option>
                <option value="status">Статус</option>
                ${METRICS.map(([k,t])=>`<option value="${k}">${esc(t)}</option>`).join('')}
                <option value="">Исходный порядок</option>
            </select>
            <button id="direction">↓</button>
            <label><input id="zeros" type="checkbox">Скрыть нули</label>
            <button id="reset">Сброс фильтров</button>
        </div>
        <div id="notice">
            <div id="message"></div>
            <div id="counts"></div>
            <div id="queueText"></div>
            <progress id="progress" max="1" value="0"></progress>
            <div id="memory"></div>
        </div>
        <div class="scroll">
            <table>
                <thead><tr>
                    <th><button data-sort="name">Товар / заявка ↕</button></th>
                    <th>Сообщество / статус</th>
                    <th><button data-sort="published">Публикация ↕</button></th>
                    <th><button data-sort="clickCount">Клики ↕</button></th>
                    <th><button data-sort="totalOrderCount">Заказы ↕</button></th>
                    <th>Эффективность</th>
                    <th><button data-sort="predictedReward">Прогноз ↕</button></th>
                    <th><button data-sort="factReward">Факт ↕</button></th>
                </tr></thead>
                <tbody id="body"></tbody>
            </table>
        </div>
        <div class="footer">
            <button id="prev">←</button><span id="pageInfo"></span>
            <button id="next">→</button>
            <span class="grow"></span>
            <span class="muted">
                «Остальные» — расчёт, не подтверждённые отказы.
                «Факт» не означает доступность денег к выводу.
            </span>
        </div>
    </section>

    <section id="analytics" class="panel" hidden>
        <div class="bar head">
            <strong>Аналитика публикаций</strong>
            <span class="grow"></span>
            <button id="aExcel" class="green">Excel</button>
            <button id="aClose">Закрыть ✕</button>
        </div>
        <div class="bar">
            <select id="period">
                <option value="today">Сегодня</option>
                <option value="week">Последние 7 дней</option>
                <option value="month">Последние 30 дней</option>
                <option value="all">Всё время</option>
                <option value="custom">Свой период</option>
            </select>
            <input id="from" type="date" title="Дата публикации от">
            <input id="to" type="date" title="Дата публикации до">
            <select id="group">
                <option value="day">Группировать по дням</option>
                <option value="week">По неделям, с понедельника</option>
                <option value="month">По месяцам</option>
            </select>
            <select id="aMetric">
                ${METRICS.map(([k,t])=>`<option value="${k}">${esc(t)}</option>`).join('')}
            </select>
            <select id="aCommunity"><option value="">Все сообщества</option></select>
        </div>
        <div class="scroll" id="dash">
            <div class="note">
                Период относится к <b>дате публикации</b>.
                Показатели накопительные, по последним полученным данным.
                Это не число событий, произошедших за выбранный период.
                Свежие публикации имели меньше времени набрать результаты.
            </div>
            <div id="quality"></div>
            <div id="kpis" class="kpis"></div>
            <div class="grid">
                <div class="card wide">
                    <h3 id="trendTitle">Показатель по публикациям</h3>
                    <div class="chart"><canvas id="trend"></canvas></div>
                </div>
                <div class="card">
                    <h3 id="topTitle">Сообщества · топ-10</h3>
                    <div class="chart"><canvas id="top"></canvas></div>
                </div>
                <div class="card">
                    <h3>Структура заказов</h3>
                    <div class="chart"><canvas id="ordersChart"></canvas></div>
                    <div id="ordersNote" class="muted"></div>
                </div>
                <div class="card wide">
                    <h3>Показатели по периодам публикации</h3>
                    <div id="analysisTable" style="overflow:auto"></div>
                </div>
            </div>
        </div>
    </section>`;

    function visibleRows() {
        let rows = allRows();
        const u = S.ui;
        const q = u.query.trim().toLocaleLowerCase('ru-RU');

        if (q) rows = rows.filter(o =>
            [o.number, o.product?.name, o.community?.title].join(' ')
                .toLocaleLowerCase('ru-RU').includes(q)
        );
        if (u.category) rows = rows.filter(o => category(o) === u.category);
        if (u.hideZero && numericSort.has(u.sort)) {
            rows = rows.filter(o => val(o, u.sort) !== 0);
        }

        rows.sort((a,b) => {
            if (!u.sort) return a.index - b.index;
            const av = val(a,u.sort), bv = val(b,u.sort);
            if (av === null && bv === null) return a.index - b.index;
            if (av === null) return 1;
            if (bv === null) return -1;
            const diff = typeof av === 'number' && typeof bv === 'number'
                ? av-bv : String(av).localeCompare(String(bv),'ru',{numeric:true});
            return diff ? diff*u.dir : a.index-b.index;
        });
        return rows;
    }

    function sum(rows, key) {
        const values = rows.map(o => val(o,key)).filter(x => x !== null);
        return {
            value: values.length ? values.reduce((a,b)=>a+b,0) : null,
            known: values.length,
            missing: rows.length-values.length
        };
    }

    // Отношения считаются только по строкам, где известны оба поля.
    // Не делим суммы с разным составом заявок.
    function aggregate(rows, key) {
        const rules = {
            buyoutRate: ['buyoutCount','totalOrderCount',100],
            conversion: ['totalOrderCount','clickCount',100],
            epc: ['factReward','clickCount',1]
        };
        if (!rules[key]) return sum(rows,key);

        const [a,b,mult] = rules[key];
        const valid = rows.filter(o => {
            const x=val(o,a), y=val(o,b);
            return x !== null && y !== null && x>=0 && y>=0 &&
                (key!=='buyoutRate' || x<=y);
        });
        const numerator = valid.reduce((s,o)=>s+val(o,a),0);
        const denominator = valid.reduce((s,o)=>s+val(o,b),0);
        return {
            value: denominator>0 ? numerator/denominator*mult : null,
            known: valid.length,
            missing: rows.length-valid.length
        };
    }

    function pendingQueue() {
        return S.queue && S.queue.index < S.queue.ids.length;
    }

    function syncControls() {
        const u=S.ui, a=S.analytics;
        $('#search').value=u.query;
        $('#category').value=u.category;
        $('#sort').value=u.sort;
        $('#zeros').checked=u.hideZero;
        $('#period').value=a.period;
        $('#from').value=a.from;
        $('#to').value=a.to;
        $('#group').value=a.group;
        $('#aMetric').value=a.metric;
        $('#from').disabled=$('#to').disabled=a.period!=='custom';
    }

    function render() {
        if (!S) return;
        const rows=allRows(), filtered=visibleRows();
                // Для старых профилей без этой настройки по умолчанию — все строки.
        const savedSize = String(S.ui.pageSize ?? 'all');
        const pageSize = ['50', '100', '200', 'all'].includes(savedSize)
            ? savedSize
            : 'all';

        S.ui.pageSize = pageSize;

        const showAll = pageSize === 'all';
        const limit = showAll
            ? Math.max(filtered.length, 1)
            : Number(pageSize);

        const pages = showAll
            ? 1
            : Math.max(1, Math.ceil(filtered.length / limit));

        S.ui.page = Math.max(1, Math.min(S.ui.page, pages));

        const pageRows = showAll
            ? filtered
            : filtered.slice(
                (S.ui.page - 1) * limit,
                S.ui.page * limit
            );

        $('#pageSize').value = pageSize;

        const forecast=sum(rows.filter(o=>category(o)!=='canceled'),'predictedReward');
        $('#forecast').textContent='Прогноз: '+rub(forecast.value);
        $('#forecast').title=
            `По всем загруженным неотменённым заявкам, независимо от фильтров.\n`+
            `Без значения прогноза: ${forecast.missing}.\n`+
            (S.scan.complete?'Список полный на момент сбора.':'Список неполный: сумма частичная.');

        $('#message').textContent=storageError || message;
        $('#counts').textContent=
            `В памяти: ${rows.length}`+
            (S.scan.total!==null?` · Сервер сообщает: ${S.scan.total}`:'')+
            ` · Список: ${S.scan.complete?'полный':'неполный'}`+
            ` · Отобрано: ${filtered.length}`+
            ` · Все основные поля есть: ${rows.filter(completeStats).length}`;

        $('#memory').textContent=storageError ||
            `Профиль: ${profile} · сохранено ${timeText(S.saved)}`;
        $('#memory').classList.toggle('error',Boolean(storageError));

        const q=S.queue;
        $('#queueText').textContent=q
            ? `Карточки: ${q.index}/${q.ids.length} · Осталось: ${q.ids.length-q.index}`+
              ` · Дополнено: ${q.added} · Без новых полей: ${q.unchanged}`+
              ` · Ошибок: ${q.failed}`
            : 'Проверка карточек запускается только отдельной кнопкой.';

        const progress=q && pendingQueue()
            ? [q.index,q.ids.length]
            : [S.scan.offset,S.scan.total||Math.max(S.scan.offset,1)];
        $('#progress').max=Math.max(progress[1],1);
        $('#progress').value=progress[0];

        for(const id of [
            'profile','newProfile','clear','load','rescan','details',
            'resume','retry','refresh','excel','aExcel'
        ]) $('#'+id).disabled=busy;

        $('#stop').disabled=!busy;
        $('#load').disabled=busy || S.scan.done;
        $('#load').textContent=S.scan.done?'Список загружен':
            S.scan.offset?'Продолжить загрузку страниц':'Загрузить всё';

        const pending=Boolean(pendingQueue());
        $('#resume').hidden=!pending;
        $('#details').disabled=busy||pending||!rows.some(o=>
            !completeStats(o)&&!S.checked[o.uuid]);
        $('#refresh').disabled=busy||pending||!rows.length;
        $('#retry').disabled=busy||pending||!rows.some(o=>o.error);
        $('#excel').disabled=$('#aExcel').disabled=busy||!rows.length;

        $('#direction').textContent=S.ui.dir===-1?'↓ Убывание':'↑ Возрастание';
        $('#zeros').disabled=!numericSort.has(S.ui.sort);
        $('#pageInfo').textContent = showAll
            ? `Показано строк: ${filtered.length} из ${rows.length}`
            : `Страница ${S.ui.page} из ${pages} · Строк: ${filtered.length}`;

        $('#prev').hidden = showAll;
        $('#next').hidden = showAll;

        $('#prev').disabled = S.ui.page <= 1;
        $('#next').disabled = S.ui.page >= pages;

        $('#body').innerHTML=pageRows.length?pageRows.map(o=>{
            const dates=Object.values(o.metricDates||{}).filter(Boolean).sort();
            const freshness=dates.length
                ? `Показатели получены: ${timeText(dates[0])} — ${timeText(dates.at(-1))}`
                : 'Показатели ещё не получены';
            const format=o.format==='PlacementFormatAuthorPost'?'Пост':
                o.format==='PlacementFormatAuthorClip'?'Клип':o.format||'';
            const metric=(name,key,title='')=>
                `<div class="metric" title="${esc(title)}">
                    <span>${esc(name)}</span><b>${metricText(key,val(o,key))}</b>
                 </div>`;
            return `<tr title="${esc(freshness)}">
                <td style="min-width:240px;max-width:330px">
                    <div class="product">${link(o.product?.externalURL,o.product?.name||'Без названия')}</div>
                    <div>${link(orderUrl(o),'№'+(o.number||'—'))}
                        <span class="muted"> · ${rub(val(o,'price'))}
                        · ${pct(val(o,'rate'))}</span></div>
                    <div class="muted">${esc(format)}</div>
                </td>
                <td>
                    ${link(o.community?.link,o.community?.title||'—')}
                    <div><span class="badge ${category(o)}"
                         title="${esc(o.status?.code)}">${esc(status(o))}</span></div>
                    ${o.error?'<div class="error muted">Ошибка проверки карточки</div>':''}
                </td>
                <td>${dateText(o.publishDateTime)}
                    <div class="muted">${o.publishedUrl?link(o.publishedUrl,'Открыть пост'):''}</div>
                    <div class="muted">Срок до ${dateText(o.status?.expiresAt)}</div>
                </td>
                <td style="font-weight:600">${fmt(val(o,'clickCount'))}</td>
                <td style="min-width:185px">
                    ${metric('Всего','totalOrderCount')}
                    ${metric('В обработке','inProgressOrderCount')}
                    ${metric('Выкуплено','buyoutCount')}
                    ${metric('Остальные*','other',
                        'Всего − в обработке − выкуплено. Не подтверждённые отказы.')}
                    <div class="muted" title="Поле отказов в API не подтверждено">Отказы: —</div>
                </td>
                <td style="min-width:175px">
                    ${metric('Доля выкупа','buyoutRate','Выкуплено / всего заказов × 100%')}
                    ${metric('Заказы / клики','conversion','Всего заказов / клики × 100%')}
                    ${metric('Факт на клик','epc','Фактическое вознаграждение / клики')}
                </td>
                <td class="money">${rub(val(o,'predictedReward'))}</td>
                <td class="money">${rub(val(o,'factReward'))}</td>
            </tr>`;
        }).join(''):'<tr><td colspan="8" style="text-align:center;padding:35px">Нет строк. Загрузите данные или измените фильтры.</td></tr>';
    }

    // ---------- Сеть и продолжение ----------

    function stopCheck() {
        if(controller?.signal.aborted) {
            throw new DOMException('Остановлено','AbortError');
        }
    }

    function wait(ms) {
        return new Promise((resolve,reject)=>{
            stopCheck();
            const signal=controller.signal;
            const abort=()=>{
                clearTimeout(timer);
                reject(new DOMException('Остановлено','AbortError'));
            };
            const timer=setTimeout(()=>{
                signal.removeEventListener('abort',abort);
                resolve();
            },ms);
            signal.addEventListener('abort',abort,{once:true});
        });
    }

    async function request(path,body) {
        stopCheck();
        const local=new AbortController();
        const parent=controller.signal;
        let timedOut=false;
        const abort=()=>local.abort();
        parent.addEventListener('abort',abort,{once:true});
        const timer=setTimeout(()=>{
            timedOut=true;
            local.abort();
        },CFG.timeout);

        try {
            const response=await fetch(path,{
                method:body===undefined?'GET':'POST',
                credentials:'same-origin',
                cache:'no-store',
                headers:body===undefined?{Accept:'application/json'}:{
                    Accept:'application/json','Content-Type':'application/json'
                },
                ...(body===undefined?{}:{body:JSON.stringify(body)}),
                signal:local.signal
            });
            if(!response.ok) {
                const e=new Error(
                    response.status===429?'HTTP 429: ограничение сервера. Подождите и продолжите.':
                    [401,403].includes(response.status)?`HTTP ${response.status}: проверьте вход и кабинет.`:
                    `Ошибка HTTP ${response.status}`
                );
                e.status=response.status;
                throw e;
            }
            return await response.json();
        } catch(e) {
            if(timedOut&&!parent.aborted) throw new Error('Тайм-аут: сервер не ответил за 30 секунд.');
            throw e;
        } finally {
            clearTimeout(timer);
            parent.removeEventListener('abort',abort);
        }
    }

    async function run(task) {
        if(busy) return;

        if(confirmedProfile!==profile) {
            if(!confirm(
                `Профиль памяти: «${profile}».\n\n`+
                'Подтвердите, что он соответствует открытому кабинету VK AdBlogger.\n'+
                'Аккаунт автоматически не определяется. Для другого кабинета выберите другой профиль.'
            )) return;
            confirmedProfile=profile;
        }

        busy=true;
        controller=new AbortController();
        render();
        try {
            await save();
            await task();
        } catch(e) {
            message=e.name==='AbortError'
                ? 'Остановлено. Сохранённую загрузку можно продолжить.'
                : e.message;
        } finally {
            busy=false;
            controller=null;
            render();
            scheduleCharts();
        }
    }

    async function loadPages(reset=false) {
        if(reset) {
            S.scan={
                offset:0,total:null,done:false,complete:false,seen:[],
                started:new Date().toISOString(),updated:null
            };
            S.queue=null;
            // Старые карточки не удаляем до окончания нового прохода.
            await save();
        }

        if(S.scan.done) return;
        if(!S.scan.started) S.scan.started=new Date().toISOString();

        while(!S.scan.done) {
            stopCheck();
            const offset=S.scan.offset;
            message=`Загрузка списка: строки с ${offset+1}. Карточки отдельно не запрашиваются.`;
            render();

            const data=await request(CFG.list,{
                pagination:{limit:CFG.limit,offset},filter:{}
            });
            if(!Array.isArray(data?.items)) throw new Error('В ответе нет массива items.');
            if(data.items.some(o=>!o?.uuid)) throw new Error('В странице есть заявки без UUID.');

            const total=n(data.totalCount);
            const seen=new Set(S.scan.seen);
            const newlySeen=data.items.filter(o=>!seen.has(o.uuid)).length;
            if(data.items.length&&!newlySeen) {
                throw new Error('Сервер повторил страницу. Обновите список с начала.');
            }

            if(total!==null) S.scan.total=total;
            if(!data.items.length && S.scan.total!==null && offset<S.scan.total) {
                throw new Error('Пустая страница раньше totalCount. Попробуйте продолжить или обновить список.');
            }

            data.items.forEach(o=>{
                ingest(o,'Список');
                seen.add(o.uuid);
            });
            S.scan.seen=[...seen];
            S.scan.offset=offset+data.items.length;
            S.scan.updated=new Date().toISOString();

            const end=S.scan.total!==null
                ? S.scan.offset>=S.scan.total
                : data.items.length<CFG.limit;

            if(end) {
                S.scan.done=true;
                S.scan.complete=S.scan.total===null || seen.size>=S.scan.total;
                if(S.scan.complete) {
                    for(const id of Object.keys(S.rows)) {
                        if(!seen.has(id)) {
                            delete S.rows[id];
                            delete S.checked[id];
                        }
                    }
                }
            }

            // Строки страницы и новый offset записываются одним снимком.
            await save();
            render();
            if(!S.scan.done) await wait(CFG.pause);
        }

        message=S.scan.complete
            ? 'Список загружен и сохранён. Уточнение карточек — только по вашей команде.'
            : 'Проход завершён, но число уникальных заявок не совпало с totalCount. Обновите список с начала.';
    }

    async function details(mode) {
        if(mode!=='resume') {
            if(pendingQueue()) throw new Error('Сначала продолжите текущую очередь.');
            const ids=allRows().filter(o=>
                mode==='all' ||
                (mode==='errors'?Boolean(o.error):
                    !completeStats(o)&&!S.checked[o.uuid])
            ).map(o=>o.uuid);
            S.queue={ids,index:0,added:0,unchanged:0,failed:0,mode};
            await save();
        }

        const q=S.queue;
        if(!q?.ids.length) {
            message='Нет карточек для проверки.';
            return;
        }

        while(q.index<q.ids.length) {
            stopCheck();
            const id=q.ids[q.index], old=S.rows[id];
            if(!old) {
                q.index++;q.failed++;
                await save();
                continue;
            }

            message=`Проверяю карточку №${old.number||id}: ${q.index+1}/${q.ids.length}.`;
            $('#message').textContent=message;
            const absent=CORE.filter(k=>n(old.stats?.[k])===null);

            try {
                const raw=await request(CFG.detail+encodeURIComponent(id));
                if(raw?.uuid!==id) throw new Error('Ответ карточки имеет неожиданную структуру.');

                ingest(raw,'Карточка');
                const added=absent.some(k=>n(S.rows[id].stats?.[k])!==null);
                if(added) q.added++; else q.unchanged++;
                S.checked[id]={at:new Date().toISOString(),result:'ok'};
            } catch(e) {
                if(e.name==='AbortError'||[401,403,429].includes(e.status)) throw e;
                S.rows[id].error=e.message;
                S.checked[id]={at:new Date().toISOString(),result:'error'};
                q.failed++;
            }

            q.index++;
            // Карточка и позиция очереди сохраняются вместе.
            await save();
            render();
            if(q.index<q.ids.length) await wait(CFG.pause);
        }

        message='Проверка завершена и сохранена. Неполученные поля не заменены нулями.';
    }

    // ---------- Аналитика по дате публикации ----------

    function periodBounds() {
        const a=S.analytics;
        if(a.period==='custom') return [a.from,a.to];
        if(a.period==='all') return ['',''];

        const end=new Date();
        const start=new Date(end.getFullYear(),end.getMonth(),end.getDate());
        start.setDate(start.getDate()-(a.period==='week'?6:a.period==='month'?29:0));
        return [localDay(start),localDay(end)];
    }

    function analyticsRows() {
        const [from,to]=periodBounds();
        if(from&&to&&from>to) return [];
        return allRows().filter(o=>{
            if(category(o)!=='published') return false;
            const d=date(o.publishDateTime);
            if(!d) return false;
            const day=localDay(d);
            return (!from||day>=from)&&(!to||day<=to)&&
                (!S.analytics.community||
                    String(o.community?.communityID)===S.analytics.community);
        });
    }

    function bucket(o) {
        const d=date(o.publishDateTime);
        if(!d) return '';
        if(S.analytics.group==='month') return localDay(d).slice(0,7);
        if(S.analytics.group==='week') {
            d.setDate(d.getDate()-((d.getDay()+6)%7));
        }
        return localDay(d);
    }

    function grouped(rows,keyFn) {
        const groups=new Map();
        for(const o of rows) {
            const key=keyFn(o);
            if(!groups.has(key)) groups.set(key,[]);
            groups.get(key).push(o);
        }
        return [...groups.entries()];
    }

    function destroyCharts() {
        charts.forEach(c=>c.destroy());
        charts=[];
    }

    function makeChart(id,type,labels,datasets,extra={}) {
        charts.push(new Chart($('#'+id),{
            type,
            data:{labels,datasets},
            options:{
                responsive:true,
                maintainAspectRatio:false,
                animation:false,
                plugins:{
                    legend:{display:type==='doughnut',position:'bottom'},
                    tooltip:{callbacks:{
                        label:ctx=>`${ctx.dataset.label||ctx.label}: ${fmt(
                            typeof ctx.parsed==='number'?ctx.parsed:
                            extra.indexAxis==='y'?ctx.parsed.x:ctx.parsed.y
                        )}`
                    }}
                },
                ...(type==='doughnut'?{}:{
                    scales:{
                        y:{beginAtZero:true,grid:{color:'#edf1f6'}},
                        x:{grid:{display:false},ticks:{maxTicksLimit:18}}
                    }
                }),
                ...extra
            }
        }));
    }

    function scheduleCharts() {
        clearTimeout(chartTimer);
        if(!$('#analytics').hidden) chartTimer=setTimeout(renderAnalytics,100);
    }

    function renderAnalytics() {
        if(!S||$('#analytics').hidden) return;
        destroyCharts();

        const communities=grouped(allRows(),o=>String(o.community?.communityID||''));
        $('#aCommunity').innerHTML='<option value="">Все сообщества</option>'+
            communities.filter(([id])=>id).map(([id,rows])=>
                `<option value="${esc(id)}">${esc(rows[0].community?.title||id)}</option>`
            ).join('');
        $('#aCommunity').value=S.analytics.community;

        const rows=analyticsRows(), key=S.analytics.metric;
        const [from,to]=periodBounds();
        const all=allRows();
        const complete=rows.filter(completeStats).length;
        const noDate=all.filter(o=>!date(o.publishDateTime)).length;
        const ages=rows.flatMap(o=>Object.values(o.metricDates||{})).filter(Boolean).sort();

        $('#quality').textContent=
            `${from||'Начало истории'} — ${to||'сегодня'} · Публикаций: ${rows.length}`+
            ` · Основные показатели заполнены у ${complete}/${rows.length}`+
            ` · Список ${S.scan.complete?'полный':'НЕПОЛНЫЙ'}`+
            ` · Без даты публикации во всей памяти: ${noDate} (в графики не входят).`+
            (ages.length?` Показатели получены ${timeText(ages[0])} — ${timeText(ages.at(-1))}.`:'')+
            (from&&to&&from>to?' Ошибка: начало периода позже конца.':'');

        const kpiKeys=[
            'clickCount','totalOrderCount','buyoutCount','inProgressOrderCount',
            'predictedReward','factReward','buyoutRate','conversion','epc'
        ];

        $('#kpis').innerHTML=kpiKeys.map(k=>{
            const a=aggregate(rows,k);
            return `<div class="card"><div>${esc(LABEL[k])}</div>
                <div class="value">${metricText(k,a.value)}</div>
                <div class="muted">Данные: ${a.known}/${rows.length} заявок
                ${a.missing?' · частичный показатель':''}</div></div>`;
        }).join('');

        const periods=grouped(rows,bucket).sort((a,b)=>a[0].localeCompare(b[0]));
        const values=periods.map(([,r])=>aggregate(r,key).value);

        $('#trendTitle').textContent=
            `${LABEL[key]} — по ${S.analytics.group==='day'?'дням':
                S.analytics.group==='week'?'неделям':'месяцам'} публикации`;

        const top=grouped(rows,o=>
            String(o.community?.communityID||'')+'|'+(o.community?.title||'Без сообщества')
        ).map(([name,r])=>({name:name.split('|').slice(1).join('|'),a:aggregate(r,key)}))
            .filter(x=>x.a.value!==null)
            .sort((a,b)=>b.a.value-a.a.value).slice(0,10);

        $('#topTitle').textContent=`Сообщества · топ-10 · ${LABEL[key]}`;

        const balanced=rows.filter(o=>val(o,'other')!==null);
        const orderParts=['buyoutCount','inProgressOrderCount','other']
            .map(k=>sum(balanced,k).value||0);

        $('#ordersNote').textContent=
            `Только заявки с известными и непротиворечивыми тремя счётчиками: `+
            `${balanced.length}/${rows.length}. Остальные — не подтверждённые отказы.`+
            (orderParts.every(v=>v===0)?' Нет ненулевых данных для диаграммы.':'');

        if(typeof Chart==='undefined') {
            $('#quality').textContent+=' Библиотека графиков не загрузилась.';
        } else {
            makeChart('trend','bar',periods.map(([p])=>p),[{
                label:LABEL[key],data:values,
                backgroundColor:'#3487ef',borderRadius:5
            }]);
            makeChart('top','bar',top.map(x=>x.name),[{
                label:LABEL[key],data:top.map(x=>x.a.value),
                backgroundColor:'#7b6aef',borderRadius:5
            }],{indexAxis:'y'});
            makeChart('ordersChart','doughnut',
                ['Выкуплено','В обработке','Остальные — расчёт'],[{
                    data:orderParts,backgroundColor:['#30b48a','#f3b447','#b6c1d2'],
                    borderWidth:2,borderColor:'#fff'
                }],{cutout:'68%'});
        }

        $('#analysisTable').innerHTML=
            '<table><thead><tr><th>Период публикации</th><th>Заявок</th>'+
            METRICS.map(([,t])=>`<th>${esc(t)}</th>`).join('')+
            '</tr></thead><tbody>'+
            periods.map(([p,r])=>`<tr><td>${p}</td><td>${r.length}</td>`+
                METRICS.map(([k])=>{
                    const a=aggregate(r,k);
                    return `<td title="Данные: ${a.known}/${r.length}">`+
                        `${metricText(k,a.value)}${a.missing?' *':''}</td>`;
                }).join('')+'</tr>').join('')+
            '</tbody></table><div class="muted">* Неполное покрытие данных. '+
            'Пустые интервалы без публикаций не добавляются. '+
            'Доли рассчитываются по суммам сопоставимых строк, не как среднее процентов.</div>';
    }

    // ---------- Excel ----------

    function sheet(book,name,headers,rows) {
        const data=[headers,...rows].map(r=>r.map(v=>
            typeof v==='string'?v.slice(0,32767):v
        ));
        const ws=XLSX.utils.aoa_to_sheet(data,{cellDates:true});
        ws['!cols']=headers.map(h=>({
            wch:/Товар|Текст|Ссылка|Ошибка|Пояснение/.test(h)?36:22
        }));
        if(rows.length) ws['!autofilter']={ref:ws['!ref']};

        for(const address of Object.keys(ws)) {
            if(address.startsWith('!')) continue;
            const cell=ws[address];
            if(cell.t==='n') cell.z='#,##0.##';
            if(cell.t==='d') cell.z='dd.mm.yyyy hh:mm';
            if(cell.t==='s'&&/^https?:\/\//i.test(cell.v)&&url(cell.v)) {
                cell.l={Target:cell.v};
            }
        }
        XLSX.utils.book_append_sheet(book,ws,name);
    }

    const exportHeaders=[
        'Номер заявки','UUID','Товар','Сообщество','Категория','Статус','Код статуса',
        'Формат','Цена товара, ₽','Процент вознаграждения',
        ...METRICS.map(([,t])=>t),
        'Подтверждённые отказы — неизвестны','goodsSold API',
        'Дата публикации','Срок до','Кампания с','Кампания до',
        'Ссылка на заявку','Ссылка на публикацию','Ссылка на товар',
        'Ссылка на сообщество','Текст публикации','Комментарий статуса',
        'Список обновлён','Карточка обновлена','Ошибка',
        'price100 API','rate1000 API','sum100 API','cpm100 API',
        ...CORE.map(k=>'Получено: '+k)
    ];

    function exportRow(o) {
        return [
            String(o.number||''),o.uuid,o.product?.name||'',o.community?.title||'',
            categoryName(category(o)),status(o),o.status?.code||'',o.format||'',
            val(o,'price'),val(o,'rate'),
            ...METRICS.map(([k])=>val(o,k)),
            null,val(o,'goodsSold'),
            date(o.publishDateTime),date(o.status?.expiresAt),
            o.campaignPeriod?.from||'',o.campaignPeriod?.to||'',
            orderUrl(o),url(o.publishedUrl),url(o.product?.externalURL),
            url(o.community?.link),o.message||'',o.statusComment||'',
            date(o.listUpdated),date(o.detailUpdated),o.error||'',
            n(o.product?.price100),n(o.product?.rate1000),n(o.sum100),n(o.stats?.cpm100),
            ...CORE.map(k=>date(o.metricDates?.[k]))
        ];
    }

    function summaryRow(name,rows) {
        return [name,rows.length,...METRICS.map(([k])=>aggregate(rows,k).value),
            ...METRICS.map(([k])=>aggregate(rows,k).missing)];
    }

    function exportExcel() {
        if(busy||!allRows().length) return;
        if(!S.scan.complete&&!confirm('Список неполный. Выгрузить сохранённые данные?')) return;

        try {
            const wb=XLSX.utils.book_new(), rows=allRows();
            const published=rows.filter(o=>category(o)==='published');
            const summaryHeaders=[
                'Группа','Заявок',...METRICS.map(([,t])=>t),
                ...METRICS.map(([,t])=>'Без данных: '+t)
            ];

            const summaries=[
                summaryRow('Все загруженные — без отменённых',
                    rows.filter(o=>category(o)!=='canceled')),
                summaryRow('Опубликованные',published)
            ];

            const months=grouped(published,o=>localDay(date(o.publishDateTime)).slice(0,7))
                .sort((a,b)=>a[0].localeCompare(b[0]));
            const years=grouped(published,o=>localDay(date(o.publishDateTime)).slice(0,4))
                .sort((a,b)=>a[0].localeCompare(b[0]));

            years.forEach(([p,r])=>summaries.push(summaryRow(p+' год',r)));
            months.forEach(([p,r])=>summaries.push(summaryRow(p,r)));
            ['unpublished','canceled','unknown'].forEach(k=>
                summaries.push(summaryRow(categoryName(k),rows.filter(o=>category(o)===k)))
            );

            sheet(wb,'Сводка',summaryHeaders,summaries);
            sheet(wb,'Все заявки',exportHeaders,rows.map(exportRow));
            sheet(wb,'Текущая выборка',exportHeaders,visibleRows().map(exportRow));

            for(const k of ['published','unpublished','canceled','unknown']) {
                const r=rows.filter(o=>category(o)===k);
                if(r.length) sheet(wb,categoryName(k),exportHeaders,r.map(exportRow));
            }
            for(const [p,r] of months) sheet(wb,p,exportHeaders,r.map(exportRow));

            const ar=analyticsRows();
            const groups=grouped(ar,bucket).sort((a,b)=>a[0].localeCompare(b[0]));
            sheet(wb,'Аналитика периода',summaryHeaders,[
                summaryRow('Итого по выбранным публикациям',ar),
                ...groups.map(([p,r])=>summaryRow(p,r))
            ]);

            sheet(wb,'Описание',['Параметр','Пояснение'],[
                ['Профиль',profile],
                ['Создано',new Date()],
                ['В памяти заявок',rows.length],
                ['totalCount API',S.scan.total],
                ['Полный список',S.scan.complete?'Да':'Нет'],
                ['Следующий offset',S.scan.offset],
                ['Последняя страница получена',date(S.scan.updated)],
                ['Период аналитики',periodBounds().join(' — ')||'Всё время'],
                ['Группировка аналитики',S.analytics.group],
                ['Сообщество аналитики',S.analytics.community||'Все'],
                ['Смысл графиков','Накопленные показатели заявок по дате публикации, не события за период'],
                ['Доля выкупа','Выкуплено / всего заказов × 100%. В сводке — сопоставимые строки'],
                ['Конверсия','Всего заказов / клики × 100%. Не доля уникальных покупателей'],
                ['Факт на клик','Фактическое вознаграждение / клики'],
                ['Остальные заказы','Всего − в обработке − выкуплено; только при наличии всех полей и неотрицательном результате'],
                ['Отказы','Подтверждённое поле отсутствует. Остальные заказы нельзя автоматически считать отказами'],
                ['Прогноз','predictedReward / '+CFG.rewardDivisor+'; не гарантированная выплата'],
                ['Факт','factReward / '+CFG.rewardDivisor+'; не баланс для вывода'],
                ['Цена','price100 / '+CFG.priceDivisor],
                ['Процент','rate1000 / '+CFG.rateDivisor],
                ['Пустые значения','Нет достоверного числа; не заменяются нулями'],
                ['Свежесть','Сохранена отдельно для каждого основного поля; обновление списка не всегда обновляет детали'],
                ['Часовой пояс',Intl.DateTimeFormat().resolvedOptions().timeZone],
                ['Повторение строк','Категории, месяцы и текущая выборка дублируют строки листа Все заявки. Не суммируйте листы между собой'],
                ['Фильтры таблицы',JSON.stringify(S.ui)]
            ]);

            XLSX.writeFile(wb,
                `VK_AdBlogger_${localDay(new Date())}.xlsx`,
                {compression:true});
        } catch(e) {
            alert('Ошибка Excel: '+e.message);
        }
    }

    // ---------- События ----------

    function changeSort(key) {
        if(S.ui.sort===key) S.ui.dir*=-1;
        else {S.ui.sort=key;S.ui.dir=-1;}
        S.ui.page=1;
        $('#sort').value=key;
        render();saveUI();
    }

    $('#launch').onclick=()=>{
        $('#panel').hidden=false;
        $('#launch').hidden=true;
        render();
    };
    $('#close').onclick=()=>{
        $('#panel').hidden=true;
        $('#launch').hidden=false;
    };
    $('#showAnalytics').onclick=()=>{
        $('#analytics').hidden=false;
        syncControls();
        scheduleCharts();
    };
    $('#aClose').onclick=()=>{
        $('#analytics').hidden=true;
        destroyCharts();
    };

    $('#load').onclick=()=>run(()=>loadPages(false));
    $('#rescan').onclick=()=>{
        if(!confirm(
            'Начать новый проход списка?\n\n'+
            'Сохранённые детали не стираются. Очередь проверки карточек сбросится. '+
            'После полного прохода заявки, которых больше нет в списке, будут удалены из памяти.'
        )) return;
        run(()=>loadPages(true));
    };
    $('#details').onclick=()=>run(()=>details('missing'));
    $('#resume').onclick=()=>run(()=>details('resume'));
    $('#retry').onclick=()=>run(()=>details('errors'));
    $('#refresh').onclick=()=>{
        if(confirm(`Повторно запросить все ${allRows().length} карточек? Это может занять несколько минут.`)) {
            run(()=>details('all'));
        }
    };
    $('#stop').onclick=()=>controller?.abort();

    $('#search').oninput=e=>{
        S.ui.query=e.target.value;S.ui.page=1;
        render();saveUI();
    };
    $('#category').onchange=e=>{
        S.ui.category=e.target.value;S.ui.page=1;
        render();saveUI();
    };
    $('#sort').onchange=e=>{
        S.ui.sort=e.target.value;S.ui.dir=-1;S.ui.page=1;
        render();saveUI();
    };
    $('#direction').onclick=()=>{
        S.ui.dir*=-1;render();saveUI();
    };
    $('#zeros').onchange=e=>{
        S.ui.hideZero=e.target.checked;S.ui.page=1;
        render();saveUI();
    };
    $('#reset').onclick=()=>{
        S.ui={query:'',category:'',sort:'',dir:-1,hideZero:false,page:1};
        syncControls();render();saveUI();
    };
    $('#prev').onclick=()=>{S.ui.page--;render();saveUI();};
    $('#next').onclick=()=>{S.ui.page++;render();saveUI();};

    root.querySelectorAll('[data-sort]').forEach(button=>{
        button.onclick=()=>changeSort(button.dataset.sort);
    });

    for(const [id,key] of [
        ['period','period'],['from','from'],['to','to'],
        ['group','group'],['aMetric','metric'],['aCommunity','community']
    ]) $('#'+id).onchange=e=>{
        S.analytics[key]=e.target.value;
        syncControls();saveUI();scheduleCharts();
    };

    $('#excel').onclick=$('#aExcel').onclick=exportExcel;

    $('#profile').onchange=async e=>{
        if(busy) return;
        try {await switchProfile(e.target.value);}
        catch(error){alert(error.message);}
    };

    $('#newProfile').onclick=async()=>{
        if(busy) return;
        const name=prompt('Название профиля для текущего кабинета (до 60 символов):');
        if(!name?.trim()) return;
        try {
            await switchProfile(name.trim().slice(0,60));
            await save();
            await refreshProfileMenu();
        } catch(e) {alert(e.message);}
    };

    $('#clear').onclick=async()=>{
        if(busy||!confirm(`Удалить все сохранённые данные профиля «${profile}»?`)) return;
        S=blankState();
        message='Память профиля очищена.';
        try {await save();}
        catch(e){message=e.message;}
        syncControls();render();scheduleCharts();
    };
    // ---------- Количество отображаемых строк ----------

    const pageSizeLabel = document.createElement('label');
    pageSizeLabel.innerHTML = `
        Показывать:
        <select id="pageSize" title="Количество строк в таблице">
            <option value="50">50 строк</option>
            <option value="100">100 строк</option>
            <option value="200">200 строк</option>
            <option value="all">Все строки</option>
        </select>
    `;

    $('#prev').before(pageSizeLabel);

    $('#pageSize').onchange = event => {
        S.ui.pageSize = event.target.value;
        S.ui.page = 1;

        render();
        saveUI();
    };

    // Сброс фильтров не должен сбрасывать выбранный размер таблицы.
    $('#reset').onclick = () => {
        const pageSize = S.ui.pageSize ?? 'all';

        S.ui = {
            query: '',
            category: '',
            sort: '',
            dir: -1,
            hideZero: false,
            page: 1,
            pageSize
        };

        syncControls();
        render();
        saveUI();
    };
    // ---------- Запуск и защита от двух пишущих вкладок ----------

    async function boot() {
        db=await openDB();
        let name=localStorage.getItem(CFG.lastProfile);
        if(!name) {
            name=prompt(
                'Назовите профиль памяти для этого кабинета.\n'+
                'Для другого кабинета нужно будет выбрать отдельный профиль.',
                'Основной кабинет'
            )?.trim().slice(0,60)||'Основной кабинет';
        }
        await switchProfile(name);
    }
    // ---------- Исправление читаемости таблицы ----------

    const readableTableStyle = document.createElement('style');

    readableTableStyle.textContent = `
        /* Только основная таблица. Графики и их таблицу не меняем. */
        #panel > .scroll > table {
            width: 100%;
            min-width: 1280px;
        }

        #panel > .scroll > table > thead > tr > th,
        #panel #body > tr > td {
            padding: 12px 14px;
            border-right: 1px solid #e4eaf2;
        }

        #panel > .scroll > table > thead > tr > th:last-child,
        #panel #body > tr > td:last-child {
            border-right: none;
        }

        #panel > .scroll > table > thead > tr > th {
            background: #edf2f9;
            color: #344b69;
            font-weight: 600;
            border-bottom: 2px solid #d9e2ee;
        }

        #panel #body > tr:nth-child(even) {
            background: #fafbfd;
        }

        #panel #body > tr:hover {
            background: #edf5ff;
        }

        /*
         * Подпись и значение больше не растягиваются
         * на всю ширину ячейки.
         */
        #panel #body .metric {
            display: grid;
            grid-template-columns: 125px minmax(65px, max-content);
            justify-content: start;
            align-items: baseline;
            column-gap: 10px;
            min-height: 25px;
            padding: 2px 0;
            white-space: nowrap;
        }

        #panel #body .metric > span {
            color: #5c6e84;
            font-size: 12px;
        }

        #panel #body .metric > b {
            display: block;
            text-align: right;
            color: #203a58;
            font-size: 13px;
            font-weight: 700;
            font-variant-numeric: tabular-nums;
        }

        /* Всего заказов — главный показатель блока. */
        #panel #body > tr > td:nth-child(5) .metric:first-child {
            border-bottom: 1px solid #dce5ef;
            padding-bottom: 5px;
            margin-bottom: 3px;
        }

        #panel #body > tr > td:nth-child(5) .metric:first-child > span {
            color: #203a58;
            font-weight: 700;
        }

        /* Различаем состояния заказов, сохраняя подписи. */
        #panel #body > tr > td:nth-child(5) .metric:nth-child(2) > b {
            color: #956300;
        }

        #panel #body > tr > td:nth-child(5) .metric:nth-child(3) > b {
            color: #197b54;
        }

        #panel #body > tr > td:nth-child(5) .metric:nth-child(4) > b {
            color: #6d7d91;
        }

        /* Ширина числовых столбцов и согласование с заголовками. */
        #panel > .scroll > table > thead > tr > th:nth-child(4),
        #panel #body > tr > td:nth-child(4) {
            text-align: right;
            white-space: nowrap;
            min-width: 85px;
        }

        #panel > .scroll > table > thead > tr > th:nth-child(7),
        #panel > .scroll > table > thead > tr > th:nth-child(8),
        #panel #body > tr > td:nth-child(7),
        #panel #body > tr > td:nth-child(8) {
            text-align: right;
            white-space: nowrap;
            min-width: 115px;
        }

        #panel #body > tr > td:nth-child(7),
        #panel #body > tr > td:nth-child(8) {
            font-size: 14px;
            font-variant-numeric: tabular-nums;
        }

        #panel #body > tr > td:nth-child(5),
        #panel #body > tr > td:nth-child(6) {
            min-width: 230px;
        }
    `;

    root.appendChild(readableTableStyle);
    // Одна активная вкладка инструмента исключает перезапись очереди
    // другой вкладкой. Родной сайт продолжает работать во всех вкладках.
    if(!navigator.locks) {
        $('#launch').textContent='Нужен актуальный Chrome / Edge';
        $('#launch').onclick=()=>alert(
            'В этом браузере недоступен Web Locks API. '+
            'Откройте страницу в актуальном Chrome или Edge.'
        );
        return;
    }

    navigator.locks.request('vk-adblogger-v15-active-tab',
        {ifAvailable:true},async lock=>{
            if(!lock) {
                $('#launch').textContent='Аналитика открыта в другой вкладке';
                $('#launch').onclick=()=>alert(
                    'Закройте другую вкладку с этим скриптом и перезагрузите эту страницу. '+
                    'Это защита локальной памяти от одновременной записи.'
                );
                return;
            }
            try {
                await boot();
                // Блокировка удерживается до закрытия или перезагрузки вкладки.
                await new Promise(()=>{});
            } catch(e) {
                $('#launch').textContent='Ошибка запуска аналитики';
                $('#launch').onclick=()=>alert(
                    'Не удалось открыть локальную память: '+e.message
                );
                console.error('VK AdBlogger v15:',e);
            }
        }).catch(e=>console.error('VK AdBlogger lock:',e));
})();