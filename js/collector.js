// Сборщик продавцов Авто.ру. Запускается ЦЕЛИКОМ в javascript_tool во вкладке auto.ru.
// Работает в фоне (async), результат лежит в window.__M. Один вызов javascript_tool
// живёт ~45 секунд, поэтому скрипт НЕ ждёт себя: он стартует и возвращает 'started'.
//
// Перед запуском замените блок CONFIG. Остальное не трогайте.
window.__M = { res: [], log: [], done: false, state: {}, dupSkipped: 0, requests: 0, stop: null };
(async () => {
  // ===== CONFIG =====
  const CONFIG = {
    region: 'moskovskaya_oblast',          // slug региона в URL Авто.ру (moskva, sankt-peterburg, krasnodarskiy_kray ...)
    models: [['toyota','camry'],['kia','rio'],['hyundai','solaris']], // [марка, модель] как в URL Авто.ру
    perModelTarget: 8,                      // сколько подходящих продавцов набрать на модель
    maxPages: 6,                            // максимум страниц выдачи на модель
    minGarage: 2,                           // минимум машин в гараже без пометок «Бывшая» и «Машина мечты»
    maxRequests: 400,                       // потолок запросов за один запуск
    delayMs: 600,                           // пауза между запросами
    knownProfileIds: []                     // ID профилей из Excel (чтобы не брать повторно)
  };
  // ==================
  const M = window.__M;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const parse = t => new DOMParser().parseFromString(t, 'text/html');
  const clean = h => h ? h.split('?')[0] : null;           // query-строки не выводим: расширение блокирует такой вывод
  const known = new Set(CONFIG.knownProfileIds);
  const seenP = new Set(), seenName = new Set();
  const blocked = html => /showcaptcha|captcha|Подтвердите, что вы не робот|Доступ ограничен/i.test(html);
  const get = async url => {
    if (M.requests >= CONFIG.maxRequests) { M.stop = 'maxRequests'; throw new Error('STOP:maxRequests'); }
    M.requests++;
    const r = await fetch(url, { credentials: 'include' });
    const html = await r.text();
    if (r.status === 429 || r.status === 403 || (blocked(html) && !/ListingItemUniversal/.test(html))) {
      M.stop = 'blocked:' + r.status; throw new Error('STOP:blocked');
    }
    return { status: r.status, doc: parse(html) };
  };
  try {
    for (const [br, mo] of CONFIG.models) {
      let q = 0; M.state.model = br + ' ' + mo;
      for (let p = 1; p <= CONFIG.maxPages && q < CONFIG.perModelTarget; p++) {
        const { status, doc } = await get(`/${CONFIG.region}/cars/${br}/${mo}/used/?seller_group=PRIVATE&page=${p}`);
        const cards = [...doc.querySelectorAll('[class*="ListingItemUniversal__snippet"]')];
        M.log.push(`${br} ${mo} p${p} status ${status} cards ${cards.length}`);
        if (!cards.length) break;
        for (const el of cards) {
          if (q >= CONFIG.perModelTarget) break;
          const sb = el.querySelector('[class*="ListingItemUniversalSeller"]');
          const sellerText = sb ? sb.textContent.trim() : '';
          const a = el.querySelector('a[href*="/cars/used/sale/"]');
          const listing = clean(a && a.getAttribute('href'));
          if (!listing) continue;
          const key = sellerText.slice(0, 40);              // грубая предпроверка: один и тот же продавец в выдаче
          if (seenName.has(key)) continue; seenName.add(key);
          try {
            const { doc: d } = await get(listing);          // страница объявления -> ссылка на профиль
            const pl = [...d.querySelectorAll('a')].map(x => clean(x.getAttribute('href'))).find(h => h && /\/profile\//.test(h));
            if (!pl) { await sleep(CONFIG.delayMs); continue; }
            const pid = pl.split('/').filter(Boolean).pop();
            if (seenP.has(pid)) { await sleep(300); continue; } seenP.add(pid);
            if (known.has(pid)) { M.dupSkipped++; await sleep(300); continue; }
            const { doc: pd } = await get(pl.replace('https://auto.ru', ''));   // профиль -> гараж
            const items = [...pd.querySelectorAll('[class*="PublicProfileGarageItem-"]')]
              // «Бывшая» (продана) и «Машина мечты» (желаемая) — это не машины в собственности, в счёт не идут
              .map(n => { const raw = n.textContent.trim(); const not = /Бывшая|Машина мечты/.test(raw);
                          return { t: raw.replace(/Бывшая|Машина мечты/, ''), former: not, tag: /Машина мечты/.test(raw) ? 'мечта' : 'бывшая' }; });
            const cur = items.filter(x => !x.former).length;
            if (cur >= CONFIG.minGarage) {
              q++;
              M.res.push({ q: br + ' ' + mo, pid, name: key, gt: items.length, gc: cur,
                           g: items.map(x => x.t + (x.former ? ' (' + x.tag + ')' : '')).join('; '), l: listing });
            }
          } catch (e) { if (String(e.message).startsWith('STOP')) throw e; }
          await sleep(CONFIG.delayMs);
        }
        await sleep(CONFIG.delayMs);
      }
    }
  } catch (e) { M.log.push('stopped: ' + e.message); }
  M.done = true;
})();
'started'
