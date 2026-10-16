// Опрос прогресса (короткий вывод!). Вызывать после ожидания <=10 сек (computer wait).
JSON.stringify({ done: window.__M.done, stop: window.__M.stop, model: window.__M.state.model,
  found: window.__M.res.length, requests: window.__M.requests, dupSkipped: window.__M.dupSkipped,
  log: window.__M.log.slice(-2) })
