---
name: "avtoru-reseller-finder"
description: "Находит на Авто.ру частных продавцов с 2+ машинами (вероятные перекупы) по заданным моделям, перед запуском уточняет критерии, создаёт на рабочем столе папку и кладёт в неё Excel-файл запуска с номером и датой (NNN_resellers_ГГГГ-ММ-ДД.xlsx), без дублей по всем прошлым файлам. Используй, когда просят найти перекупов, собрать продавцов с несколькими авто, пополнить базу Авто.ру или запустить ежедневный сбор по расписанию, даже без слов «скилл» и «Авто.ру». Телефоны не собирает, сообщений не отправляет."
---

# Поиск перекупов на Авто.ру

Скилл превращает «найди мне перекупов на Авто.ру» в повторяемый процесс: выдача по модели, профиль каждого частного продавца, оценка числа его машин по гаражу, дедупликация по ID профиля, запись в Excel. Он рассчитан на автономный запуск другим экземпляром Claude, поэтому здесь много конкретики: что делать, чем проверять и когда остановиться.

## Кто пользователь и зачем это нужно

Владелец базы хочет связываться с перекупами (продавцами с несколькими машинами) по ссылкам на их объявления, через штатный чат площадки, небольшими порциями в день. Из этого следует, что нужны именно устойчивые идентификаторы (ID профиля, ссылка на объявление) и статусы «писали / не писали», а не контактные данные.

## Границы (прочитайте до начала работы)

1. **Телефоны и другие личные контакты не собираются.** Кнопку «Показать телефон» не нажимать, номера в таблицу не вносить, даже если пользователь просит. Это массовое извлечение персональных данных частных лиц. Связь идёт через чат площадки по ссылке на объявление. Если просят номера, коротко откажите и предложите ссылки.
2. **Капчу не обходить.** При капче, 403/429 или странице «Подтвердите, что вы не робот» остановиться, записать причину в журнал и сообщить пользователю.
3. **Сообщений продавцам скилл не отправляет.** Отправка от имени пользователя требует его явного подтверждения по каждому действию. Скилл готовит список, пишет человек (или отдельная задача с подтверждением).
4. **Разрешения расширения не обходить.** Если появляется запрос разрешения на auto.ru, значит, оно не выдано: остановиться и сообщить (см. раздел «Автозапуск»).
5. Не заходить на страницы входа, не вводить пароли, не создавать аккаунты.

## Что нужно на входе

| Параметр | Значение по умолчанию | Примечание |
|---|---|---|
| Модели | Toyota Camry, Kia Rio, Hyundai Solaris | 2-3 модели за запуск. Формат `[марка, модель]` как в URL Авто.ру: `['toyota','camry']`, `['volkswagen','polo']` |
| Регион | `moskovskaya_oblast` | slug из URL (`moskva`, `sankt-peterburg` ...) |
| На модель | 8 продавцов | хватает для порции в день |
| Минимум машин в гараже | 2 | без пометок «Бывшая» и «Машина мечты» |
| Кого искать | только частники | автосалоны и компании не берём (см. «Критерий») |
| Папка | `AutoRu_Resellers` на рабочем столе | создаётся шагом 1; в ней каждый запуск даёт новый файл `NNN_resellers_ГГГГ-ММ-ДД.xlsx` |

Если пользователь не назвал модели, берите ходовые у перекупов и меняйте по кругу между запусками: Camry, Rio, Solaris, Polo, Octavia, Rapid, Granta, Creta, Tiguan, Focus, Cerato, Lancer, Qashqai, X-Trail, Corolla, Mazda 3, Astra. Следующие модели в ротации берутся после тех, что записаны в `criteria.json` как последние (шаг 0 и шаг 1).

## Предусловия (проверьте и не продолжайте, если не выполнены)

- Доступны инструменты Claude in Chrome (`mcp__claude-in-chrome__*`). Если они отложенные, загрузите нужные одним вызовом ToolSearch: `select:mcp__claude-in-chrome__tabs_context_mcp,mcp__claude-in-chrome__navigate,mcp__claude-in-chrome__javascript_tool,mcp__claude-in-chrome__computer`.
- Для auto.ru выдано постоянное разрешение. Проверка выполняется в шаге 2: первый `javascript_tool` не должен вызывать окно подтверждения. Если вызвал и подтверждения нет, остановитесь и скажите пользователю, что нужно один раз выдать постоянное разрешение для auto.ru в настройках расширения. Это лучше, чем засыпать его запросами.
- Есть Python с `openpyxl` (обычно предустановлен; в Windows команда может называться `python` или `py`) и доступ к рабочему столу (шаг 1).

## Процесс

### Шаг 0. Уточните критерии (всегда, до любых действий на сайте)

Пользователь просил, чтобы перед запуском критерии подтверждались. Поэтому в обычной (живой) сессии первым делом задайте короткие вопросы, даже если есть значения по умолчанию. Если доступен `AskUserQuestion`, задайте одним вызовом 3 вопроса (в каждом первым вариантом поставьте текущее значение с пометкой «(Рекомендуется)»; «Другое» пользователь впишет сам):
1. **Модели**: «Camry, Rio, Solaris» / «следующие по ротации: …» / свои.
2. **Регион**: «Московская область» / «Москва» / «Санкт-Петербург» / свой.
3. **Размер порции и строгость**: «8 на модель, от 2 машин в гараже» / «8 на модель, от 3 машин (строже)» / «15 на модель, от 2 машин».

Если `AskUserQuestion` нет, напишите те же вопросы обычным текстом с готовыми значениями «ответьте "ок", если подходит» и дождитесь ответа. Если критерии уже названы в просьбе («найди по Polo и Octavia, Москва»), не переспрашивайте всё заново: повторите их одной строкой, подставьте остальное по умолчанию и попросите подтвердить («ок» или правки). Не запускайте сбор, пока нет ответа.

Подтверждённые критерии сохраняются на шаге 1 в `criteria.json` (папка запусков). Их значения идут в `CONFIG` сборщика и в `--run-note` Excel.

**Автозапуск и запуск по расписанию (человека нет рядом):** вопросов не задавайте, они зависнут. Возьмите критерии из `criteria.json` (с ротацией моделей на следующие три) или, если файла нет, значения по умолчанию из таблицы выше. Первой строкой итога напишите, с какими критериями был запуск.

### Шаг 1. Папка на рабочем столе и «известные» ID

Все файлы живут в одной папке `AutoRu_Resellers` на рабочем столе пользователя. Не угадывайте путь: рабочий стол бывает на другом месте (OneDrive, другая локализация), а сессия может не иметь доступа к его файлам. Определите, где вы работаете:

**А. Облачная сессия, связанная с компьютером пользователя** (в списке инструментов есть `mcp__remote-devices__*`; на компьютере пользователя обычно Windows):
1. `get_device_info`: посмотрите `connectedFolders`. Если рабочего стола (`Desktop`) там нет, один раз вызовите `device_request_folder_access` с путём `~/Desktop` и коротким `reason` («Создать папку AutoRu_Resellers и сохранять в неё Excel с продавцами»). Просите только рабочий стол, не всю домашнюю папку. Дождитесь одобрения.
2. Папка в `device_bash`: `mkdir -p "$HOME/mnt/Desktop/AutoRu_Resellers"` (имя после `mnt/` смотрите по `ls "$HOME/mnt"`, оно совпадает с подключённой папкой), затем проверьте запись и номера: `python3 setup_folder.py --dir "$HOME/mnt/Desktop/AutoRu_Resellers"`.
3. Если в `device_bash` нет Python с `openpyxl` (и `pip install openpyxl` не работает), делайте Excel в облачной рабочей папке: перенесите имеющиеся файлы папки через `device_stage_files`, запустите там `known_ids.py` и `build_xlsx.py`, готовый файл верните `device_commit_files` по пути вида `<путь рабочего стола из connectedFolders>\AutoRu_Resellers\NNN_resellers_ГГГГ-ММ-ДД.xlsx`. Не пересобирайте файл из памяти.
4. Если пользователь отказал в доступе к рабочему столу: работайте в рабочей папке сессии, файл отправьте через `SendUserFile` и скажите, что на рабочий стол он не попал.

**Б. Агент запущен прямо на компьютере пользователя** (локальная сессия, расписание в приложении): `python setup_folder.py` (на Windows это `python` или `py`, на macOS/Linux `python3`). Скрипт сам находит настоящий рабочий стол (в том числе перенесённый в OneDrive), создаёт `AutoRu_Resellers` и печатает JSON.

`setup_folder.py` (Приложение F) печатает `folder`, `files` (уже существующие файлы запусков), `next_number` и `criteria` (последние сохранённые). После шага 0 сохраните подтверждённые критерии: `python3 setup_folder.py --dir <папка> --save-criteria '{"region":"moskovskaya_oblast","models":[["toyota","camry"],["kia","rio"],["hyundai","solaris"]],"per_model":8,"min_garage":2}'` (на Windows в PowerShell экранируйте кавычки или запишите JSON файлом `criteria.json` в папку напрямую).

Если ответ `error: desktop_not_found`, спросите у пользователя путь к рабочему столу и передайте его через `--dir`.

Затем получите ID, которые уже есть в папке (по всем прошлым файлам):
```bash
python3 known_ids.py "<папка>"
```
Получится JSON-массив (для пустой папки `[]`). Он вставляется в `CONFIG.knownProfileIds` сборщика, чтобы повторно никого не брать. Это первый уровень защиты от дублей; второй (по ID внутри самого запуска) есть в сборщике, третий (отсев по всем файлам папки) делает `build_xlsx.py`, четвёртый (формула «Дубль?») стоит в таблице. Старый файл `resellers.xlsx` из пробного запуска, если лежит в папке, тоже учитывается.

### Шаг 2. Вкладка
1. `tabs_context_mcp` с `createIfEmpty: true`. Каждый запуск работает в своей новой вкладке; tabId из прошлых сессий не использовать.
2. `navigate` на `https://auto.ru/{регион}/cars/{марка}/{модель}/used/?seller_group=PRIVATE` для первой модели. Это нужно, чтобы запросы шли с домена auto.ru (иначе `fetch` упрётся в CORS) и с cookies пользователя.
3. Один пробный `javascript_tool`: `document.querySelectorAll('[class*="ListingItemUniversal__snippet"]').length`. Ожидается около 37. Если 0, прочитайте раздел «Заметки о сайте»: возможно, изменилась вёрстка.

### Шаг 3. Запуск сборщика
1. Возьмите код сборщика из Приложения A, подставьте в блок `CONFIG` регион, модели, цели и `knownProfileIds`.
2. Вставьте скрипт целиком в `javascript_tool`. Он вернёт `started` сразу, а работать будет в фоне.
   **Почему так:** один вызов `javascript_tool` живёт примерно 45 секунд (потом «CDP timed out»), а сбор по трём моделям занимает 5-10 минут. Фоновая async-функция живёт, пока открыта вкладка.
3. Не закрывайте вкладку и не уходите с неё навигацией до конца (иначе фоновая задача умрёт вместе со страницей).

### Шаг 4. Ожидание и опрос
Повторяйте: `computer` с `action: "wait"`, `duration: 10` (больше 10 секунд инструмент не принимает), затем код из Приложения B через `javascript_tool`. Смотрите на `done`, `stop`, `found`, `requests`.
- `done: true` и `stop: null`: сбор завершён.
- `stop` не пуст (`blocked:...` или `maxRequests`): сбор остановлен сборщиком. Это нормальная остановка: заберите то, что собрано, запишите причину в сводку. При `blocked` не повторяйте запуск сегодня.
- Не опрашивайте чаще, чем раз в 10 секунд: это ничего не ускоряет.

### Шаг 5. Выгрузка результатов
Вывод `javascript_tool` обрезается (около 1-1,5 тыс. символов) и блокируется, если в нём есть query-строки со знаком `?`. Поэтому выгружайте порциями через код из Приложения C: меняйте `START`/`END` по 3-4 записи (`0..4`, `4..8`, ...), пока не получите все `found` записей. Складывайте строки в текстовый файл `rows.txt` (одна запись = одна строка, формат описан в начале Приложения C). Если результат вернулся с пометкой `[TRUNCATED]`, уменьшите шаг и повторите именно эту порцию. Независимые порции можно запрашивать параллельно.

Перечитайте итог: число строк в `rows.txt` должно равняться `found`.

### Шаг 6. Excel (новый файл с номером и датой)
```bash
python3 build_xlsx.py rows.txt "<папка>" --run-note "Camry/Rio/Solaris, moskovskaya_oblast, 8 на модель, от 2 машин"
```
Скрипт (Приложение D, сохраните как `build_xlsx.py`) каждый запуск создаёт **новый** файл `NNN_resellers_ГГГГ-ММ-ДД.xlsx`: `NNN` это следующий порядковый номер в папке (001, 002 ...), дата сегодняшняя. Два запуска в один день дают два файла с разными номерами, старые файлы не изменяются (в них остаются статусы контакта, которые пользователь проставил руками). Дубли отсекаются по ID профиля среди **всех** файлов папки; если новых продавцов нет, файл не создаётся (`"file": null`). В файле: лист «Продавцы» (гиперссылки, «Уверенность», выпадающий «Статус контакта» со значением `не писали`, формула «Дубль?») и лист «Журнал запусков» с критериями запуска. Скрипт печатает JSON `file / added / duplicates_skipped`.

Проверка: откройте созданный файл через `openpyxl`, убедитесь, что число строк равно `added`, ссылки имеют вид `https://auto.ru/profile/<ID>/` и `https://auto.ru/cars/used/sale/...`, а в «Дубль?» нет «ДУБЛЬ». Прогоните пересчёт (`recalc.py` из xlsx-скилла, если он есть), чтобы формулы имели значения.

### Шаг 7. Итог пользователю
Коротко: с какими критериями шёл запуск, сколько добавлено, сколько дублей отсеяно, по каким моделям, были ли остановки, полный путь к созданной папке и имя файла (например, «Рабочий стол → AutoRu_Resellers → 002_resellers_2026-10-16.xlsx»). Одна фраза про ограничение (ниже). Без пересказа шагов.

## Как работает критерий «перекуп» и его честные пределы

Продавец проходит, если в его публичном профиле гараж содержит **2 и более машин без пометок «Бывшая» (продана) и «Машина мечты» (желаемая, не в собственности)**. Профиль частника открыт без входа в аккаунт, поэтому это самый доступный признак. Но гараж не равен «объявлениям на продаже»: человек может держать в гараже семейную машину, которую не продаёт. Поэтому:

- В таблице стоит колонка «Уверенность»: 3+ машин в гараже означает «средняя», ровно 2 означает «низкая-средняя, проверить».
- Перед рассылкой человек должен по ссылке на объявление и профиль убедиться, что это похоже на перекупа (одинаковые описания, быстрая смена машин, разные модели в разных ценовых классах).
- Статус «не перекуп» в списке выбора нужен, чтобы вручную вычёркивать ошибочные записи, и такие ID остаются в базе, чтобы агент не взял их повторно.
- Сильный признак: в гараже несколько одинаковых моделей или быстро сменившиеся машины. Слабый: две машины разных классов (может быть семья).
- Не называйте найденных людей «перекупами» как факт. Говорите «вероятные перекупы» или «продавцы с 2+ машинами».

Профи-продавцы (автосалоны, компании) нужны не всегда. Сборщик ищет только частников (`seller_group=PRIVATE`), у них гараж виден. У профи вместо гаража есть счётчик «Объявления» (50, 99+). Если пользователь просит и их, это отдельный режим (см. раздел «Заметки о сайте»): считайте профили со счётчиком, но не смешивайте с частниками в одном листе без колонки «Тип».

## Скорость и нагрузка

- Пауза 0,6 с между запросами, не больше 400 запросов за запуск, не больше 6 страниц выдачи на модель. Это сознательно мягкий режим: капча приходит от резких всплесков, и один остановленный день дороже, чем пара лишних минут ожидания.
- Не запускайте несколько сборщиков параллельно в разных вкладках и не уменьшайте паузу.
- Не прогоняйте один и тот же запрос выдачи по несколько раз за сутки.

## Типичные сбои и что делать

| Симптом | Причина | Действие |
|---|---|---|
| Окно подтверждения от расширения / «No response to the permission prompt» | нет постоянного разрешения для auto.ru | остановиться, попросить выдать разрешение (один раз) |
| «Cannot attach to this target» | вкладка в служебном состоянии или редирект | новая вкладка через `tabs_create_mcp`, снова навигация на auto.ru |
| «tab group no longer exists» / «Tab not in Claude's tab group» | потеряна группа вкладок | `tabs_context_mcp` с `createIfEmpty: true`, новый tabId |
| `[BLOCKED: Cookie/query string data]` в выводе | в выводе есть URL с `?` | выводить пути без query (сборщик уже очищает через `split('?')[0]`) |
| «CDP Runtime.evaluate timed out after 45000ms» | синхронный скрипт слишком долгий | запускать фоном, как в Приложении A, и опрашивать |
| cards 0 на странице выдачи | изменилась вёрстка или показана капча | раздел «Заметки о сайте», проверить на странице вручную; при капче остановиться |
| found мало (0-2) | модель редкая или в регионе мало частников | расширить регион (`moskva` вместо области), увеличить `maxPages` до 10, добавить модель |
| `dupSkipped` растёт, `found` не растёт | выдача по тем же людям | сменить модель по ротации |

## Что скилл сознательно не делает

Не собирает телефоны и e-mail, не нажимает «Показать телефон», не пишет продавцам, не заходит в аккаунт, не использует Авито и Дром (на Авито расширение запрашивает разрешение на каждый шаг, на Дроме продавца в объявлении не видно; подробности в разделе «Заметки о сайте»), не обходит капчу и защиту.

## Как пользоваться приложениями ниже
Скилл самодостаточен: весь код лежит в этом файле. Python-скрипты (Приложения D, E и F) сохраните в рабочую папку как `build_xlsx.py`, `known_ids.py` и `setup_folder.py` (инструментом Write или командой `cat > build_xlsx.py <<'EOF' ... EOF`). JavaScript-блоки (A, B, C) вставляются прямо в `javascript_tool`, сохранять их не нужно.

## Заметки о сайте

Читайте этот раздел, если сборщик вернул 0 карточек или 0 профилей: вёрстка Авто.ру меняется, и селекторы могут устареть. Проверено 12.10.2026.

### URL
- Выдача по модели, только частники: `/{регион}/cars/{марка}/{модель}/used/?seller_group=PRIVATE&page=N` (регионы: moskva, moskovskaya_oblast, sankt-peterburg, krasnodarskiy_kray, sverdlovskaya_oblast, respublika_tatarstan, novosibirskaya_oblast ...). Одна страница = 37 карточек.
- Объявление: `/cars/used/sale/{марка}/{модель}/{id}-{хеш}/`
- Профиль частника: `/profile/{ID}/` (ID вида `QbCQf1Xr56eXo_TKLd9-Ig`); профиль профи-продавца: `/profile/{slug}/` (например `/profile/avg_motors/`).
- Страницы `/dealer/resellers/` и `/resellers/` — реклама сервиса для профи, списка продавцов там нет. Не заходите.

### Селекторы (CSS-классы с хешем, ищем по префиксу)
- карточка в выдаче: `[class*="ListingItemUniversal__snippet"]`
- блок продавца в карточке: `[class*="ListingItemUniversalSeller"]` (ник, рейтинг, метка «Проверенный продавец», город/адрес)
- в объявлении ссылка на профиль: `a[href*="/profile/"]`
- гараж в профиле: `[class*="PublicProfileGarageItem-"]` (одна машина = один элемент)
- метки в гараже: «Бывшая» (продана) и «Машина мечты» (желаемая, не в собственности). Обе в счёт не идут. Если встретите новую метку, которой нет в списке, не считайте такую машину автоматически: покажите её в колонке «Что в гараже» и отметьте в комментарии
- у профи-продавцов в профиле есть счётчик «Объявления» (50, 99+), у частников его нет, только гараж

### Что видно без входа в аккаунт
Профиль частника: имя, город, стаж на Авто.ру, гараж (марка, поколение, год, метки). Списка активных объявлений в профиле нет. Поэтому «машин в гараже без бывших и мечты» — это ОЦЕНКА числа машин продавца, а не точное число объявлений на продаже.

### Кодировка и прочее
- Авто.ру отдаёт UTF-8. (Дром отдаёт windows-1251 и требует TextDecoder, но Дром в этом скилле не используется.)
- Анонимные продавцы в выдаче подписаны «Частное лицо»; их профиль всё равно открывается по ссылке из объявления.
- `fetch` из вкладки с `credentials: 'include'` работает без капчи на скорости ~1 запрос в 0,6 с (проверено на ~1000 запросов за сессию).

### Что не работает / не делать
- Дром: объявление не содержит продавца (ни профиля, ни имени), по ссылкам перекупа не определить.
- Авито: расширение Claude in Chrome запрашивает разрешение на каждое действие, автономный запуск невозможен.
- «Показать телефон»: не нажимать, номера не собирать (см. раздел «Границы»).

## Автозапуск

Цель: компьютер включён, раз в день агент сам собирает небольшую порцию новых продавцов и сохраняет её в новый Excel-файл с номером и датой.

### Что должно быть настроено ОДИН РАЗ человеком (агент сам этого не сделает и не должен пытаться)
1. Chrome запущен, расширение Claude in Chrome установлено и подключено.
2. Для сайта auto.ru в расширении выдано постоянное разрешение (иначе каждое действие потребует подтверждения и автозапуск встанет). Проверка: агент выполняет `javascript_tool` на вкладке auto.ru, и никакого окна подтверждения не появляется.
3. Компьютер не уходит в сон: Windows — «Параметры → Система → Питание → Сон: никогда» (при питании от сети); macOS — «Не переходить в режим сна при питании от сети», либо `caffeinate -s`.
4. Приложение Claude на компьютере открыто и вошло в аккаунт; расписание создаётся как «запланированная задача» (scheduled task) в приложении, а не как cron внутри сессии.
5. У сессии есть доступ к рабочему столу (подключённая папка) или агент запускается прямо на компьютере. Папка `AutoRu_Resellers` создаётся скиллом сама (шаг 1); путь не нужно хранить вручную.

### Расписание
- 1 раз в день, в рабочее время (например 11:00 по Москве), порция: 3 модели x 8 продавцов, не больше 400 запросов.
- Не запускайте чаще и не увеличивайте порцию: это лишняя нагрузка на сайт и риск капчи без выигрыша.
- Модели меняйте по кругу (ротация), чтобы каждый день были новые объявления: список в разделе «Что нужно на входе».

### Текст запланированной задачи (скопируйте целиком)
«Запусти скилл avtoru-reseller-finder в режиме без вопросов (человека рядом нет). Критерии: из criteria.json в папке AutoRu_Resellers на рабочем столе, модели возьми следующие по ротации; если файла нет, значения по умолчанию. Создай новый файл NNN_resellers_ГГГГ-ММ-ДД.xlsx в этой папке. В конце напиши итог: критерии запуска, сколько добавлено, сколько дублей, были ли остановки, путь к файлу. Сообщения продавцам не отправляй.»

### Правила остановки (агент обязан следовать)
- Капча, статус 403/429, страница «Подтвердите, что вы не робот»: остановиться, записать в журнал «остановлено: капча», уведомить человека. Не повторять, не обходить.
- Окно подтверждения разрешения в расширении: остановиться и сообщить, что разрешение для auto.ru не выдано.
- 0 карточек на первой странице: вероятно, поменялась вёрстка. Прочитать раздел «Заметки о сайте», проверить селекторы на одной странице; если не помогло, остановиться и сообщить.
- Вкладка закрыта/группа вкладок потеряна: `tabs_context_mcp` с `createIfEmpty: true`, затем снова навигация на auto.ru. Старые tabId не использовать.

### Что агент пишет по итогам запуска
Одна короткая сводка: критерии запуска, добавлено N, дублей M, моделей K, остановки (если были), путь к новому файлу. Строка в листе «Журнал запусков» нового файла (её добавляет build_xlsx.py).

## Приложение A. Сборщик (вставить в javascript_tool целиком, предварительно заполнив CONFIG)
```javascript
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
```

## Приложение B. Опрос прогресса
```javascript
// Опрос прогресса (короткий вывод!). Вызывать после ожидания <=10 сек (computer wait).
JSON.stringify({ done: window.__M.done, stop: window.__M.stop, model: window.__M.state.model,
  found: window.__M.res.length, requests: window.__M.requests, dupSkipped: window.__M.dupSkipped,
  log: window.__M.log.slice(-2) })
```

## Приложение C. Выгрузка порциями
```javascript
// Выгрузка результатов порциями. Вывод инструмента обрезается примерно на 1-1,5 тыс. символов,
// поэтому выгружайте по 3-4 записи за вызов: меняйте START и END (по умолчанию 0..4).
// Формат строки: модель|ID профиля|ник|машин в гараже|из них без «бывших»|гараж|путь объявления
const START = 0, END = 4;
window.__M.res.slice(START, END).map(x => [x.q, x.pid, x.name.replace(/Москва.*|, .*/, '').slice(0, 20), x.gt, x.gc,
  x.g.slice(0, 120), x.l.replace('https://auto.ru/cars/used/sale/', '')].join('|')).join('\n')
```

## Приложение D. build_xlsx.py
```python
#!/usr/bin/env python3
"""Создаёт НОВЫЙ Excel-файл запуска с номером и датой: NNN_resellers_ГГГГ-ММ-ДД.xlsx.

Использование:
  python3 build_xlsx.py rows.txt /путь/к/папке [--run-note "Camry/Rio/Solaris, moskovskaya_oblast"]

rows.txt - строки из Приложения C (по одной на продавца, поля через |).
Папка    - папка на рабочем столе (её создаёт setup_folder.py). Номер NNN = следующий по порядку (001, 002, ...),
           дата = сегодняшняя; два запуска в один день дают два файла с разными номерами.
Дубли отсекаются по ID профиля среди ВСЕХ файлов папки (старых запусков тоже). Если новых нет, файл не создаётся.
Печатает JSON: file, added, duplicates_skipped. Телефонов в таблице нет и быть не должно.
"""
import sys, json, re, datetime
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

F = 'Arial'
HDR = ['№', 'Дата добавления', 'Запрос (модель)', 'Продавец (ник)', 'ID профиля', 'Ссылка на профиль',
       'Ссылка на объявление', 'Гараж: всего', 'Гараж: без «бывших»', 'Что в гараже', 'Уверенность',
       'Дубль?', 'Статус контакта', 'Дата контакта', 'Комментарий']
STATUSES = '"не писали,написали,ответил,отказ,не отвечает,не перекуп"'
FILE_RE = re.compile(r'^(\d{3})_resellers_(\d{4}-\d{2}-\d{2})\.xlsx$')
PROFILE_RE = re.compile(r'auto\.ru/profile/([^/\s?#]+)')

def parse(path):
    out = []
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        p = line.split('|')
        if len(p) < 7:
            continue
        q, pid, name, gt, gc, g, l = p[:7]
        out.append(dict(q=q, pid=pid.strip(), name=name, gt=int(gt), gc=int(gc), g=g, l=l.strip()))
    return out

def confidence(gc):
    # гараж не равен объявлениям на продаже, поэтому честно называем это оценкой
    return 'средняя (3+ машин)' if gc >= 3 else 'низкая-средняя (2 машины, проверить)'

def existing_ids(folder):
    """ID профилей из всех .xlsx папки (старый resellers.xlsx и файлы с номерами тоже)."""
    ids = set()
    for f in folder.glob('*.xlsx'):
        if f.name.startswith('~$'):
            continue
        try:
            ws = load_workbook(f)['Продавцы']
        except Exception:
            continue
        hdr = {str(c.value): c.column for c in ws[1] if c.value}
        col = hdr.get('ID профиля')
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if col and c.column == col and c.value:
                    ids.add(str(c.value).strip())
                elif not col:
                    s = (c.hyperlink.target if c.hyperlink and c.hyperlink.target else c.value)
                    m = PROFILE_RE.search(s) if isinstance(s, str) else None
                    if m:
                        ids.add(m.group(1))
    return ids

def main():
    rows_path, folder = sys.argv[1], Path(sys.argv[2]).expanduser()
    note = sys.argv[sys.argv.index('--run-note') + 1] if '--run-note' in sys.argv else ''
    folder.mkdir(parents=True, exist_ok=True)
    rows = parse(rows_path)
    seen = existing_ids(folder)
    today = datetime.date.today().isoformat()
    fresh, skipped = [], 0
    for x in rows:
        if x['pid'] in seen:
            skipped += 1; continue
        seen.add(x['pid']); fresh.append(x)
    if not fresh:
        print(json.dumps({'file': None, 'added': 0, 'duplicates_skipped': skipped}, ensure_ascii=False)); return
    nums = [int(m.group(1)) for n in (p.name for p in folder.iterdir()) for m in [FILE_RE.match(n)] if m]
    path = folder / f'{(max(nums) + 1 if nums else 1):03d}_resellers_{today}.xlsx'
    wb = Workbook(); ws = wb.active; ws.title = 'Продавцы'; ws.append(HDR)
    for i, x in enumerate(fresh, 1):
        prof = f"https://auto.ru/profile/{x['pid']}/"
        lst = 'https://auto.ru/cars/used/sale/' + x['l'].lstrip('/')
        ws.append([i, today, x['q'], x['name'], x['pid'], prof, lst, x['gt'], x['gc'], x['g'],
                   confidence(x['gc']), None, 'не писали', '', ''])
        ws.cell(i + 1, 6).hyperlink = prof; ws.cell(i + 1, 7).hyperlink = lst
    last = ws.max_row
    for r in range(2, last + 1):
        ws.cell(r, 12).value = f'=IF(COUNTIF($E$2:$E${last},E{r})>1,"ДУБЛЬ","ок")'
    for c in ws[1]:
        c.font = Font(name=F, bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
        c.alignment = Alignment(wrap_text=True, vertical='center')
    for row in ws.iter_rows(min_row=2, max_row=last):
        for c in row:
            c.font = Font(name=F, color='0563C1', underline='single') if c.column in (6, 7) else Font(name=F)
            c.alignment = Alignment(vertical='top', wrap_text=c.column in (10, 15))
    for col, w in zip('ABCDEFGHIJKLMNO', [5, 13, 18, 22, 26, 30, 36, 10, 12, 44, 22, 9, 16, 13, 28]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = f'A1:O{last}'
    dv = DataValidation(type='list', formula1=STATUSES, allow_blank=True); ws.add_data_validation(dv); dv.add('M2:M2000')
    log = wb.create_sheet('Журнал запусков')
    log.append(['Дата', 'Критерии и запрос', 'Найдено', 'Добавлено', 'Дублей пропущено'])
    log.append([today, note, len(rows), len(fresh), skipped])
    for c in log[1]:
        c.font = Font(name=F, bold=True)
    log.column_dimensions['B'].width = 70
    wb.save(path)
    sys.stdout.reconfigure(encoding='utf-8')
    print(json.dumps({'file': str(path), 'added': len(fresh), 'duplicates_skipped': skipped}, ensure_ascii=False))

if __name__ == '__main__':
    main()
```

## Приложение E. known_ids.py
```python
#!/usr/bin/env python3
"""Печатает ID профилей из ВСЕХ Excel-файлов папки одной строкой JSON-массива:
вставьте её в CONFIG.knownProfileIds сборщика (Приложение A).
Использование: python3 known_ids.py /путь/к/папке   (можно передать и один .xlsx)
Читает все .xlsx в папке (кроме временных ~$...), ищет ссылки/ID вида auto.ru/profile/<ID> на листе «Продавцы»."""
import sys, json, re
from pathlib import Path
from openpyxl import load_workbook

PROFILE_RE = re.compile(r'auto\.ru/profile/([^/\s?#]+)')

def known_ids(target):
    t = Path(target)
    files = [t] if t.is_file() else sorted(p for p in t.glob('*.xlsx') if not p.name.startswith('~$'))
    ids = []
    for f in files:
        try:
            wb = load_workbook(f)
            ws = wb['Продавцы']
        except Exception:
            continue
        hdr = {str(c.value): c.column for c in ws[1] if c.value}
        col = hdr.get('ID профиля')
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if col and c.column == col and c.value:
                    ids.append(str(c.value).strip())
                elif c.hyperlink and c.hyperlink.target:
                    m = PROFILE_RE.search(c.hyperlink.target)
                    if m and not col:
                        ids.append(m.group(1))
                elif isinstance(c.value, str) and not col:
                    m = PROFILE_RE.search(c.value)
                    if m:
                        ids.append(m.group(1))
    return list(dict.fromkeys(ids))

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    print(json.dumps(known_ids(sys.argv[1]) if len(sys.argv) > 1 else [], ensure_ascii=False))
```

## Приложение F. setup_folder.py
```python
#!/usr/bin/env python3
"""Находит рабочий стол, создаёт на нём папку AutoRu_Resellers и печатает JSON с состоянием.

Использование:
  python3 setup_folder.py                       # найти рабочий стол, создать папку, показать состояние
  python3 setup_folder.py --dir /путь/к/папке   # использовать готовую папку (например, смонтированную: $HOME/mnt/Desktop/AutoRu_Resellers)
  python3 setup_folder.py --save-criteria '{"region":"moskovskaya_oblast","models":[["toyota","camry"]],"per_model":8,"min_garage":2}'

Печатает: folder, desktop, writable, files (уже готовые файлы), next_number, criteria (последние сохранённые или null).
"""
import sys, os, json, re, datetime
from pathlib import Path

FOLDER_NAME = 'AutoRu_Resellers'
FILE_RE = re.compile(r'^(\d{3})_resellers_(\d{4}-\d{2}-\d{2})\.xlsx$')

def desktop_path():
    home = Path.home()
    if sys.platform == 'win32':
        try:  # настоящий путь рабочего стола (учитывает перенос в OneDrive)
            import ctypes
            from ctypes import wintypes
            buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
            if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0000, None, 0, buf) == 0 and buf.value:
                return Path(buf.value)
        except Exception:
            pass
    for cand in [home / 'Desktop', home / 'OneDrive' / 'Desktop', home / 'OneDrive' / 'Рабочий стол', home / 'Рабочий стол']:
        if cand.is_dir():
            return cand
    return None

def main():
    out = {}
    if '--dir' in sys.argv:
        folder = Path(sys.argv[sys.argv.index('--dir') + 1]).expanduser()
        out['desktop'] = None
    else:
        d = desktop_path()
        if d is None:
            print(json.dumps({'error': 'desktop_not_found', 'hint': 'рабочий стол не найден: спросите путь у пользователя и вызовите с --dir'}, ensure_ascii=False))
            sys.exit(2)
        folder = d / FOLDER_NAME
        out['desktop'] = str(d)
    folder.mkdir(parents=True, exist_ok=True)
    try:
        t = folder / '.write_test'; t.write_text('ok'); t.unlink(); out['writable'] = True
    except Exception as e:
        out['writable'] = False; out['write_error'] = str(e)
    cfg = folder / 'criteria.json'
    if '--save-criteria' in sys.argv:
        data = json.loads(sys.argv[sys.argv.index('--save-criteria') + 1])
        data['saved_at'] = datetime.date.today().isoformat()
        cfg.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    files = sorted(p.name for p in folder.iterdir() if FILE_RE.match(p.name))
    nums = [int(FILE_RE.match(n).group(1)) for n in files]
    out.update(folder=str(folder), files=files, next_number=(max(nums) + 1 if nums else 1),
               criteria=(json.loads(cfg.read_text(encoding='utf-8')) if cfg.exists() else None))
    sys.stdout.reconfigure(encoding='utf-8')
    print(json.dumps(out, ensure_ascii=False))

if __name__ == '__main__':
    main()
```
