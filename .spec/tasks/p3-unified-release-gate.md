# P3.1 — Контракт единого release-gate CreateOrder

**Дата:** 29.08.2026  
**Статус:** реализовано; спецификация подтверждена Verify/Release 30.08.2026  
**Граница:** этот документ описывает проверку. Он не выпускает, не импортирует и не изменяет рабочую книгу сам по себе.

## 1. Цель

Дать одну воспроизводимую команду, которая перед выпуском CreateOrder проверяет исходники, встроенную VBA-книгу, адресные сценарии, Word/Open XML и полный acceptance. При любой ошибке gate останавливается до изменения рабочей книги и не создаёт релизный артефакт.

## 2. Режимы

### Verify

- работает с указанной рабочей книгой или disposable-копией;
- выполняет preflight, сверку исходников с книгой и все доступные автоматические проверки;
- создаёт JSON и Markdown отчёты в одноразовой папке `Trash\release-gate-*`;
- не копирует результат в `CreateOrderReleases` и не изменяет рабочую `CreateOrder.xlsm`;
- возвращает `0`, если обязательные автоматические gate зелёные, `10` — ошибка preflight, `20` — рассинхронизация исходников/книги, `30` — упавший тест, `40` — ошибка упаковки/артефакта, `50` — ручной gate не подтверждён.

### Release

- сначала выполняет весь `Verify` на disposable-копии;
- перед изменением источника делает резервную копию с timestamp;
- собирает релиз только после `Verify=0`;
- проверяет выходной `.xlsm` read-only через Excel/Open XML;
- не выполняет импорт в рабочую книгу; импорт проверенных исходников в закрытую рабочую книгу остаётся отдельным установщиком с собственным backup;
- возвращает тот же код, а путь backup/release/report выводит явно.

## 3. Порядок gate

1. **Preflight Office:** `EXCEL.EXE` и `WINWORD.EXE` не запущены. Gate не завершает чужие процессы и не использует `Stop-Process`.
2. **Входы:** существуют исходная книга, каталог `CreateOrder.xlsm.modules`, `resources/customUI14.xml` и все явно перечисленные тесты; пути разрешаются в пределах проекта.
3. **Источник → книга:** для стандартных `.bas/.cls/.frm` сравниваются исходные тексты после удаления `Attribute VB_Name`; для UserForm дополнительно проверяются `.layout.csv` и наличие `.frx`; document/class modules не импортируются как обычные `.bas`.
4. **Структура/Open XML:** книга является ZIP, содержит `xl/vbaProject.bin`, обязательные листы/компоненты и не содержит обратных слэшей в ZIP-путях.
5. **Дизайнерские формы:** P0/P1/P2 designer-тесты подтверждают имена, типы, контейнеры, страницы и отсутствие runtime-геометрии/`Controls.Add`.
6. **Адресные тесты:** V2 preview/E2E, data-integrity fixture, Ribbon и справочники ФИЗО/тарифов/медалей/компактной формы.
7. **Сквозная регрессия:** `Test-PersonnelEvents.ps1`, `Test-PaymentsEnrollmentAcceptance.ps1`, FIO/ZP12 и доступные smoke-проверки.
8. **Word:** созданные DOCX существуют, не пусты, читаются Open XML и не содержат необработанных технических плейсхолдеров; визуальный просмотр остаётся owner/manual gate.
9. **Отчёт:** каждый gate имеет `id`, `status`, `duration_ms`, `command`, `exit_code`, `message`; персональные значения и содержимое кадровых строк в отчёт не попадают.
10. **Stop rule:** первый обязательный FAIL прекращает следующий изменяющий шаг. Предыдущие одноразовые тестовые копии сохраняются в `Trash`; рабочая книга и `CreateOrderReleases` не меняются.

## 4. Матрица существующих проверок

| Gate | Команда | Что доказывает |
|---|---|---|
| Personnel V2 designer | `tools/Test-PersonnelActionWizardV2Designer.ps1 -ExpectedActiveVersion V2` | 80 строк manifest, 3 страницы, V1 сохранена, V2 активна |
| Personnel V2 E2E | `tools/Test-PersonnelActionWizardV2Safe.ps1` | preview/confirm, два DOCX, duplicate-confirm и no-write cancel |
| Preview read-only | `tools/Test-PersonnelActionPreviewSafe.ps1` | builder не меняет шесть кадровых реестров |
| Data-integrity fixture | `tools/Test-PersonnelDataIntegritySafe.ps1` | clean=0; corrupt-категории; schema finding; хеши без мутаций |
| Integrity designer | `tools/Test-DataIntegrityCenterDesigner.ps1` | 11 design-time контролов, read-only output, Ribbon contract |
| Personnel history center | `tools/Test-PersonnelHistoryCenterSafe.ps1`, `tools/Test-PersonnelHistoryCenterDesigner.ps1` | read-only chronology, missing/ambiguous handling, 16 static controls |
| Grouped personnel order | `tools/Test-GroupedPersonnelOrderSafe.ps1` | пять событий, отдельные `§`, персональные выплаты, блокировка неполного основания, DOCX/Open XML |
| Personnel events | `Test-PersonnelEvents.ps1` | зачисление/перемещение/исключение и связи событий |
| Ribbon | `tools/Test-PersonnelRibbonSafe.ps1` | команды ленты и безопасная маршрутизация |
| Enrollment references | `tools/Test-EnrollmentFizoReferenceSafe.ps1`, `Test-EnrollmentTariffReferenceSafe.ps1`, `Test-EnrollmentMedalReferenceSafe.ps1` | справочники и расчётные ограничения |
| Enrollment compact UI | `tools/Test-EnrollmentCompactUiSafe.ps1` | сохранённая дизайнерская геометрия без перекрытий |
| Full acceptance | `Test-PaymentsEnrollmentAcceptance.ps1` | сквозная зачисление/выплаты/Word/регрессия |
| Text/validation | `Test-FIODeclension.ps1`, `Test-ZP12Validation.ps1` | вспомогательная текстовая и валидационная регрессия |

Скрипты с длительной COM-приёмкой запускаются последними среди автоматических тестов. Gate обязан показывать начало и завершение каждого теста, но не считать отсутствие ручного визуального просмотра доказанным автоматически.

## 5. Артефакты и коды

- `report.json` — машинно-читаемый итог с версиями, SHA-256 входов и списком gate;
- `report.md` — краткий отчёт для владельца: итог, время, backup, release (если создан), пропущенные ручные проверки и первый FAIL;
- `CreateOrder_Release_<timestamp>.xlsm` — только в режиме Release после нулевого кода;
- backup — копия до любой операции, затрагивающей книгу/релизный процесс;
- `0` — все обязательные автоматические gate пройдены;
- ненулевой код — выпуск не готов, причина и gate указаны в обоих отчётах.

## 6. Ручные/владелецские границы

Автоматический gate не подменяет:

- визуальный просмотр форм владельцем после ручной компоновки;
- визуальная сверка grouped DOCX с утверждённым PDF-образцом P5;
- утверждение НПА, ролей и расчётной матрицы для P6.

Такие пункты записываются как `MANUAL_REQUIRED`, а не как зелёный автоматический тест.

## 7. Фактическая проверка 30.08.2026

`Verify -SkipManual` с P5 завершился `exit=0` в отчёте
`Trash/release-gate-20260830-025544/report.md` после source/book sync, всех
адресных тестов P1–P5, полного acceptance и read-only проверки. Grouped-тест
зафиксировал пять событий, три параграфа, пять независимых выплат, негативные
проверки основания и `REQUIRES_DECISION`, а также DOCX/Open XML.

`Release -SkipManual` завершился `exit=0`, создал
`CreateOrderReleases/CreateOrder_Release_20260830_031131.xlsm`, backup
`CreateOrderBackups/release-gate-20260830-030504/CreateOrder.xlsm.before-release-gate.xlsm`
и отчёт `Trash/release-gate-20260830-030504/report.md`/`report.json`. Артефакт
прошёл ZIP/Open XML и Excel read-only. `-SkipManual` остаётся явным предупреждением,
а визуальная приёмка владельца не считается выполненной автоматически.

## 8. Приёмка P3

P3 считается готовым, когда:

1. `Verify` на чистой disposable-копии возвращает `0` и не меняет рабочую книгу;
2. намеренно повреждённая копия останавливается с ненулевым кодом и понятным первым FAIL;
3. `Release` создаёт один проверенный `.xlsm` только после зелёного gate;
4. повторный read-only/Open XML осмотр релиза проходит;
5. рабочий источник, `CreateOrder.xlsm` и `CreateOrderReleases` при неуспехе остаются без частичных изменений;
6. статус, UserGuide (если меняется пользовательский workflow) и Active State/Log обновлены.

## Следующий шаг

P3, P4 и P5 завершены. Следующий шаг — ручная визуальная сверка grouped DOCX;
после неё можно открыть owner gate P6 для утверждения юридических правил.
