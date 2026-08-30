# Runbook внедрения CreateOrder P1–P6 для Luna

**Дата:** 29.08.2026  
**Состояние:** P1.1–P1.6, P2.1–P2.4, P3.1–P3.4, P4.1–P4.4 и P5.1–P5.4 выполнены/импортированы; P6 остановлен на owner gate  
**Базовая ветка:** `main` (`c41e93b`)  
**Предшествующее условие:** P0 завершён: `frmPersonnelActionWizardV2` активна, V1 сохранена для отката.

## Original Request

Ну, сделай все необходимое. Когда все подготовишь все эти моменты, ты мне в итоговом сообщении напишешь, что Луне надо будет прочитать. И я уже в чате переключу на младшую модель, чтобы он уже начал внедрять эти фичи.

## Настройки выполнения

- **Исполнитель:** Luna выполняет только одну задачу за раз и останавливается на каждом owner gate.
- **Тесты:** обязательны. Любая новая логика сначала получает адресный тест в одноразовой копии книги.
- **Логирование:** `DEBUG` для вычислений без персональных данных, `INFO` для выполненных операторских действий, `WARN` для блокирующих данных, `ERROR` с кодом/контекстом ошибки без ФИО, номеров, документов и путей к персональным файлам.
- **Документация:** после каждого законченного этапа обновить `docs/PROJECT_STATUS.md`, `.spec/PROJECT_CONTEXT.md` (Active State и Log), а для видимой функции — `docs/UserGuide.md`.
- **Кодировка:** `.bas`, `.frm`, `.cls` — CP1251 и CRLF; `.frx` вручную не редактировать; все новые русские подписи — через `ModuleLocalization` и лист `Localization`.
- **Книга:** перед Excel COM проверить висящие `EXCEL.EXE`; рабочую `CreateOrder.xlsm` не изменять, пока адресная проверка не прошла на копии. Импортировать только после закрытия Excel и с резервной копией.
- **Сторонние изменения:** не трогать существующие `graphify-out/`, `.agents/`, `.ai-factory*`, `state/` и любые файлы, не созданные задачей.

## Общий протокол Luna для каждой задачи

1. Прочитать этот runbook, `.spec/PROJECT_CONTEXT.md`, `docs/PROJECT_STATUS.md` и спецификацию текущего этапа.
2. Выполнить `git status --short`; сохранить и не индексировать чужие изменения.
3. До изменения создать/уточнить отдельную спецификацию `.spec/tasks/pN-<slug>.md` с границами, полями данных и приёмкой.
4. Изменять только перечисленные в задаче файлы. Если потребовался другой файл — остановиться и объяснить причину.
5. Запустить адресный тест на временной копии книги, затем обязательную регрессию этапа.
6. Проверить, что рабочая книга закрыта; создать резервную копию; безопасно импортировать проверенные VBA/формы; повторить адресный тест.
7. Обновить документацию, показать `git diff --check`, перечень затронутых файлов, команды/результаты и известные ограничения.
8. Не выполнять commit, merge, push или импорт в книгу, если это не указано в отдельной задаче владельца.

## Карта зависимостей

```text
P0 (готов) → P1
           ├→ P2 → P4
            └→ P3 → P5 [готов]
                 ↑
             P2 + P3 → P6 [owner gate: НПА, роли, переходы]
```

P1, P2 и P3 можно планировать независимо, но выполнять последовательно по приоритету P1 → P2 → P3. P4 начинает только после зелёного P2. P5 owner gate закрыт и выполнен; P6 остаётся условным этапом до утверждения юридической модели.

---

## P1. Подтверждение кадрового изменения «до → после»

### P1.1 — Зафиксировать контракт предпросмотра

- **Изменить:** создать `.spec/tasks/p1-personnel-action-confirmation.md`; при необходимости только Markdown-план формы.
- **Изучить:** `frmPersonnelActionWizardV2.frm` (`WriteValues`, `SaveAction`, `CurrentSignature`), `mdlPersonnelEvents.bas` (`GetCurrentPersonnelState`, `SavePersonnelEventInput`, `SavePersonnelEvent`), `mdlPersonnelEventOrderExport.bas`.
- **Сделать:** перечислить отображаемые поля TRANSFER/EXCLUSION, правила diff, статусы выплат, предупреждения и точное правило «отмена не пишет ни в один реестр».
- **Запрет:** не использовать `ExportPersonnelEventOrder` для предпросмотра: он требует EventID, создаёт `DocumentRegistry` и меняет статус события.
- **Готовность:** спецификация содержит таблицу «поле / before / after / источник / отображение Word», а также критерии совпадения итогового Word с экраном.

### P1.2 — Чистый расчёт draft и diff без записи

**Фактический статус 29.08.2026:** выполнено. Добавлен `CreateOrder.xlsm.modules/mdlPersonnelActionPreview.bas`; адресный `tools/Test-PersonnelActionPreviewSafe.ps1` подтвердил TRANSFER/EXCLUSION/invalid draft и нулевую мутацию кадровых реестров. Безопасный импорт выполнен `tools/Install-PersonnelActionPreviewSafe.ps1` после staging-проверки и backup. P1.3 и P1.4 также выполнены; следующий незакрытый пункт P1 — P1.5.

- **Изменить:** `CreateOrder.xlsm.modules/mdlPersonnelEvents.bas`; при росте модуля создать `mdlPersonnelActionPreview.bas`.
- **Сделать:** добавить публичный builder `BuildPersonnelActionPreview` (имя можно уточнить в спецификации), который принимает черновые значения, читает только текущее состояние и возвращает Dictionary/Collection: до/после, список изменённых полей, прогноз назначений/прекращений выплат, предупреждения и данные для Word.
- **Ограничение:** builder не вызывает `SetPersonnelWizardValue`, `SavePersonnelEventInput`, `SavePersonnelEvent`, `SaveStateSnapshot`, `SavePaymentAssignments`, `SaveDocumentRecords`, `SetPersonnelEventStatus` и не создаёт листов.
- **Логирование:** `DEBUG preview-built` только с event type, количеством полей/выплат/предупреждений; `ERROR preview-failed` с кодом ошибки.
- **Проверка:** новый адресный PowerShell-тест создаёт fixture TRANSFER и EXCLUSION, дважды строит preview и доказывает неизменность значений/числа строк всех кадровых реестров.

### P1.3 — Единый источник существенного текста для preview и Word

**Фактический статус 29.08.2026:** выполнено. Добавлен чистый `mdlPersonnelOrderText.bas`; preview вкладывает его в `order_projection.text_model`, а индивидуальный экспорт использует `WriteTransferOrderFromModel`/`WriteExclusionOrderFromModel` без изменения API. V2 E2E подтвердил точное совпадение общего transfer core text с DOCX и существующую кадровую регрессию. P1.4 после решения владельца выполнен; следующий пункт P1.5 — двухфазное сохранение.

- **Изменить:** `mdlPersonnelEventOrderExport.bas`; возможно новый чистый модуль `mdlPersonnelOrderText.bas`.
- **Сделать:** вынести из `WriteTransferOrder`/`WriteExclusionOrder` формирование существенных строк в чистую модель/функцию. Финальный экспорт и P1 preview используют одинаковые поля/формулировки, но только экспорт создаёт Word и регистрирует документ.
- **Не менять:** существующий API `ExportPersonnelEventOrder(eventID)` и индивидуальный Word-экспорт.
- **Проверка:** fixture подтверждает, что перечисленные на preview дата, должность, место убытия, основание, отпускно-финансовые сведения и выплаты присутствуют в созданном DOCX; использовать существующий XML/DOCX-паттерн из `tools/Test-PersonnelActionWizardV2Safe.ps1`.

### P1.4 — Design-time экран подтверждения и локализация

**Фактический статус 29.08.2026:** выполнено по решениям владельца. В `frmPersonnelActionWizardV2` добавлена отдельная страница `pgPreview` с четырьмя read-only/multiline полями существенных строк и кнопками `Подтвердить`/`Отмена`; добавлены ключи `personnel.preview.*` в `ModuleLocalization.bas` и `Localization`. Designer-test подтвердил 80 manifest-строк, 3 страницы, уникальные контролы, отсутствие runtime-геометрии и лог-маркеры confirmation. Ручная художественная подгонка `pgPreview` выполнена владельцем; экспорт зафиксировал `geometryChanges=5`.

- **Изменить:** `frmPersonnelActionWizardV2.frm/.frx/.layout.csv` только через дизайнер/штатный экспортёр, `ModuleLocalization.bas`, лист `Localization`, `tools/Export-PersonnelActionWizardV2Designer.ps1`, `tools/Test-PersonnelActionWizardV2Designer.ps1`.
- **Сделать:** добавить статичный блок/страницу с read-only before/after, выплатами, предупреждениями, кнопками `Подтвердить` и `Отмена`.
- **Owner gate:** если нужно менять геометрию, Luna создаёт/экспортирует набор контролов с именами и функциональностью, но останавливается до художественного выравнивания. Владелец вручную располагает элементы в VBE; Luna затем только экспортирует и проверяет.
- **Запрет:** не применять `Controls.Add`, не присваивать `Left`, `Top`, `Width`, `Height` в коде, не менять `(Name)`, тип или контейнер действующих контролов.
- **Логирование:** `INFO confirmation-opened/cancelled/confirmed`; без персональных значений.
- **Проверка:** designer-test подтверждает уникальные контролы, страницы и отсутствие runtime-геометрии.

### P1.5 — Подключить двухфазное сохранение

**Фактический статус 29.08.2026:** выполнено. Форма собирает transient draft и открывает preview по «Сохранить»; только `ConfirmConfirmationPreview` вызывает запись, `CancelConfirmationPreview` не меняет реестры, duplicate-confirm идемпотентен, stale draft сбрасывается, экспорт доступен только в `SAVED`. Обновлённый установщик заменяет код существующей UserForm без `Attribute`-строк, сохраняя owner geometry и `.frx`. Staging/рабочая post-install проверка и `PERSONNEL_ACTION_V2_E2E_OK` зелёные; P1.6 автоматическая приёмка также завершена.

- **Изменить:** `frmPersonnelActionWizardV2.frm`, `mdlPersonnelEvents.bas`, локализация.
- **Сделать:** текущая кнопка сохранения сначала собирает transient draft и открывает preview. Только `Подтвердить` один раз вызывает `WriteValues` и `SaveAction`; `Отмена` очищает transient draft и возвращает к редактированию. Экспорт недоступен до успешного подтверждения.
- **Краевые случаи:** повторное подтверждение, изменённое поле после открытия preview, ошибка расчёта, попытка экспортировать несохранённое, меню формы и режим ENROLLMENT (P1 не меняет ENROLLMENT).
- **Проверка:** расширить `tools/Test-PersonnelActionWizardV2Safe.ps1`: cancel не меняет `PersonnelEvents`, snapshots, assignments, documents; confirm создаёт ровно одно событие; повторный click не создаёт второе.

### P1.6 — Приёмка, безопасный импорт и документация

**Фактический статус 29.08.2026:** автоматическая часть выполнена. На рабочей книге прошли designer-test, preview smoke, V2 E2E (`PERSONNEL_ACTION_V2_E2E_OK`), `Test-PersonnelActionWizardSafe.ps1`, `Test-PersonnelEvents.ps1`, `Test-PersonnelRibbonSafe.ps1` и полный `Test-PaymentsEnrollmentAcceptance.ps1` (`ACCEPTANCE_SMOKE_OK`). Документация обновлена; отдельное визуальное открытие preview владельцем остаётся рекомендуемой ручной проверкой.

- **Изменить:** адресные `tools/Test-PersonnelActionWizardV2Safe.ps1`, при необходимости `tools/Test-PersonnelActionWizardSafe.ps1`, документация.
- **Проверки:** designer-test; P1 address test; `tools/Test-PersonnelActionWizardSafe.ps1`; `Test-PersonnelEvents.ps1`; полный `Test-PaymentsEnrollmentAcceptance.ps1` с `ACCEPTANCE_SMOKE_OK`; ручное открытие preview без сохранения данных.
- **Внедрение:** резервная копия, импорт в закрытую книгу, повтор критичного адресного теста; обновить UserGuide с операторским порядком «проверить → подтвердить → экспортировать».
- **Готовность P1:** отмена не создаёт артефактов; подтверждение создаёт одну запись; финальный Word согласован с preview; P0 V2 и V1-regression зелёные.

---

## P2. Центр целостности бизнес-данных

### P2.1 — Спецификация реестров и каталог findings

- **Фактический статус 29.08.2026:** выполнено. Создана `.spec/tasks/p2-business-data-integrity.md` с контрактом Finding, каталогом проверок и fixture/severity-матрицей; автоисправление и `HealWorkbookStructure` исключены.
- **Сделать:** зафиксировать Finding (`Severity`, `Category`, `EntityType`, `EntityID`, `Message`, `SuggestedAction`) и полный список проверок: ID/дубликаты, Event→Employee, Event→snapshots, snapshots→event/employee/kind, assignments→employee/event/act, docs→event/file, current state→employee/LastEventID, chronology, exclusion/active payments, связь со `Штат`, ссылки НПА/классификации.
- **Ограничение:** первая версия только read-only; `HealWorkbookStructure` и любые автоисправления исключены.
- **Готовность:** каждой категории соответствует fixture и ожидаемая severity.

### P2.2 — Детерминированный сканер реестров

**Фактический статус 29.08.2026:** выполнено. `mdlPersonnelDataIntegrity.bas` установлен в рабочую книгу; `tools/Test-PersonnelDataIntegritySafe.ps1` возвращает `PERSONNEL_DATA_INTEGRITY_OK` (clean=0, corrupt categories покрыты, schema finding, нулевая мутация). Сканирование реальной книги read-only выявляет baseline 8 findings (6 ошибок, 2 предупреждения), что зафиксировано как диагностический результат, а не автоисправление.

- **Изменить:** новый `mdlPersonnelDataIntegrity.bas`, минимально `mdlDataValidation.bas`.
- **Сделать:** независимые функции сканирования по registry sheets, без `Select`, `Activate`, MessageBox и любых записей. Возвращать findings и статистику.
- **Логирование:** `INFO integrity-scan-complete` с числом сущностей/findings; `WARN integrity-finding` без PII; `ERROR integrity-scan-failed`.
- **Проверка:** `tools/Test-PersonnelDataIntegritySafe.ps1` на одноразовой копии делает снимки/хеши всех реестров до/после, seeding ровно одного дефекта каждого вида, доказывает точную категорию и нулевую мутацию.

### P2.3 — Операторский центр и Ribbon

**Фактический статус 29.08.2026:** выполнено. Design-time `frmDataIntegrityCenter` содержит 11 статических контролов и фильтры severity/category; локализация, Ribbon-кнопка и handler подключены. `tools/Test-DataIntegrityCenterDesigner.ps1` зелёный; рабочий импорт выполнен после staging и backup.

- **Изменить:** design-time `frmDataIntegrityCenter.frm/.frx`, `mdlRibbonHandlers.bas`, при необходимости Ribbon XML/локализация, тест дизайнерской формы.
- **Сделать:** запуск scan, фильтр severity/category, read-only просмотр findings, явное пояснение «исправление не выполняется». Не показывать персональные значения в журнале; в UI раскрывать только столько, сколько нужно оператору.
- **Проверка:** открыть центр на эталонной и повреждённой копиях; подтвердить, что закрытие/повторный scan ничего не меняют.

### P2.4 — Регрессия и внедрение P2

**Фактический статус 29.08.2026:** выполнено. Пройдены P2 safe-test, integrity designer-test, кадровый V2 designer/E2E, personnel events, Ribbon и полный `Test-PaymentsEnrollmentAcceptance.ps1` (`ACCEPTANCE_SMOKE_OK`). Документация обновляется этим чекпойнтом; визуальный просмотр центра остаётся ручной проверкой владельца.

- **Проверки:** P2 safe-test; `Test-PersonnelEvents.ps1`; `tools/Test-PersonnelActionWizardV2Safe.ps1`; полный acceptance.
- **Документация:** UserGuide описывает диагностику как read-only и путь передачи finding на исправление.
- **Готовность P2:** эталон не имеет ложных ошибок; все заявленные fixture находки видны; ни один лист/реестр не изменён проверкой.

---

## P3. Единый release gate

### P3.1 — Контракт gate и матрица существующих проверок

**Фактический статус 29.08.2026:** выполнено. Создан `.spec/tasks/p3-unified-release-gate.md` с режимами Verify/Release, кодами возврата, stop rules, матрицей тестов, артефактами и ручными границами. Следующий пункт — P3.2.

- **Изменить:** создать `.spec/tasks/p3-unified-release-gate.md`.
- **Сделать:** описать режимы `Verify` и `Release`, входы/выходы, коды возврата, порядок gate, временные копии, отчёт и stop rules. Сопоставить существующие `Test-*.ps1` с целями и временем выполнения.
- **Ограничение:** `Build-Release.ps1` пока является упаковщиком; он не должен выпускать файл до зелёных gate.

### P3.2 — Оркестратор и preflight

**Фактический статус 29.08.2026:** выполнено. `tools/Invoke-CreateOrderReleaseGate.ps1` реализует preflight, disposable copy, source/book sync, stop rules, отчёты и коды возврата; `Build-Release.ps1` делегирует ему Release по умолчанию, а `-SkipGate` используется только внутренним шагом упаковки. Office-процессы не завершаются принудительно.

- **Изменить:** новый `tools/Invoke-CreateOrderReleaseGate.ps1`, минимальная делегация из `Build-Release.ps1`.
- **Сделать:** параметры workbook/source/output/mode/skip-manual; fail-fast при открытом Excel/Word; резервная копия; disposable copy; хронологический запуск тестов; структурированный JSON/Markdown отчёт.
- **Запрет:** не завершать пользовательские Office-процессы, не импортировать VBA в рабочую книгу, не удалять существующий релиз при сбое.
- **Проверка:** отсутствие книги, открытый Office и невалидный параметр дают ненулевой код, отчёт и ноль изменений рабочей книги.

### P3.3 — Сверка исходников, тесты, Word/OpenXML gates

**Фактический статус 29.08.2026:** выполнено. Gate прошёл полный список designer/E2E, Ribbon, P2 integrity, справочники enrollment, кадровые события, FIO, ZP12 и `Test-PaymentsEnrollmentAcceptance.ps1`; source/book sync учитывает служебные Attribute и кодировки CP1251/UTF-8. Негативный mismatch остановился до тестов с `exit=20`.

- **Изменить:** gate и точечно существующие тесты только для параметра пути к disposable copy.
- **Сделать:** COM-экспорт компонентов с нормализацией CP1251/CRLF и сравнение с исходниками; document modules сравнивать через `CodeModule`, не через Import. Запускать V2 designer/E2E, ribbon, compact/tariff/fizo/medal, personnel events, FIO и полный acceptance. Проверить DOCX и ZIP: обязательные части, `xl/vbaProject.bin`, отсутствие `\\` в ZIP entry, открытие в Excel read-only.
- **Проверка:** намеренный mismatch исходника и намеренное падение теста прекращают релиз до упаковки и сохраняют backup.

### P3.4 — Выпуск на чистой копии

**Фактический статус 29.08.2026:** выполнено. Release завершился `exit=0`, создал проверенный `CreateOrderReleases/CreateOrder_Release_20260829_231525.xlsm`, JSON/Markdown-отчёт и backup `CreateOrderBackups/release-gate-20260829-230934`. Артефакт прошёл Open XML и Excel read-only; рабочая книга не пересобиралась. `-SkipManual` явно оставлен как предупреждение о ручной визуальной проверке.

- **Сделать:** зелёный `Release` на чистой копии, затем проверка созданного файла; рабочая книга не сохранена/не изменена.
- **Готовность P3:** одна команда воспроизводимо даёт exit 0, артефакт, backup, machine-readable report и доказательство Excel/OpenXML; все negative fixtures безопасно останавливаются.

---

## P4. Центр истории и документов сотрудника

### P4.1 — Контракт read-only просмотра

**Фактический статус 29.08.2026:** выполнено. Контракт зафиксирован в
`.spec/tasks/p4-personnel-history-document-center.md`: read-only read-model,
точный поиск и три явных действия без скрытой записи.

- **Изменить:** создать `.spec/tasks/p4-personnel-history-document-center.md`.
- **Изучить:** `mdlPersonnelHistory.bas` и существующий лист `PersonnelHistory`.
- **Сделать:** определить, что новая форма читает registry sheets без записи; legacy sheet допустимо оставлять как совместимый derived view. Зафиксировать явные действия: Open document, repeat export, prepare correction.

### P4.2 — Read model и безопасное открытие документа

**Фактический статус 29.08.2026:** выполнено. Добавлен
`mdlPersonnelHistoryCenter.bas` с детерминированным отчётом по кадровым
реестрам, проверкой локального пути документа, повторным экспортом и
подготовкой корректировки; missing document и нарушенная связь возвращают
понятные ошибки/предупреждения.

- **Изменить:** `mdlPersonnelHistory.bas` либо новый `mdlPersonnelHistoryCenter.bas`.
- **Сделать:** поиск/хронология событий, snapshots, assignments, documents, StaffStateSyncLog; `OpenDocument` проверяет наличие файла и локальный разрешённый путь. Отсутствующий файл — WARN, не реконструировать его автоматически.
- **Запрет:** просмотр не вызывает Save, correction/export выполняются только по явной команде.

### P4.3 — Design-time форма и подключение входов

**Фактический статус 29.08.2026:** выполнено. Созданы
`frmPersonnelHistoryCenter.frm/.frx/.layout.csv` с 16 статическими контролами,
локализация и Ribbon-кнопка `openPersonnelHistoryCenter`; runtime `Controls.Add`
и runtime-геометрия отсутствуют. `Test-PersonnelHistoryCenterDesigner.ps1` и
`Test-PersonnelRibbonSafe.ps1` зелёные.

- **Изменить:** `frmPersonnelHistoryCenter.frm/.frx`, `mdlRibbonHandlers.bas`, `frmPersonnelActionWizardV2.frm` только если нужен явный переход, локализация.
- **Проверка:** дизайнерская проверка, поиск с нулём/одним/несколькими совпадениями, browse без мутации, missing document.

### P4.4 — E2E истории

**Фактический статус 30.08.2026:** адресная fixture-проверка выполнена:
`Test-PersonnelHistoryCenterSafe.ps1` подтвердил хронологию, read-only browse,
ошибки missing/ambiguous и отсутствие мутаций; установщик завершён после
backup `CreateOrderBackups/personnel-history-center-installed-20260829-234654`.
Полный Verify и Release-gate после P4 завершились `exit=0`; актуальный артефакт
`CreateOrderReleases/CreateOrder_Release_20260830_003625.xlsm` прошёл Open XML и
Excel read-only. `-SkipManual` оставляет визуальную приёмку владельца отдельной.

- **Изменить:** `tools/Test-PersonnelHistoryCenterSafe.ps1`.
- **Проверка:** fixture с событиями/документами; снимок реестров до/после browse; повторный экспорт имеет ожидаемую регистрацию; коррекция создаётся только после отдельного save оператора.
- **Готовность P4:** просмотр не меняет реестры; документ открывается/ошибка понятна; экспорт и correction остаются явными действиями.

---

## P5. Единый многопараграфный/групповой приказ — завершён

### P5.0 — Owner gate (обязательная остановка)

**Фактический статус 30.08.2026:** owner gate P5 закрыт решениями владельца,
зафиксированными в `.spec/tasks/p5-grouped-personnel-order.md`. Утверждён весь
PDF-образец отображения; параграфы создаются отдельно по присутствующим
аналогичным категориям событий; выплаты разных сотрудников никогда не
объединяются; выбранная выплата без основания подсвечивается и блокирует пакет,
невыбранная выплата не проверяется; строки читаются в порядке заведения.

Согласованные значения `DocumentType`, статуса и legacy-ссылки зафиксированы и
реализованы в P5-спецификации; P6 по-прежнему не активируется автоматически.

### P5.1 — Спецификация данных и обратной совместимости — готово

- **Изменить:** создать `.spec/tasks/p5-grouped-personnel-order.md`.
- **Сделать:** добавить модель связей `DocumentEventLinks` (`DocumentID`, `EventID`, `ParagraphNo`, `Role`, `CreatedAt`); не записывать несколько EventID в существующий `DocumentRegistry.EventID`; индивидуальный export не менять.

### P5.2 — Детектор совместимости и выборка — готово

- **Изменить:** новый `mdlGroupedPersonnelOrderExport.bas`; pure helper при необходимости.
- **Сделать:** выбор только сохранённых событий из таблицы в порядке заведения,
  группировка по первой встреченной категории события, персональная проверка
  выплат и диагностические причины отказа; при ошибке не создавать частичный
  DOCX/реестр.

### P5.3 — Форма и Word renderer — готово

- **Изменить:** design-time `frmGroupedPersonnelOrderWizard`, локализация, `mdlPersonnelEventOrderExport.bas` только для общего безопасного formatting helper.
- **Сделать:** параграфы `§`, согласованные шаблоны, единый DocumentRegistry record и отдельные DocumentEventLinks.

### P5.4 — Grouped E2E и регрессия индивидуального экспорта — готово

- **Изменить:** `tools/Test-GroupedPersonnelOrderSafe.ps1`.
- **Проверка:** события одной категории дают один `§`, категории и сотрудники
  сохраняют порядок таблицы, выплаты не объединяются, включённая выплата без
  основания подсвечивается/блокирует пакет, невыбранная не блокирует; связи
  полные; individual export не регрессирует; DOCX визуально проверен.
- **Готовность P5:** owner gate закрыт, P3 зелёный; ни одна несовместимая группа не создаёт файл/реестр. Адресный тест и общий Verify/Release завершены `exit=0`.

---

## P6. Жизненный цикл НПА и симулятор правил — условный этап

### P6.0 — Owner gate (обязательная остановка)

**Фактический статус 30.08.2026:** подготовлена
`.spec/tasks/p6-legal-rule-lifecycle-simulation.md`. Зафиксированы наблюдаемые
поля `LegalActs`/`PaymentRules` и вопросы по статусам, ролям, датам, приоритетам
и симуляции; юридические ответы не придуманы, код не начинался.

До кода владелец утверждает: роли и матрицу переходов `DRAFT → REVIEWED → ACTIVE → RETIRED`, юридические реквизиты/даты, политику исправления старых данных и конкретные правила расчёта. Luna не придумывает проценты, НПА или права ролей.

### P6.1 — Спецификация версий и симуляции

- **Изменить:** создать `.spec/tasks/p6-legal-rule-lifecycle-simulation.md`.
- **Сделать:** additive schema для версий/аудита (`RuleLifecycleAudit`, derived `RuleSimulationResults`), матрица переходов, as-of semantics, неизменность истории и критерии no-op симуляции.

### P6.2 — Модуль lifecycle и совместимый выбор ACTIVE-версий

- **Изменить:** `mdlLegalActs.bas`, `mdlPersonnelAllowanceRules.bas`, новый `mdlRuleLifecycle.bas`, инфраструктура в `mdlPersonnelEvents.bas`.
- **Сделать:** validate transition/role, audit, поиск только ACTIVE и effective версии. Сначала добавить adapter, сохраняющий текущие подтверждённые правила и `MOBILIZED` 158 000 без регрессии.
- **Проверка:** invalid transition запрещён; DRAFT/REVIEWED/RETIRED не влияют на расчёт; ACTIVE влияет только в своём периоде.

### P6.3 — Чистая симуляция

- **Изменить:** новый `mdlRuleSimulation.bas`.
- **Сделать:** in-memory сравнение затронутых сотрудников/назначений/неоднозначностей до activation, без записи реестров. Результат может быть показан в derived форме/листе только как явный отдельный отчёт.
- **Проверка:** хеш реестров до/после симуляции совпадает; fixture показывает ожидаемые изменения и неоднозначности.

### P6.4 — Design-time администрирование, тесты и выпуск

- **Изменить:** `frmRuleLifecycleCenter.frm/.frx`, Ribbon/localization, `tools/Test-RuleLifecycleSimulationSafe.ps1`, документация.
- **Проверка:** роли, переходы, version history, no-op simulation, baseline regression, P2 integrity, P3 release gate.
- **Готовность P6:** только ACTIVE меняет расчёт; симуляция не меняет данные; аудит воспроизводим; юридические входы владельца указаны в спецификации.

## Контрольные точки коммитов

1. После P1: `feat: add personnel action confirmation preview`.
2. После P2: `feat: add business data integrity center`.
3. После P3: `build: add unified CreateOrder release gate`.
4. После P4: `feat: add personnel history and document center`.
5. После P5 (только после owner gate): `feat: add grouped personnel order export`.
6. После P6 (только после owner gate): `feat: add legal rule lifecycle and simulation`.

Перед каждым коммитом: адресные тесты зелёные, полный relevant regression зелёный, `git diff --check`, документация актуальна, рабочая книга импортирована только по правилам проекта. Merge/push выполняются лишь по явной команде владельца.

## Что считать итогом очереди

- P1–P4 готовы и внедрены; единый post-P4 release-gate зелёный.
- P5 ждёт утверждённые образцы Word и правила группировки.
- P6 ждёт утверждённые НПА, роли и матрицу переходов.
- Нельзя считать автоматические тесты заменой ручной визуальной проверки Excel/Word; такой результат всегда указывается отдельно.
