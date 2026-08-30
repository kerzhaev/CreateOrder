# P1: подтверждение кадрового изменения «до → после»

**Статус:** P1.1–P1.6 реализованы и импортированы; автоматическая приёмка зелёная, ручная визуальная проверка preview рекомендована владельцу  
**Дата:** 29.08.2026  
**Зависимость:** P0 завершён; активная форма — `frmPersonnelActionWizardV2`  
**Область:** только `TRANSFER` и `EXCLUSION`; `ENROLLMENT` остаётся в текущем потоке

**Фактическое состояние P1.2–P1.4 (29.08.2026):** `mdlPersonnelActionPreview.bas` содержит read-only `BuildPersonnelActionPreview(draftValues)`, а `mdlPersonnelOrderText.bas` — общий чистый DTO/text model для preview и индивидуального Word. `tools/Test-PersonnelActionPreviewSafe.ps1` проверяет отсутствие мутаций; `tools/Test-PersonnelActionWizardV2Safe.ps1` проверяет preview до SaveAction и точное совпадение transfer core text с DOCX. Design-time `frmPersonnelActionWizardV2` теперь содержит страницу `pgPreview` и четыре read-only/multiline поля для существенных строк, кнопки `btnPreviewConfirm`/`btnPreviewCancel`, локализацию `personnel.preview.*` и confirmation-логирование. Designer-test подтвердил 80 manifest-строк, 3 страницы и отсутствие runtime-геометрии; форма импортирована в рабочую книгу после staging и backup `CreateOrderBackups/personnel-action-v2-installed-20260829-202814/CreateOrder.before-personnel-action-v2-install.xlsm`. Ручная художественная подгонка `pgPreview` выполнена владельцем и может продолжаться отдельно.

**Фактическое состояние P1.5 (29.08.2026):** `frmPersonnelActionWizardV2` собирает transient draft и заполняет `pgPreview`; `ConfirmConfirmationPreview` — единственная точка, которая вызывает сохранение, `CancelConfirmationPreview` не меняет реестры, а повторное подтверждение после `SAVED` возвращает тот же EventID. Экспорт заблокирован до `SAVED`, а изменение draft после открытия preview сбрасывает устаревший просмотр. Установщик обновляет существующую UserForm через очищенный `CodeModule.AddFromString` без `Attribute`-строк, поэтому ручная геометрия и `.frx` сохраняются. Staging и рабочий импорт выполнены с backup `CreateOrderBackups/personnel-action-v2-installed-20260829-202814/CreateOrder.before-personnel-action-v2-install.xlsm`; owner geometry экспортирована с `geometryChanges=5` и backup `CreateOrderBackups/personnel-action-v2-owner-layout-20260829-202943`. Адресный V2 E2E подтвердил cancel/confirm/duplicate-confirm и два DOCX. Автоматическая часть P1.6 завершена; отдельное визуальное открытие preview владельцем остаётся рекомендуемой ручной проверкой.

## Цель

До записи кадрового события оператор должен увидеть вычисленный результат изменения, проверить его и явно подтвердить сохранение. Предпросмотр должен быть построен из transient draft и текущего состояния сотрудника, не создавать EventID и не менять ни один служебный реестр. После подтверждения сохранение и последующий Word-экспорт используют те же существенные данные.

## Не входит в P1

- новые проценты, НПА, нормативные формулировки или правила расчёта;
- исправление противоречий в существующих реестрах (это P2);
- групповой Word-приказ (P5);
- автоматическое исправление данных;
- изменение ручной геометрии существующих контролов без owner gate;
- изменение поведения мастера зачисления `ENROLLMENT`.

## Текущая цепочка, которую нужно разделить

Сейчас `frmPersonnelActionWizardV2.btnImportResponse_Click` вызывает `SaveAction`. `SaveAction` вызывает `WriteValues`, а тот записывает видимые поля в `PersonnelEventInput` через `mdlPersonnelEvents.SetPersonnelWizardValue`; затем вызывается `SavePersonnelWizardAction` → `SavePersonnelEventInput` → `SavePersonnelEvent`. `SavePersonnelEvent` создаёт BEFORE/AFTER snapshots, событие, назначения, документы и изменяет текущее состояние.

В P1 эта запись разрешена только из обработчика явной кнопки подтверждения. Нажатие отмены, повторное открытие предпросмотра, ошибка расчёта и закрытие формы без подтверждения не должны вызывать эту цепочку.

## Контракт draft и источники

Transient draft собирается из текущих значений статических design-time контролов через `TextOf`, но до подтверждения не записывается в `PersonnelEventInput`.

| Группа | Поля TRANSFER | Поля EXCLUSION | Источник before | Источник after/проектируемого значения |
|---|---|---|---|---|
| Идентификация | `employee_id` | `employee_id` | `mdlPersonnelEvents.GetCurrentPersonnelState` и Employees | draft, EmployeeID не изменяется |
| Даты и реквизиты | `event_date`, `effective_date`, `order_reference`, `basis_text`, `comment` | те же | текущая карточка/пустое значение | draft |
| Кадровое состояние | `new_rank`, `new_position`, `new_section`, `new_military_unit`, `new_vus` | — | current state | draft; пустое поле сохраняет подтверждённое значение по действующему правилу |
| Даты перемещения | `handover_date`, `acceptance_date`, `duty_start_date` | `handover_date` | текущее значение при наличии | draft или fallback по существующей логике |
| Место/назначение | `destination_unit`, `destination_location` | `destination_unit`, `destination_location` | текущее значение | draft |
| Финансово-отпускные сведения | — | `material_assistance_status`, `main_leave_status`, `additional_leave_status` | текущее значение | draft |
| Служебное состояние | `status` | `status` | не редактируется оператором | UI-only; не включается в diff |

Поля поиска (`search`, `search_results`) и `saved_event_id` не являются изменением события и не попадают в preview diff. Реальный набор видимых полей берётся из `VisibleFieldKeys` в форме; при расхождении со спецификацией сначала обновляется эта спецификация.

## Модель preview

Builder (предлагаемое имя `BuildPersonnelActionPreview`) возвращает объект/Dictionary с полями:

- `event_type`, `employee_id`, `is_valid`, `can_confirm`;
- `changed_fields`: ordered collection объектов `key`, `label`, `before`, `after`, `change_kind` (`CHANGED`, `ADDED`, `REMOVED`, `UNCHANGED`);
- `payment_changes`: ordered collection `payment_code`, `change_kind` (`START`, `CONTINUE`, `STOP`, `UNCHANGED`, `REQUIRES_DECISION`), `amount_kind`, `amount_value`, `cap_group`, `act_id`, `explanation`;
- `warnings`: ordered collection `severity` (`INFO`, `WARN`, `ERROR`), `code`, `label`, `detail` без персональных значений;
- `order_projection`: нормализованные существенные значения, которые будут использованы финальным Word renderer;
- `counts`: число изменённых полей, стартующих/останавливаемых выплат и предупреждений.

Порядок элементов стабилен: сначала кадровые поля в порядке формы, затем даты/реквизиты, затем выплаты и предупреждения. Сравнение строк нормализует пробелы и регистр кодовых значений; даты сравниваются по дате, а не по отображаемому формату. Пустое и отсутствующее значение считаются равными только после явного преобразования `Null/Empty` в единое представление.

## Правила diff и выплат

1. `TRANSFER`: before — состояние до события, after — draft с применением `new_*` и уже действующих fallback-правил. Показывать только отличающиеся значения, но предупреждать о пропущенном обязательном поле.
2. `EXCLUSION`: before — текущее состояние; after — проектируемое состояние после исключения. Отдельно показывать прекращение активных назначений и переход `Employees.IsActive` в `NO`; это прогноз, а не запись.
3. Выплаты рассчитываются тем же существующим `mdlPersonnelAllowanceRules` и входными условиями. Preview не меняет `PaymentAssignments` и не активирует неподтверждённые правила.
4. Платёж со статусом `REQUIRES_DECISION` отображается как предупреждение и не считается активируемым.
5. Для `MOBILIZED_FIXED_158000` сохраняются текущие подтверждённые `ActID`/порядок; P1 не меняет сумму, основание или пределы.
6. Если изменение состояния невозможно однозначно вычислить, `can_confirm = False`, выводится код ошибки/предупреждения, а сохранение не вызывается.

## Предупреждения и блокирующие условия

Блокирующие (`ERROR`): нет EmployeeID/current state, невалидная дата, неизвестный тип действия, нарушение существующей валидации кадрового события, отсутствует обязательное поле для этого типа.

Неблокирующие (`WARN`): остановлены активные выплаты, отсутствует optional документ, значение требует решения по НПА, after совпадает с before несмотря на выбранный тип действия. Информационные (`INFO`): поле не изменилось, выплата продолжается, экспорт станет доступен после подтверждения.

В журнал нельзя писать ФИО, личный/табельный/внутренний номер, текст основания, реквизиты документа или полный путь к DOCX. Разрешены тип события, коды полей/выплат, количества и коды ошибок.

## Word-согласованность

P1 preview не создаёт Word и не вызывает `mdlPersonnelEventOrderExport.ExportPersonnelEventOrder`, потому что этот метод требует сохранённый EventID, регистрирует документ и переводит событие в `EXPORTED`.

Существенные строки для preview и финального Word должны собираться общим чистым renderer/DTO, вынесенным из `WriteTransferOrder` и `WriteExclusionOrder`. Финальный экспорт после подтверждения получает persisted EventID и выполняет текущую регистрацию. Preview обязан показывать или явно помечать все существенные данные, проверяемые в DOCX: вид/состав изменения, даты, реквизиты приказа, место/подразделение, основание, оклады/финансово-отпускные строки и применимые выплаты.

## UI-состояния

```text
EDITING → PREVIEW_READY → CONFIRMED → SAVED → EXPORTABLE
    │           │
    └───────────┴→ CANCELLED/EDITING
```

- `EDITING`: можно менять draft; экспорт запрещён.
- `PREVIEW_READY`: controls preview read-only; подтверждение и отмена доступны.
- `CONFIRMED`: transient draft ещё не считается сохранённым до успешного SaveAction; повторное подтверждение идемпотентно и блокируется.
- `SAVED`: записан ровно один EventID; экспорт разрешён.
- `CANCELLED`: transient preview удалён; реестры и `saved_event_id` не изменены.

Если владелец вручную добавляет controls, их `(Name)`, тип и контейнер фиксируются экспортом. VBA не должен использовать `Controls.Add` и runtime `Left/Top/Width/Height`; `.frx` не редактировать вручную.

## Приёмочная матрица P1

| Сценарий | Ожидаемый результат |
|---|---|
| TRANSFER с изменением должности и дат | preview показывает before/after и будущие даты |
| EXCLUSION с активными выплатами | preview показывает STOP и IsActive→NO как прогноз |
| Cancel из preview | нет новых строк Events/Snapshots/Assignments/Documents, input не сохранён |
| Confirm один раз | ровно одно событие и два snapshots; UI получает EventID |
| Confirm дважды | второе нажатие ничего не добавляет |
| Изменить draft после preview | старый preview сброшен; требуется новый preview |
| Ошибка обязательного поля | `can_confirm=False`, запись и Word недоступны |
| Попытка Export до save | Word не создаётся и DocumentRegistry не меняется |
| ENROLLMENT | существующий поток не меняется |
| Финальный Word | существенные поля совпадают с подтверждённым preview |

## Файлы и проверки для следующего шага

- Код: `CreateOrder.xlsm.modules/frmPersonnelActionWizardV2.frm`, `mdlPersonnelEvents.bas`, `mdlPersonnelEventOrderExport.bas` или новый чистый модуль.
- Форма/ресурсы: только design-time экспорт `frmPersonnelActionWizardV2.frm/.frx/.layout.csv`; `.frx` не редактировать текстом.
- Локализация: `ModuleLocalization.bas`, лист `Localization`.
- Тесты: расширить `tools/Test-PersonnelActionWizardV2Safe.ps1`; при необходимости добавить `tools/Test-PersonnelActionPreviewSafe.ps1`.
- Обязательные команды после реализации: designer-test, P1 address-test, `tools/Test-PersonnelActionWizardSafe.ps1`, `Test-PersonnelEvents.ps1`, `Test-PaymentsEnrollmentAcceptance.ps1`, затем read-only ручное открытие preview.
- Перед импортом в закрытую книгу: проверить `EXCEL.EXE`, создать backup, импортировать только после зелёных тестов, повторить P1 address-test.

## Definition of Done

P1.1 считается закрытым этой спецификацией. P1 полностью готов только когда cancel/confirm/duplicate-confirm и Word-согласованность проходят матрицу, существующая индивидуальная регрессия зелёная, документация обновлена, а импорт в рабочую книгу подтверждён post-install тестом. P2 не начинается автоматически.

## Вопросы владельцу перед UI-реализацией

1. Размещать preview на отдельной странице `MultiPage` или в статичном блоке текущей страницы? **Решение: отдельная страница `MultiPage` (`pgPreview`).**
2. Нужен ли оператору полный текст будущего Word или достаточно существенных строк с кнопкой «открыть полный preview»? **Решение: только существенные строки.**
3. Должно ли изменение after, совпадающее с before, блокировать подтверждение или только показываться предупреждением? **Решение: только предупреждение, подтверждение не блокировать.**
