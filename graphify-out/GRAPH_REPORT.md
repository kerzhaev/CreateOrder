# Graph Report - CreateOrder  (2026-07-14)

## Corpus Check
- 53 files · ~56,747 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 667 nodes · 685 edges · 64 communities (54 shown, 10 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `4c434330`
- Run `git rev-parse HEAD` and compare to check if the graph is stale.
- Run `graphify update .` after code changes (no API cost).

## Community Hubs (Navigation)
- Test-PaymentsEnrollmentAcceptance.ps1
- Block 2. Enrollment workflow: new tab and rules engine
- Руководство пользователя: Формирователь приказов (CreateOrder)
- Block 1. Payments Without Periods: list-based assignment mode
- Project Context & Knowledge Base
- ensure_localization_sheet.py
- Feature Specification: Payments Without Periods and Enrollment Workflow
- Test-ZP12Validation.ps1
- Technical Notes
- Спецификация: Гибридная система лицензирования (Time-Bomb + File-Based Activation)
- plan.md
- release-protection.md
- План реализации: кадровые события и движок надбавок
- Спецификация: правила назначения надбавок при зачислении
- 11. Этапы реализации
- 4. Справочники правил
- 7. Пользовательские сценарии
- 8. Word-архитектура
- ������������: ������������� ������� ������ � ���������� ��� ��������
- Tasks: [FEATURE NAME]
- ����������� ����: ������������� ������� ������ � ���������� ��� ��������
- Технический план: Массовое добавление сотрудников на лист "Выплаты_Без_Периодов"
- 2. ��������� ���� ����������
- update-agent-context.ps1
- ���������� �� ��������� ������� �������� ��� ��������
- 3. Детальный план изменений
- User Scenarios & Testing *(mandatory)*
- Core Principles
- common.ps1
- Implementation Plan: [FEATURE]
- 2. Структура изменений
- [PROJECT NAME] Development Guidelines
- 3. Целевая модель данных
- 2. Требования
- create-new-feature.ps1
- 2. Требования
- 6. Единый движок оценки выплат
- [CHECKLIST TYPE] Checklist: [FEATURE NAME]
- Спецификация: Импорт утвержденных рапортов (Word) в ДСО
- 12. Тестовая стратегия
- 5. Сервис связывания со Штат
- D89 Validation
- 9. Интерфейс
- tasks.md
- alu-frp-export-plan.md
- plan.md
- spec.md
- План: приказы о перемещении и исключении из списков
- Передача проекта — 14.07.2026
- Статус проекта CreateOrder
- Разбор источника: приказ МО РФ №430дсп — особые достижения
- Спецификация: единый мастер кадровых действий
- Правила работы с проектом CreateOrder
- Test-PersonnelEvents.ps1
- Test-PersonnelActionWizardSafe.ps1
- Test-PersonnelRibbonSafe.ps1
- plan.md

## God Nodes (most connected - your core abstractions)
1. `План реализации: кадровые события и движок надбавок` - 23 edges
2. `Спецификация: правила назначения надбавок при зачислении` - 17 edges
3. `Tasks: [FEATURE NAME]` - 13 edges
4. `11. Этапы реализации` - 12 edges
5. `Block 1. Payments Without Periods: list-based assignment mode` - 12 edges
6. `Block 2. Enrollment workflow: new tab and rules engine` - 12 edges
7. `Технический план: Массовое добавление сотрудников на лист "Выплаты_Без_Периодов"` - 10 edges
8. `Спецификация: Массовое добавление сотрудников на лист "Выплаты_Без_Периодов"` - 10 edges
9. `������������: ������������� ������� ������ � ���������� ��� ��������` - 10 edges
10. `План: приказы о перемещении и исключении из списков` - 10 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Import Cycles
- None detected.

## Communities (64 total, 10 thin omitted)

### Community 0 - "Test-PaymentsEnrollmentAcceptance.ps1"
Cohesion: 0.09
Nodes (18): Assert-CustomUiUsesLocalizationCallbacks(), Assert-EnrollmentLocalizationKeysSeeded(), Assert-LocalizationKeysSeededSafely(), Assert-PaymentModulesUseLocalization(), Assert-RibbonHandlersUseLocalization(), Assert-True(), Get-BackendFieldRow(), Get-BackendValue() (+10 more)

### Community 1 - "Block 2. Enrollment workflow: new tab and rules engine"
Cohesion: 0.09
Nodes (23): 1. Employee, 2. Enrollment event, 3. Allowance assignment, Acceptance criteria, Block 2. Enrollment workflow: new tab and rules engine, Business event model, Integration requirements, Main entities (+15 more)

### Community 2 - "Руководство пользователя: Формирователь приказов (CreateOrder)"
Cohesion: 0.07
Nodes (28): 1. Введение и подготовка к работе <a name="введение"></a>, 2.1. Группа «Приказы и Рапорты» <a name="группа-приказы"></a>, 2.2. Группа «Надбавки» <a name="группа-надбавки"></a>, 2.3. Группа «Импорт и Отчеты» <a name="группа-импорт"></a>, 2.4. Группа «Проверка и Настройки» <a name="группа-проверка"></a>, 2. Панель управления (Лента «СВО Макросы») <a name="панель-управления"></a>, 3.1. Лист «ДСО» (Периоды) <a name="лист-дсо"></a>, 3.2. Лист «Надбавки без периодов» <a name="лист-надбавки"></a> (+20 more)

### Community 3 - "Block 1. Payments Without Periods: list-based assignment mode"
Cohesion: 0.08
Nodes (24): Acceptance criteria, Block 1. Payments Without Periods: list-based assignment mode, Current implementation scope, Export requirements, Feature Specification: Payments Without Periods and Enrollment Workflow, Functional concept, Implementation priority, Mandatory implementation workflow (+16 more)

### Community 4 - "Project Context & Knowledge Base"
Cohesion: 0.14
Nodes (13): 0. System Instructions (Agent Rules), 1. Глобальная цель, 2. Текущее состояние (Active State), 3.1. AI Factory / Pipeline Notes, 3. Архитектурные решения (ADR), 4.1. Основные модули экспорта документов, 4.2. Модули работы с данными, 4.3. Вспомогательные и Системные модули (+5 more)

### Community 5 - "ensure_localization_sheet.py"
Cohesion: 0.38
Nodes (10): ArgumentParser, Path, build_parser(), get_excel_open_path(), get_or_create_sheet(), main(), parse_translations(), read_module_text() (+2 more)

### Community 6 - "Feature Specification: Payments Without Periods and Enrollment Workflow"
Cohesion: 0.06
Nodes (32): 1.1. Массовое добавление по номерам, 1.2. Добавление одного сотрудника через форму поиска, 1. Описание задачи (пользовательский сценарий), 2.1. Массовое добавление, 2.2. Добавление через форму, 2. Функциональные требования, 3.1. Структура листа "Выплаты_Без_Периодов", 3.2. Структура листа "Штат" (+24 more)

### Community 7 - "Test-ZP12Validation.ps1"
Cohesion: 0.24
Nodes (3): Apply-RegressionFixes(), Get-ZP12State(), Invoke-Python()

### Community 8 - "Technical Notes"
Cohesion: 0.33
Nodes (5): CodeModule import normalization, Mandatory Excel safety workflow, Technical Notes, VBA localization encoding rule, VBA module import/export rule

### Community 9 - "Спецификация: Гибридная система лицензирования (Time-Bomb + File-Based Activation)"
Cohesion: 0.50
Nodes (3): Описание, Спецификация: Гибридная система лицензирования (Time-Bomb + File-Based Activation), Требования

### Community 13 - "План реализации: кадровые события и движок надбавок"
Cohesion: 0.13
Nodes (14): 10. Миграция текущей реализации, 13. Приёмочные критерии, 14. Блокирующие условия финальной приёмки, 15. Порядок работы по каждой надбавке, 16. Контроль отклонений от подтверждённых норм, 17. Подготовка корректировки из истории, 18. Сохранность условий назначения выплат, 19. Безопасный сброс кадровой формы (+6 more)

### Community 14 - "Спецификация: правила назначения надбавок при зачислении"
Cohesion: 0.05
Nodes (36): 10. Разделение справочника сотрудников и журнала событий, 11. Сценарий исключения из списков организации, 12.1. Связь существующего Word-потока зачисления с кадровым событием, 12. Общая архитектура кадровых Word-документов, 13.1. Категория военнослужащих по мобилизации, 13.2. Формулировки Word-приказа для подтверждённых НПА, 13.3. Защитные ограничения подтверждённого набора, 13. Предварительное правило объединения особых достижений (+28 more)

### Community 15 - "11. Этапы реализации"
Cohesion: 0.17
Nodes (12): 11. Этапы реализации, Этап 0. Уточнение и фиксация предметной области, Этап 10. Стабилизация и миграция, Этап 1. Инфраструктура данных, Этап 2. Справочник сотрудников, Этап 3. Каркас справочников выплат, Этап 4. Движок оценки, Этап 5. ФИЗО как первое полностью согласованное правило (+4 more)

### Community 16 - "4. Справочники правил"
Cohesion: 0.22
Nodes (9): 4.1. Нормативные акты, 4.2. Виды выплат, 4.3. Правила назначения, 4.4. Группы ограничений, 4.5. Классификатор должностей, 4.6. ФИЗО, 4.7. Медали, 4.8. Профили организации (+1 more)

### Community 17 - "7. Пользовательские сценарии"
Cohesion: 0.33
Nodes (6): 7.1. Зачисление из Штат, 7.2. Ручное зачисление, 7.3. Перемещение, 7.4. Исключение, 7.5. Просмотр истории, 7. Пользовательские сценарии

### Community 18 - "8. Word-архитектура"
Cohesion: 0.40
Nodes (5): 8.1. Общие блоки, 8.2. Зачисление и перемещение, 8.3. Исключение, 8.4. Версионирование, 8. Word-архитектура

### Community 19 - "������������: ������������� ������� ������ � ���������� ��� ��������"
Cohesion: 0.06
Nodes (30): ������������: ������������� ������� ������ � ���������� ��� ��������, 1. �������� ������ (������� ������������), 2. ������� ��������� �������, 2.1. ������������ ������� � ���������, 2.2. ��������� ������ ���������������, 3. ���������� � ����� ����� ������, 3.1. �������� ��������� ��������� ���, 3.2. �������� �� ������ (+22 more)

### Community 20 - "Tasks: [FEATURE NAME]"
Cohesion: 0.07
Nodes (26): Dependencies & Execution Order, Format: `[ID] [P?] [Story] Description`, Implementation for User Story 1, Implementation for User Story 2, Implementation for User Story 3, Implementation Strategy, Incremental Delivery, MVP First (User Story 1 Only) (+18 more)

### Community 21 - "����������� ����: ������������� ������� ������ � ���������� ��� ��������"
Cohesion: 0.09
Nodes (22): ����������� ����: ������������� ������� ������ � ���������� ��� ��������, 1. ����� ���������, 1.1. ����� ������ VBA, ��� 1: �������� ������� (��� �����������), 1.2. ���������� ������ VBA, 1.3. ����� ����� Excel, 1.4. ���������� ���������������� �����, ��� 2: ���������� ������ �� ������������� (+14 more)

### Community 22 - "Технический план: Массовое добавление сотрудников на лист "Выплаты_Без_Периодов""
Cohesion: 0.09
Nodes (21): 1. Обзор изменений, 2.1. Модули VBA (расширение существующих), 2.2. Новые файлы, 2.3. Обновление Ribbon XML, 2. Затронутые файлы, 4.1. Алгоритм поиска по табельному номеру, 4.2. Алгоритм массового импорта, 4.3. Алгоритм формы выбора сотрудника (+13 more)

### Community 23 - "2. ��������� ���� ����������"
Cohesion: 0.11
Nodes (19): 1.1. ���� "�������_���_��������", 1.2. ���� "����������_���_������", 1.3. ���� "����������_����_������", ���� 1: �������� ��������� ������ (����� Excel), 2. ��������� ���� ����������, 2.1. ���� ������, 2.2. �������� �������, ���� 2: ������ ������ �� ������������� (`mdlReferenceData.bas`) (+11 more)

### Community 24 - "update-agent-context.ps1"
Cohesion: 0.28
Nodes (18): Extract-PlanField(), Format-TechnologyStack(), Get-CommandsForLanguage(), Get-LanguageConventions(), Get-ProjectStructure(), Main(), New-AgentFile(), Parse-PlanData() (+10 more)

### Community 25 - "���������� �� ��������� ������� �������� ��� ��������"
Cohesion: 0.11
Nodes (17): ���������� �� ��������� ������� �������� ��� ��������, 1.1. ���� "�������_���_��������", 1.2. ���� "����������_���_������", 1.3. ���� "����������_����_������", 1. �������� ������ Excel, 2. ���������� ���������������� ����� (XML), 3.1. ������� ��� ����� ������, 3.2. ������ ������������� ������ (+9 more)

### Community 26 - "3. Детальный план изменений"
Cohesion: 0.12
Nodes (17): 1.1. Добавление функции поиска колонки "Лицо", 1.2. Добавление функции поиска по табельному номеру, 1.3. Добавление универсальной функции поиска, 2.1. Добавление функции массового импорта, 2.2. Добавление функции обработки диапазона, 3.1. Структура формы, 3.2. Публичные переменные, 3.3. События формы (+9 more)

### Community 27 - "User Scenarios & Testing *(mandatory)*"
Cohesion: 0.17
Nodes (11): Edge Cases, Feature Specification: [FEATURE NAME], Functional Requirements, Key Entities *(include if feature involves data)*, Measurable Outcomes, Requirements *(mandatory)*, Success Criteria *(mandatory)*, User Scenarios & Testing *(mandatory)* (+3 more)

### Community 28 - "Core Principles"
Cohesion: 0.18
Nodes (10): Core Principles, Governance, [PRINCIPLE_1_NAME], [PRINCIPLE_2_NAME], [PRINCIPLE_3_NAME], [PRINCIPLE_4_NAME], [PRINCIPLE_5_NAME], [PROJECT_NAME] Constitution (+2 more)

### Community 29 - "common.ps1"
Cohesion: 0.36
Nodes (5): Get-CurrentBranch(), Get-FeatureDir(), Get-FeaturePathsEnv(), Get-RepoRoot(), Test-HasGit()

### Community 30 - "Implementation Plan: [FEATURE]"
Cohesion: 0.22
Nodes (8): Complexity Tracking, Constitution Check, Documentation (this feature), Implementation Plan: [FEATURE], Project Structure, Source Code (repository root), Summary, Technical Context

### Community 31 - "2. Структура изменений"
Cohesion: 0.25
Nodes (7): 1. Архитектура решения, 2. Структура изменений, План реализации: Импорт рапортов из Word, Фаза 1: UI и Подготовка, Фаза 2: Механизм экстракции (Extract), Фаза 3: Механизм трансформации (Transform), Фаза 4: Механизм загрузки (Load)

### Community 32 - "[PROJECT NAME] Development Guidelines"
Cohesion: 0.25
Nodes (7): Active Technologies, Code Style, Commands, [PROJECT NAME] Development Guidelines, Project Structure, Recent Changes, Workflow Rules

### Community 33 - "3. Целевая модель данных"
Cohesion: 0.22
Nodes (9): 3.1. Справочник сотрудников, 3.2. Текущее состояние сотрудника, 3.3. Журнал кадровых событий, 3.4. Снимки состояния, 3.5. Назначения выплат, 3.6. Реестр документов, 3.7. Экран истории сотрудника — реализуемый этап, 3.8. Безопасное связывание с выгрузкой Штат — реализуемый этап (+1 more)

### Community 34 - "2. Требования"
Cohesion: 0.29
Nodes (6): 1. Цель, 2.1. Многоколоночный список (Multi-column ListBox), 2.2. Живой поиск (Live Filtering), 2.3. Навигация (Keyboard Navigation), 2. Требования, Спецификация: Улучшение UI поиска сотрудников

### Community 35 - "create-new-feature.ps1"
Cohesion: 0.43
Nodes (5): ConvertTo-CleanBranchName(), Get-BranchName(), Get-HighestNumberFromBranches(), Get-HighestNumberFromSpecs(), Get-NextBranchNumber()

### Community 36 - "2. Требования"
Cohesion: 0.33
Nodes (5): 1. Описание задачи, 2.1 Интерфейс, 2.2 Логика данных, 2. Требования, Спецификация: Обновление выгрузки Алушта/ФРП

### Community 37 - "6. Единый движок оценки выплат"
Cohesion: 0.33
Nodes (6): 6.1. Входные данные, 6.2. Результат оценки, 6.3. Применение ограничений, 6.4. Жизненный цикл выплат, 6.5. Особый профиль `MOBILIZED`, 6. Единый движок оценки выплат

### Community 38 - "[CHECKLIST TYPE] Checklist: [FEATURE NAME]"
Cohesion: 0.40
Nodes (4): [Category 1], [Category 2], [CHECKLIST TYPE] Checklist: [FEATURE NAME], Notes

### Community 39 - "Спецификация: Импорт утвержденных рапортов (Word) в ДСО"
Cohesion: 0.50
Nodes (3): 1. Цель (User Stories), 2. Требования (Requirements), Спецификация: Импорт утвержденных рапортов (Word) в ДСО

### Community 40 - "12. Тестовая стратегия"
Cohesion: 0.50
Nodes (4): 12.1. Уровни тестов, 12.2. Обязательные сценарии, 12.3. Регрессия, 12. Тестовая стратегия

### Community 41 - "5. Сервис связывания со Штат"
Cohesion: 0.50
Nodes (4): 5.1. Импорт и поиск, 5.2. Сопоставление, 5.3. Разрешение расхождений, 5. Сервис связывания со Штат

### Community 42 - "D89 Validation"
Cohesion: 0.50
Nodes (3): D89 Validation, Goal, Scope

### Community 43 - "9. Интерфейс"
Cohesion: 0.67
Nodes (3): 9.1. Разделы, 9.2. Требования, 9. Интерфейс

### Community 53 - "План: приказы о перемещении и исключении из списков"
Cohesion: 0.10
Nodes (20): 1. Структура образцов, 2. Модель Word-экспорта, 3. Сценарий `EXCLUSION`, 4. Сценарий `TRANSFER`, 5. Изменения в приложении, 6. Валидация, 7. Этапы реализации, 8. Приёмочные сценарии (+12 more)

### Community 54 - "Передача проекта — 14.07.2026"
Cohesion: 0.12
Nodes (16): Автоматические проверки, Где искать код и документацию, Единый вход в кадровые действия, Надбавки и нормативная логика, Ошибки VBA, встреченные при предыдущем импорте, Ошибки и ограничения Excel/интерфейса, Ошибки и принятые решения, Передача проекта — 14.07.2026 (+8 more)

### Community 55 - "Статус проекта CreateOrder"
Cohesion: 0.18
Nodes (10): Аудит готовности на 13.07.2026, Документы и ориентиры, Единый вход в кадровые действия, Зафиксированные решения, Как продолжить работу, Направления работы, Общая картина, Открытые вопросы (+2 more)

### Community 56 - "Разбор источника: приказ МО РФ №430дсп — особые достижения"
Cohesion: 0.25
Nodes (7): Влияние на текущую спецификацию, Ограничение источника, Правила, явно видимые в переданном фрагменте, Разбор источника: приказ МО РФ №430дсп — особые достижения, Решения, которые ожидаются от владельца проекта, Связанное локальное правило из другого НПА: военнослужащие по мобилизации, Условия расчёта, которые нужно реализовать отдельно от самих процентов

### Community 57 - "Спецификация: единый мастер кадровых действий"
Cohesion: 0.29
Nodes (6): Вторая поставка: кадровый мастер, Принципы, Приёмка, Состав первой поставки, Спецификация: единый мастер кадровых действий, Цель

### Community 58 - "Правила работы с проектом CreateOrder"
Cohesion: 0.40
Nodes (4): В начале каждой сессии, Документация — обязательная часть работы, Правила работы с проектом CreateOrder, Формат статуса

## Knowledge Gaps
- **404 isolated node(s):** `0. System Instructions (Agent Rules)`, `1. Глобальная цель`, `2. Текущее состояние (Active State)`, `3. Архитектурные решения (ADR)`, `3.1. AI Factory / Pipeline Notes` (+399 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **10 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `План реализации: кадровые события и движок надбавок` connect `План реализации: кадровые события и движок надбавок` to `3. Целевая модель данных`, `6. Единый движок оценки выплат`, `12. Тестовая стратегия`, `5. Сервис связывания со Штат`, `9. Интерфейс`, `11. Этапы реализации`, `4. Справочники правил`, `7. Пользовательские сценарии`, `8. Word-архитектура`?**
  _High betweenness centrality (0.011) - this node is a cross-community bridge._
- **Why does `Block 2. Enrollment workflow: new tab and rules engine` connect `Block 2. Enrollment workflow: new tab and rules engine` to `Block 1. Payments Without Periods: list-based assignment mode`?**
  _High betweenness centrality (0.003) - this node is a cross-community bridge._
- **Why does `Feature Specification: Payments Without Periods and Enrollment Workflow` connect `Block 1. Payments Without Periods: list-based assignment mode` to `Block 2. Enrollment workflow: new tab and rules engine`?**
  _High betweenness centrality (0.003) - this node is a cross-community bridge._
- **What connects `0. System Instructions (Agent Rules)`, `1. Глобальная цель`, `2. Текущее состояние (Active State)` to the rest of the system?**
  _404 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Test-PaymentsEnrollmentAcceptance.ps1` be split into smaller, more focused modules?**
  _Cohesion score 0.09090909090909091 - nodes in this community are weakly interconnected._
- **Should `Block 2. Enrollment workflow: new tab and rules engine` be split into smaller, more focused modules?**
  _Cohesion score 0.08695652173913043 - nodes in this community are weakly interconnected._
- **Should `Руководство пользователя: Формирователь приказов (CreateOrder)` be split into smaller, more focused modules?**
  _Cohesion score 0.06896551724137931 - nodes in this community are weakly interconnected._