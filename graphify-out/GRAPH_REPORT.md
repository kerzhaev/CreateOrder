# Graph Report - CreateOrder  (2026-07-13)

## Corpus Check
- 40 files · ~35,026 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 146 nodes · 164 edges · 13 communities (11 shown, 2 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Graph Freshness
- Built from commit: `e69e9275`
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

## God Nodes (most connected - your core abstractions)
1. `Block 1. Payments Without Periods: list-based assignment mode` - 12 edges
2. `Block 2. Enrollment workflow: new tab and rules engine` - 12 edges
3. `Feature Specification: Payments Without Periods and Enrollment Workflow` - 9 edges
4. `Project Context & Knowledge Base` - 8 edges
5. `Руководство пользователя: Формирователь приказов (CreateOrder)` - 8 edges
6. `Read-VbaText()` - 7 edges
7. `Required tab structure` - 7 edges
8. `main()` - 7 edges
9. `Assert-True()` - 6 edges
10. `4. Карта модулей (Module Map)` - 5 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Import Cycles
- None detected.

## Communities (13 total, 2 thin omitted)

### Community 0 - "Test-PaymentsEnrollmentAcceptance.ps1"
Cohesion: 0.09
Nodes (18): Assert-CustomUiUsesLocalizationCallbacks(), Assert-EnrollmentLocalizationKeysSeeded(), Assert-LocalizationKeysSeededSafely(), Assert-PaymentModulesUseLocalization(), Assert-RibbonHandlersUseLocalization(), Assert-True(), Get-BackendFieldRow(), Get-BackendValue() (+10 more)

### Community 1 - "Block 2. Enrollment workflow: new tab and rules engine"
Cohesion: 0.09
Nodes (23): 1. Employee, 2. Enrollment event, 3. Allowance assignment, Acceptance criteria, Block 2. Enrollment workflow: new tab and rules engine, Business event model, Integration requirements, Main entities (+15 more)

### Community 2 - "Руководство пользователя: Формирователь приказов (CreateOrder)"
Cohesion: 0.13
Nodes (14): 1. Введение и подготовка к работе <a name="введение"></a>, 2.1. Группа «Приказы и Рапорты» <a name="группа-приказы"></a>, 2.2. Группа «Надбавки» <a name="группа-надбавки"></a>, 2.3. Группа «Импорт и Отчеты» <a name="группа-импорт"></a>, 2.4. Группа «Проверка и Настройки» <a name="группа-проверка"></a>, 2. Панель управления (Лента «СВО Макросы») <a name="панель-управления"></a>, 3.1. Лист «ДСО» (Периоды) <a name="лист-дсо"></a>, 3.2. Лист «Надбавки без периодов» <a name="лист-надбавки"></a> (+6 more)

### Community 3 - "Block 1. Payments Without Periods: list-based assignment mode"
Cohesion: 0.14
Nodes (14): Acceptance criteria, Block 1. Payments Without Periods: list-based assignment mode, Export requirements, Functional concept, Problem, Reference data requirements, Required allowance logic, Required sheet-level model (+6 more)

### Community 4 - "Project Context & Knowledge Base"
Cohesion: 0.15
Nodes (12): 0. System Instructions (Agent Rules), 1. Глобальная цель, 2. Текущее состояние (Active State), 3.1. AI Factory / Pipeline Notes, 3. Архитектурные решения (ADR), 4.1. Основные модули экспорта документов, 4.2. Модули работы с данными, 4.3. Вспомогательные и Системные модули (+4 more)

### Community 5 - "ensure_localization_sheet.py"
Cohesion: 0.38
Nodes (10): ArgumentParser, Path, build_parser(), get_excel_open_path(), get_or_create_sheet(), main(), parse_translations(), read_module_text() (+2 more)

### Community 6 - "Feature Specification: Payments Without Periods and Enrollment Workflow"
Cohesion: 0.18
Nodes (10): Current implementation scope, Feature Specification: Payments Without Periods and Enrollment Workflow, Implementation priority, Mandatory implementation workflow, Phase 1, Phase 2, Phase 3, Status (+2 more)

### Community 7 - "Test-ZP12Validation.ps1"
Cohesion: 0.24
Nodes (3): Apply-RegressionFixes(), Get-ZP12State(), Invoke-Python()

### Community 8 - "Technical Notes"
Cohesion: 0.33
Nodes (5): CodeModule import normalization, Mandatory Excel safety workflow, Technical Notes, VBA localization encoding rule, VBA module import/export rule

### Community 9 - "Спецификация: Гибридная система лицензирования (Time-Bomb + File-Based Activation)"
Cohesion: 0.50
Nodes (3): Описание, Спецификация: Гибридная система лицензирования (Time-Bomb + File-Based Activation), Требования

## Knowledge Gaps
- **68 isolated node(s):** `0. System Instructions (Agent Rules)`, `1. Глобальная цель`, `2. Текущее состояние (Active State)`, `3. Архитектурные решения (ADR)`, `3.1. AI Factory / Pipeline Notes` (+63 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **2 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `Block 2. Enrollment workflow: new tab and rules engine` connect `Block 2. Enrollment workflow: new tab and rules engine` to `Feature Specification: Payments Without Periods and Enrollment Workflow`?**
  _High betweenness centrality (0.072) - this node is a cross-community bridge._
- **Why does `Feature Specification: Payments Without Periods and Enrollment Workflow` connect `Feature Specification: Payments Without Periods and Enrollment Workflow` to `Block 2. Enrollment workflow: new tab and rules engine`, `Block 1. Payments Without Periods: list-based assignment mode`?**
  _High betweenness centrality (0.070) - this node is a cross-community bridge._
- **Why does `Block 1. Payments Without Periods: list-based assignment mode` connect `Block 1. Payments Without Periods: list-based assignment mode` to `Feature Specification: Payments Without Periods and Enrollment Workflow`?**
  _High betweenness centrality (0.050) - this node is a cross-community bridge._
- **What connects `0. System Instructions (Agent Rules)`, `1. Глобальная цель`, `2. Текущее состояние (Active State)` to the rest of the system?**
  _68 weakly-connected nodes found - possible documentation gaps or missing edges._
- **Should `Test-PaymentsEnrollmentAcceptance.ps1` be split into smaller, more focused modules?**
  _Cohesion score 0.09090909090909091 - nodes in this community are weakly interconnected._
- **Should `Block 2. Enrollment workflow: new tab and rules engine` be split into smaller, more focused modules?**
  _Cohesion score 0.08695652173913043 - nodes in this community are weakly interconnected._
- **Should `Руководство пользователя: Формирователь приказов (CreateOrder)` be split into smaller, more focused modules?**
  _Cohesion score 0.13333333333333333 - nodes in this community are weakly interconnected._