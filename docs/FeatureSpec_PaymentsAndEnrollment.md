# Feature Specification: Payments Without Periods and Enrollment Workflow

## Status

Accepted as the next implementation goal for `CreateOrder`.

## Current implementation scope

This goal is being implemented in two immediate waves:

1. `Wave A. Grouped non-periodic payments`
- extend `Выплаты_Без_Периодов` from 6 columns to 14 columns;
- add package metadata, per-row parameter, shared basis, export flag, status, source enrollment id;
- add package creation and recalculation buttons on the sheet;
- add automatic amount resolution for class qualification;
- group export by `payment type + package id`;
- add duplicate detection inside one package.

2. `Wave B. Enrollment tab`
- create sheet `Зачисление`;
- add validation of required personnel and date fields;
- add button to generate payment rows into `Выплаты_Без_Периодов`;
- create one generated package per enrollment row;
- prefill class qualification, ФИЗО, secrecy, and listed standard payments.

3. `Wave C. Enrollment master form and payment split`
- create sheet-form `Мастер_Зачисления` for operator input instead of raw entry directly into the journal;
- split enrollment proposals into two visible buckets:
  - position-based allowances;
  - personal allowances;
- allow saving one form record into the `Зачисление` journal;
- allow saving one form record and immediately generating rows in `Выплаты_Без_Периодов`.

## Technical acceptance for the current wave

The current wave is accepted only if all items below are true:

1. Workbook self-healing recreates:
- `ДСО`;
- `Выплаты_Без_Периодов` with the expanded column model;
- `Зачисление`.

2. Ribbon command `Проверить лист` works on:
- `ДСО`;
- `Выплаты_Без_Периодов`;
- `Зачисление`.

3. Export of non-periodic payments creates separate files for different packages of the same payment type.

4. `Классная квалификация` can be generated from:
- direct entry on `Выплаты_Без_Периодов`;
- generated rows from `Зачисление`.

5. Before every workbook import run:
- hanging `Excel` processes are terminated;
- backup copy of `CreateOrder.xlsm` is created.

6. Enrollment operator workflow must be understandable without sheet-overlaid buttons:
- all main actions are available from ribbon group `Зачисление`;
- raw journal sheet remains available, but the preferred entry point is `Мастер_Зачисления`.

## Mandatory implementation workflow

Before any code import into `CreateOrder.xlsm`:

1. Stop hanging `Excel` processes.
2. Create a backup copy of the workbook.
3. Only then import modified VBA and run validation.

This rule is mandatory for every implementation pass on these features.

---

## Block 1. Payments Without Periods: list-based assignment mode

### Problem

The current sheet `Выплаты_Без_Периодов` works for simple manual entries but is weak for common real workflows where one allowance is assigned to a list of people in one paragraph of one draft order.

Typical examples:

- class qualification allowance;
- physical training allowance;
- secrecy allowance;
- other non-periodic allowances with one shared basis and per-person parameters.

The current model forces too much manual work:

- manual amount entry for many people;
- repeated filling of the same basis;
- weak support for one group paragraph;
- weak distinction between shared allowance logic and person-specific parameters.

### Target result

Add a controlled list-based mode for `Выплаты_Без_Периодов` where the operator creates one allowance package and fills a group of people inside it.

### Functional concept

The feature must support two modes:

1. `Individual mode`
For one person and one allowance record.

2. `List mode`
For one shared allowance package:
- one allowance type;
- one shared basis or shared basis block;
- one export paragraph or one grouped order fragment;
- multiple employees;
- one per-person parameter set.

### Required sheet-level model

The data model for non-periodic payments must distinguish:

1. `Package-level fields`
- package id;
- package type;
- allowance type;
- shared basis text;
- shared document number;
- shared document date;
- export group flag;
- package comment.

2. `Person-level fields`
- employee identifier;
- FIO;
- personal number;
- rank;
- position;
- organization / staff section;
- allowance parameter;
- calculated amount or percent;
- individual basis override, if needed;
- validation status.

### Required allowance logic

The system must support at least these allowance logic categories:

1. `Fixed by manual amount`
Operator enters amount or percent directly.

2. `Calculated from parameter`
Example: class qualification.
Operator selects class, program determines percent/amount from reference data.

3. `Conditional by eligibility`
Example: secrecy allowance.
Payment is available only if required condition is present, such as clearance/admission or relevant role.

4. `Mixed`
Shared package with optional per-person override.

### User workflow

#### Workflow A: create a grouped package

1. Operator selects `allowance type`.
2. Operator chooses `list mode`.
3. Operator enters shared basis fields once.
4. Operator adds employees by:
   - search form;
   - mass import of numbers;
   - copy from selection.
5. Program auto-fills employee data from `Штат`.
6. Operator sets per-person parameter where needed:
   - class;
   - physical qualification;
   - secrecy/admission flag;
   - custom amount if allowed.
7. Program calculates resulting amount/percent.
8. Program validates all rows.
9. Program exports one grouped paragraph or one grouped draft order.

#### Workflow B: individual payment

1. Operator selects one employee.
2. Program fills reference data from `Штат`.
3. Operator fills basis and allowance parameter.
4. Program validates and exports the single record.

### UI requirements

Implementation may use the existing sheet plus forms, but the following capabilities are required:

1. Quick creation of a payment package.
2. Clear visual separation between shared package fields and per-person rows.
3. Search-based employee add.
4. Mass employee add by numbers.
5. Bulk fill of shared values across selected rows.
6. Recalculation button for computed payments.
7. Preview of resulting export text.
8. Validation indicator per row.

### Reference data requirements

The feature must use or extend reference sheets for:

- allowance types;
- parameter-to-percent rules;
- export template mapping;
- required fields by allowance type;
- conditions for eligibility.

For class qualification specifically, the system must support:

- class name or code;
- reference percent or amount by class;
- automatic fill of resulting payment size.

### Validation rules

Minimum validation set:

1. Employee must exist in `Штат`.
2. Allowance type must exist in the reference.
3. Shared basis must be filled for grouped export.
4. Required per-person parameter must be filled.
5. Duplicate employee in the same package must be highlighted.
6. Suspicious zero or empty calculated amount must be highlighted.
7. Employee belonging to another organization or staff section must be highlighted.
8. Missing eligibility condition must block export or mark row as warning, depending on rule settings.

### Export requirements

The export layer must support:

1. grouped paragraph generation;
2. individual paragraph generation;
3. one package -> one paragraph block in Word;
4. correct use of FIO/rank/position declension from the updated helper engine;
5. optional grouping by allowance subtype if required later.

### Acceptance criteria

The block is accepted when:

1. Operator can create one allowance package for at least 20 people.
2. Shared basis is entered once and reused for all package rows.
3. Class qualification can be selected per person and converted automatically into result size.
4. Export produces one grouped order block without manual editing of each line.
5. Duplicate and incomplete rows are visibly highlighted.
6. Employee data is auto-filled from `Штат`.

---

## Block 2. Enrollment workflow: new tab and rules engine

### Problem

There is no dedicated process for handling a newly arrived employee and automatically determining:

- personnel data;
- key кадровые dates;
- supporting documents;
- which allowances are mandatory;
- which allowances are optional or named;
- from what exact date each allowance should start.

This leads to manual reasoning across several documents and repeated entry into multiple sheets.

### Target result

Create a new tab and workflow for `Зачисление` that turns one personnel event into structured data and a set of ready-to-apply payment assignments.

### Business event model

One enrollment record represents the event:

`employee arrived -> accepted duty -> enrolled into organization -> eligible allowances determined`

### Main entities

#### 1. Employee

Persistent identity fields:

- FIO;
- personal number;
- rank;
- organization;
- current position;
- staff section.

#### 2. Enrollment event

Document and date fields:

- order on personnel;
- report with registration;
- assignment/prescription document;
- date of arrival;
- date of duty acceptance;
- date of enrollment into unit;
- effective personnel date if separate.

#### 3. Allowance assignment

Derived or manual payment decision:

- allowance type;
- start date;
- end date if needed;
- basis document;
- parameter value;
- calculated size;
- assignment status;
- reason why not assigned automatically, if blocked.

### Required tab structure

The new tab must contain at least these sections.

#### Section A. Identity and position

- FIO;
- personal number;
- rank;
- position;
- organization / staff section;
- unit data.

#### Section B. Key personnel dates

- personnel order date;
- arrival date;
- duty acceptance date;
- enrollment date;
- date from which standard allowances begin.

#### Section C. Documents and grounds

- order number and date;
- report number and date;
- prescription / assignment document;
- optional notes.

#### Section D. Automatic allowances

Allowances that can be proposed from role/rules automatically.

Examples:

- standard payments tied to position;
- payments always assigned after enrollment if role requires them.

#### Section E. Named / conditional allowances

Allowances that require explicit employee-specific evidence.

Examples:

- class qualification;
- physical training category;
- secrecy allowance if clearance exists;
- other named allowances.

#### Section F. Result

- ready to export flag;
- missing data list;
- warnings;
- list of generated assignments;
- action buttons.

### Workflow requirements

#### Workflow A: create enrollment record

1. Operator adds or selects employee.
2. Operator fills source documents and dates.
3. Program loads staff-related data if employee is already known.
4. Program evaluates rules for automatic allowances.
5. Program shows conditional allowances requiring manual confirmation.
6. Operator confirms or fills named allowance parameters.
7. Program generates resulting assignments.
8. Program exports or transfers them to the payment workflow.

#### Workflow B: derive payments from enrollment

1. Program calculates which allowances are mandatory.
2. Program calculates start date for each allowance.
3. Program checks whether any required supporting data is missing.
4. Program marks records:
   - ready;
   - warning;
   - blocked.

### Rules engine requirements

The enrollment feature requires a rules engine or rules table with these fields:

- allowance type;
- mandatory or optional;
- depends on position;
- depends on organization;
- depends on named parameter;
- depends on additional document;
- start date source;
- export destination;
- validation severity if data is missing.

### Start date rules

The implementation must support allowance start dates based on:

- duty acceptance date;
- enrollment date;
- personnel order date;
- manually chosen date;
- rule-specific date source.

### Integration requirements

The new tab must integrate with:

1. `Штат`
For employee and position data.

2. `Выплаты_Без_Периодов`
To generate ready payment assignments or packages.

3. Word export layer
To build draft orders from enrollment decisions.

4. Validation layer
To block incomplete enrollment records from producing incorrect outputs.

### Validation rules

Minimum validation set:

1. Required identity fields must be present.
2. Required source documents must be present.
3. Start-date-driving field must be present for every generated allowance.
4. Position-dependent allowances must not be created without a resolved position.
5. Named allowance must not be generated without named parameter or document.
6. Conflicts with current `Штат` state must be highlighted.

### Acceptance criteria

The block is accepted when:

1. Operator can enter one enrollment event from the usual document set.
2. Program determines which allowances are standard and which are named.
3. Program calculates start dates according to configured rules.
4. Program can prepare output records for `Выплаты_Без_Периодов`.
5. Missing data is shown explicitly with actionable status.

---

## Implementation priority

Recommended order:

1. Block 1: grouped `Выплаты_Без_Периодов`.
2. Block 2: `Зачисление` tab and workflow.
3. Rules engine refactor to reduce duplicated logic.
4. Export polish and advanced edge cases.

## Suggested delivery phases

### Phase 1

- grouped package model for non-periodic payments;
- class qualification reference logic;
- mass add and validation;
- grouped export.

### Phase 2

- enrollment tab layout;
- enrollment data model;
- automatic allowance proposal;
- transfer to payment records.

### Phase 3

- rules table generalization;
- advanced eligibility checks;
- improved export scenarios;
- edge-case validation.
