# Technical Notes

## VBA module import/export rule

When we update VBA from exported files, we must keep the component type aligned with the target component in the workbook:

- Standard modules go in `CreateOrder.xlsm.modules/*.bas`.
- UserForms go in `CreateOrder.xlsm.modules/*.frm` plus the paired `.frx`.
- Exported `UserForm` `.frm` files must be kept with Windows `CRLF` line endings. If a form file is rewritten with plain `LF`, Excel VBA can mis-detect it during import and create a standard module instead of a real `MSForm`.
- Workbook and worksheet document modules go only in `workbook-modules/*.bas`.
- `ThisWorkbook` and sheet code modules must be imported through the document-module flow in `VbaModuleManager`; they must never be pasted into a standard module.
- Class-module headers such as `VERSION 1.0 CLASS`, `BEGIN`, `MultiUse`, and `Attribute VB_PredeclaredId` belong only in class/document modules. If that text appears inside a normal `.bas` module, VBA will throw compile errors like `Expected: end of statement`.

Before editing workbook contents or re-importing modules, check for running Excel and Word processes. Open Office COM instances can keep the workbook/template locked or cause imports and export smoke tests to land in the wrong place.

## Mandatory Excel safety workflow

Before any VBA import, workbook patching, or automated Excel validation, the following sequence is mandatory:

1. Stop all hanging `Excel` and `WINWORD` processes first.
2. Create a backup copy of the target workbook before importing modified code.
3. Only after steps 1-2 open the workbook for automation, import modules, save, and validate.

Operational rule:

- Never trust a successful import if there were stale Excel/Word COM instances or visible `Excel`/`WINWORD` processes before the run.
- If import/validation behaves inconsistently, stop `Excel` and `WINWORD` again and retry from a clean process state.
- For this project, skipping the Office cleanup step is treated as a workflow error because it can silently corrupt the resulting workbook, leave a partial VBA import, or keep Word templates/documents locked.

## CodeModule import normalization

When a standard VBA module is imported by replacing `VBComponent.CodeModule` text with `AddFromString`, the raw exported `.bas` file must be normalized first:

- remove the leading `Attribute VB_Name = "..."` line before `AddFromString`;
- keep `Option Explicit` and the procedure code;
- do not paste class/document module headers into a standard module.

Reason:

- exported `.bas` files contain metadata that is valid for file import, but not always for raw `CodeModule` text insertion;
- if `Attribute VB_Name` is inserted into module text, Excel may accept the write but the module can compile incompletely, after which public procedures appear to be "missing" or workbook behavior diverges from the source files.

## VBA localization encoding rule

Visible Russian UI text in VBA must not depend on the source file codepage:

- UserForms, MsgBox text, ribbon captions, worksheet headers, comments, and Word-export text must be routed through `ModuleLocalization.t/tf` or the local wrappers (`L`, `ET`) used by the module.
- New built-in Russian localization entries must be seeded with `AddSafe` and Unicode code points, not with direct Cyrillic string literals in `AddTranslation`.
- `AddTranslation` intentionally rejects likely mojibake values before they enter the localization cache. Do not bypass this filter.
- Acceptance tests must verify that critical localization keys are present through `AddSafe`; old corrupted `AddTranslation` entries are not considered safe.
- If a form or message shows mojibake, fix the localization key/source first. Do not paste decoded text directly into `.frm` or `.bas` files as a quick fix.
