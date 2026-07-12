# Release Protection Specification

- [x] Security: Расширить массив $modulesToHide в скрипте Build-Release.ps1. Добавить все модули бизнес-логики (mdlMainExport, mdlRaportExport, mdlSpravkaExport, mdlRiskExport, mdlUniversalPaymentExport, mdlFRPExport, mdlWordImport) для предотвращения прямого запуска макросов через Alt+F8 в обход ленты.
- [x] Pipeline Rule: Не использовать `DPB -> DPx` в релизном пайплайне как обязательный слой защиты. Этот хак вызывает `invalid key 'DPx'`, recovery-диалог Excel и не дает надежной защиты проекта.
- [x] Pipeline Rule: Патчить `xl/vbaProject.bin` только на уровне байтовых последовательностей фиксированной длины. Перекодировка всего `vbaProject.bin` в строку и обратно запрещена.
- [x] Pipeline Rule: `ThisWorkbook` / `ЭтаКнига` и модули листов Excel не импортируются через `VBComponents.Import` как `.bas`. Это document-модули класса. Их нужно обновлять только через код существующего компонента (`CodeModule.DeleteLines` + `AddFromString`) или через специальную ветку `importDocumentModule` в `VbaModuleManager`.
- [x] Verification: Если после автоматической сборки в `ThisWorkbook` видны строки `VERSION 1.0 CLASS`, `BEGIN`, `MultiUse = -1`, значит document-модуль был загружен неправильным способом и сборка считается дефектной.
- [x] Current Release Model: Релизная защита строится по схеме `штатный пароль VBA-проекта + Ghost Modules`.
- [x] Release Output: Релизные `.xlsm` складывать в папку `CreateOrderReleases` внутри проекта, а не в корень репозитория.
- [x] Verification: После сборки релиза проверять:
- [x] В `vbaProject.bin` отсутствуют `Module=...` для скрываемых модулей.
- [x] В релизе нет `DPx=`.
- [x] Книга открывается в Excel без recovery/repair-диалога.
