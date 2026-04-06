# Release Protection Specification

- [x] Security: Расширить массив $modulesToHide в скрипте Build-Release.ps1. Добавить все модули бизнес-логики (mdlMainExport, mdlRaportExport, mdlSpravkaExport, mdlRiskExport, mdlUniversalPaymentExport, mdlFRPExport, mdlWordImport) для предотвращения прямого запуска макросов через Alt+F8 в обход ленты.
- [x] Pipeline Rule: Не использовать `DPB -> DPx` в релизном пайплайне как обязательный слой защиты. Этот хак вызывает `invalid key 'DPx'`, recovery-диалог Excel и не дает надежной защиты проекта.
- [x] Pipeline Rule: Патчить `xl/vbaProject.bin` только на уровне байтовых последовательностей фиксированной длины. Перекодировка всего `vbaProject.bin` в строку и обратно запрещена.
- [x] Current Release Model: Релизная защита строится по схеме `штатный пароль VBA-проекта + Ghost Modules`.
- [x] Verification: После сборки релиза проверять:
- [x] В `vbaProject.bin` отсутствуют `Module=...` для скрываемых модулей.
- [x] В релизе нет `DPx=`.
- [x] Книга открывается в Excel без recovery/repair-диалога.
