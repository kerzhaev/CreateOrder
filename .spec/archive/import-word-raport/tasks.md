# Задачи: Импорт рапортов из Word

- [x] T001: [UI] Обновить `customUI14.xml`, добавив кнопку `importWordRaport` (imageMso="ImportWord").
- [x] T002: [UI] Добавить заглушку обработчика `RunWordRaportImport` в `mdlRibbonHandlers.bas`.
- [x] T003: [Module] Создать `mdlWordImport.bas` с Option Explicit.
- [x] T004: [Extract] Реализовать функцию `ConvertWordToTempHTML(filePath)`.
- [x] T005: [Transform] Создать структуру UDT `ImportedPeriod` и логику `ParseHTMLToDict(htmlPath)`.
- [x] T006: [Load] Реализовать `ApplyDictToDSOSheet(dict)`.
- [x] T007: [Load] Реализовать функцию сортировки периодов в строке после добавления.
- [x] T008: [Integration] Связать все этапы в главной процедуре `ExecuteWordImport` и подключить к кнопке Ribbon.
