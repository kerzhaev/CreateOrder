# Release Protection Specification

- [x] Security: Расширить массив $modulesToHide в скрипте Build-Release.ps1. Добавить все модули бизнес-логики (mdlMainExport, mdlRaportExport, mdlSpravkaExport, mdlRiskExport, mdlUniversalPaymentExport, mdlFRPExport, mdlWordImport) для предотвращения прямого запуска макросов через Alt+F8 в обход ленты.
