# Спецификация: надёжность текущего релизного выпуска

**Дата:** 29.08.2026

**Статус:** реализовано и проверено в текущем выпуске

## Цель

Сделать выпуск текущей версии CreateOrder воспроизводимым в Windows PowerShell и
гарантировать, что созданный `CreateOrder_Release_*.xlsm` является корректным
Open XML-контейнером, открываемым Excel.

## Реализованные изменения

- Новые генераторы, экспортёры, установщики и V2-тесты в `tools/` сохранены в
  UTF-8 с BOM, поэтому Windows PowerShell корректно разбирает русские строки.
- Значения путей в параметрах новых скриптов задаются после `param`, а не через
  `$PSScriptRoot` в значении параметра по умолчанию; запуск без аргументов
  снова работает в Windows PowerShell.
- `Test-PersonnelActionWizardV2Safe.ps1` ожидает штатное завершение Excel/Word
  до пяти секунд вместо ложной ошибки при короткой задержке COM.
- `Build-Release.ps1` пересобирает ZIP-контейнер с путями Open XML через `/`,
  а не с Windows-разделителем `\`.

## Граница

- Бизнес-логика кадровых событий, выплат, форм и Word-экспорта не меняется.
- Рабочая `CreateOrder.xlsm` не изменяется процессом создания релизной копии.
- Старый некорректный релизный файл создан только в этой попытке и перенесён в
  `Trash/rejected-releases/`; он не является выпуском.

## Проверка

- `Test-PaymentsEnrollmentAcceptance.ps1` — `ACCEPTANCE_SMOKE_OK`.
- `Test-EnrollmentWizardV2Logic.ps1` и `Test-EnrollmentWizardV2Designer.ps1`.
- `Test-PersonnelActionWizardSafe.ps1`,
  `Test-PersonnelActionWizardV2Designer.ps1` и
  `Test-PersonnelActionWizardV2Safe.ps1`.
- Тарифная, ФИЗО-, медальная, компактная UI- и ribbon-проверки.
- Релиз `CreateOrder_Release_20260829_160203.xlsm`: архив содержит
  `[Content_Types].xml`, `xl/workbook.xml`, `xl/vbaProject.bin`, не содержит
  путей с `\`, открыт Excel read-only; подтверждены 29 листов и 61 VBA-компонент.

## Приёмочные критерии

1. Каждый затронутый `tools/*.ps1` запускается в Windows PowerShell без
   ошибки кодировки или пустого `$PSScriptRoot`.
2. Полный acceptance завершается строкой `ACCEPTANCE_SMOKE_OK`.
3. Релизная книга содержит корректную структуру Open XML и открывается Excel.
4. Релизная процедура не сохраняет изменения в рабочей книге.
