# Fix Encoding Issue - План реализации

## Этап 1: Подготовка (ручные действия в Excel)

### 1.1 Изменение кодовых имён листов
Пользователь должен вручную изменить кодовые имена листов в Excel (Alt+F11, окно Properties):

| Отображаемое имя | Старое кодовое имя | Новое кодовое имя |
|------------------|-------------------|-------------------|
| ДСО              | (определить)      | `wsDSO`           |
| Штат             | (определить)      | `wsStaff`         |
| Выплаты_Без_Периодов | (определить) | `wsPayments`      |
| Справочник_ВУС_Экипаж | (определить) | `wsRefVUSCrew`    |
| Справочник_Типы_Выплат | (определить) | `wsRefPaymentTypes` |

## Этап 2: Рефакторинг модулей (автоматически)

### 2.1 mdlHelper.bas
Заменить:
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`

### 2.2 mdlMainExport.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.3 mdlRaportExport.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.4 mdlSpravkaExport.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.5 mdlRiskExport.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.6 mdlFRPExport.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.7 mdlRibbonHandlers.bas
Заменить:
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`
- `ThisWorkbook.Sheets(mdlReferenceData.SHEET_REF_PAYMENT_TYPES)` → `wsRefPaymentTypes`

### 2.8 mdlReferenceData.bas
Вариант А (рекомендуется): Полностью убрать константы и использовать кодовые имена
Вариант Б: Оставить константы, но они не будут использоваться в коде

### 2.9 frmSelectEmployee.frm
Заменить:
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`

### 2.10 frmSearchFIO.frm
Заменить:
- `ThisWorkbook.Sheets("Штат")` → `wsStaff`
- `ThisWorkbook.Sheets("ДСО")` → `wsDSO`

## Порядок работы

1. Пользователь вручную изменяет кодовые имена листов в Excel
2. Пользователь экспортирует модули через VbaModuleManager (если нужно)
3. AI выполняет замену строковых имён на кодовые во всех модулях
4. Пользователь импортирует обновлённые модули
5. Тестирование компиляции

## Особенности
- В mdlDataValidation.bas используется параметр `Name` - это не имя листа, оставляем как есть
- Все замены должны быть последовательными и полными
- После рефакторинга код будет работать независимо от кодировки системы
