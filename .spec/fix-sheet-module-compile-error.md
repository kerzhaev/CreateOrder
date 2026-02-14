# Исправление ошибки компиляции в модуле листа «ДСО»

**Ошибка:** `Compile error: Method or data member not found` на строке с `ThisWorkbook.Sheets("ДСО")` в процедуре `Worksheet_BeforeDoubleClick`.

Код этой процедуры находится в **модуле листа** (например, «ДСО» или «Лист1 (ДСО)») внутри книги .xlsm и не экспортируется в репозиторий.

---

## Вариант 1: Через вспомогательные функции (рекомендуется)

В модуле **mdlHelper** добавлены функции `GetSheetDSO()` и `GetSheetStaff()`.

В редакторе VBA откройте **модуль листа**, где срабатывает двойной клик (в дереве проекта: Microsoft Excel Objects → лист «ДСО» или с кодовым именем в скобках).

**Замените строку:**
```vba
Set wsDSO = ThisWorkbook.Sheets("ДСО")
```
**на:**
```vba
Set wsDSO = GetSheetDSO()
```

Если в том же модуле есть обращение к листу «Штат», замените:
```vba
Set wsStaff = ThisWorkbook.Sheets("Штат")
```
**на:**
```vba
Set wsStaff = GetSheetStaff()
```

Сохраните (Ctrl+S), затем **Отладка → Скомпилировать VBAProject**.

---

## Вариант 2: Замена Sheets на Worksheets

В том же модуле листа замените:
- `ThisWorkbook.Sheets("ДСО")` → `ThisWorkbook.Worksheets("ДСО")`
- `ThisWorkbook.Sheets("Штат")` → `ThisWorkbook.Worksheets("Штат")`

После правки: **Отладка → Скомпилировать VBAProject**.

---

## Если ошибка не исчезла: проверка ссылок

1. В VBA: **Сервис (Tools)** → **Ссылки (References...)**.
2. Убедитесь, что отмечена **«Microsoft Excel xx.0 Object Library»** (xx — версия, например 14.0 или 16.0).
3. Если стоит «ОТСУТСТВУЕТ» (MISSING) — снимите галочку и выберите рабочую ссылку на Excel из списка.
4. **OK**, затем снова **Отладка → Скомпилировать VBAProject**.
