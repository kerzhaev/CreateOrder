# Технический план: Универсальная система работы с надбавками без периодов

**Версия:** 1.0.0  
**Дата:** 01.12.2025  
**Автор:** Кержаев Евгений, ФКУ "95 ФЭС" МО РФ  
**Основан на спецификации:** `.spec/universal-payments.md`

## 1. Обзор изменений

### 1.1. Новые модули VBA
1. `mdlReferenceData.bas` - работа со справочниками
2. `mdlPaymentTypes.bas` - конфигурация типов выплат
3. `mdlPaymentValidation.bas` - валидация надбавок
4. `mdlUniversalPaymentExport.bas` - универсальный экспорт надбавок

### 1.2. Изменяемые модули VBA
1. `mdlRibbonHandlers.bas` - добавление новых обработчиков для ленты
2. `mdlHelper.bas` - возможно добавление вспомогательных функций

### 1.3. Новые листы Excel
1. `Выплаты_Без_Периодов` - основная таблица с данными о надбавках
2. `Справочник_ВУС_Экипаж` - справочник пар ВУС-должность для экипажа
3. `Справочник_Типы_Выплат` - конфигурация типов выплат и их шаблонов

### 1.4. Обновление пользовательской ленты
- Обновление XML ленты (файл в Excel, не в modules)
- Добавление группы "Надбавки" с тремя кнопками

## 2. Детальный план реализации

### Этап 1: Создание структуры данных (листы Excel)

#### 1.1. Лист "Выплаты_Без_Периодов"
**Структура:**
- Колонка A: № (порядковый номер)
- Колонка B: Тип выплаты
- Колонка C: ФИО
- Колонка D: Личный номер
- Колонка E: Размер выплаты
- Колонка F: Основание

**Заголовки (строка 1):**
| № | Тип выплаты | ФИО | Личный номер | Размер выплаты | Основание |

**Реализация:** Создание листа программно при первом запуске или вручную пользователем.

#### 1.2. Лист "Справочник_ВУС_Экипаж"
**Структура:**
- Колонка A: ВУС
- Колонка B: Должность
- Колонка C: Примечание (опционально)

**Заголовки (строка 1):**
| ВУС | Должность | Примечание |

**Реализация:** Создание листа программно при первом запуске или вручную пользователем.

#### 1.3. Лист "Справочник_Типы_Выплат"
**Структура:**
- Колонка A: Тип выплаты
- Колонка B: Код
- Колонка C: Шаблон Word
- Колонка D: Описание

**Заголовки (строка 1):**
| Тип выплаты | Код | Шаблон Word | Описание |

**Начальные данные:**
- Водители СдЕ | DRIVER_SDE | Шаблон_Водители.docx | Надбавка водителям категории СдЕ
- Экипаж | CREW | Шаблон_Экипаж.docx | Надбавка за экипаж
- ФИЗО | FIZO | Шаблон_ФИЗО.docx | Надбавка за уровень ФИЗО
- Секретность | SECRECY | Шаблон_Секретность.docx | Надбавка за секретность

**Реализация:** Создание листа программно с начальными данными.

---

### Этап 2: Модуль работы со справочниками (`mdlReferenceData.bas`)

#### 2.1. Типы данных
```vba
' Тип для пары ВУС-Должность
Public Type VUSPositionPair
    VUS As String
    Position As String
End Type
```

#### 2.2. Основные функции

**`InitializeReferencesSheet(ws As Worksheet)`**
- Инициализация листа справочников (если создается программно)
- Создание структуры заголовков

**`LoadVUSPositionPairs() As Collection`**
- Загрузка всех пар ВУС-должность из листа "Справочник_ВУС_Экипаж"
- Возвращает коллекцию объектов VUSPositionPair

**`CheckVUSPositionPair(vus As String, position As String) As Boolean`**
- Проверка наличия пары (ВУС, должность) в справочнике
- Используется для валидации надбавки за экипаж

**`GetPaymentTypeConfig(paymentType As String) As PaymentTypeConfig`**
- Получение конфигурации типа выплаты из справочника "Справочник_Типы_Выплат"
- Возвращает объект PaymentTypeConfig (определен в mdlPaymentTypes.bas)

**`GetAllPaymentTypes() As Collection`**
- Получение списка всех типов выплат из справочника
- Возвращает коллекцию строк с названиями типов

---

### Этап 3: Модуль конфигурации типов выплат (`mdlPaymentTypes.bas`)

#### 3.1. Типы данных
```vba
' Тип для конфигурации типа выплаты
Public Type PaymentTypeConfig
    TypeName As String              ' "Водители СдЕ"
    TypeCode As String              ' "DRIVER_SDE"
    WordTemplate As String          ' "Шаблон_Водители.docx"
    Description As String           ' Описание
    DefaultTemplate As String       ' Единый шаблон по умолчанию
End Type

' Тип для данных о выплате без периодов
Public Type PaymentWithoutPeriod
    fio As String
    lichniyNomer As String
    Rank As String                  ' Из листа "Штат"
    Position As String              ' Из листа "Штат"
    VoinskayaChast As String        ' Из листа "Штат"
    PaymentType As String
    Amount As String
    Foundation As String
End Type
```

#### 3.2. Основные функции

**`GetPaymentTypeConfig(paymentType As String) As PaymentTypeConfig`**
- Получение конфигурации типа выплаты
- Загружает из справочника "Справочник_Типы_Выплат"
- Если тип не найден - возвращает конфигурацию по умолчанию

**`GetDefaultTemplate() As String`**
- Получение имени единого универсального шаблона
- Может быть константой или настройкой

**`GetTemplatePath(templateName As String) As String`**
- Получение полного пути к шаблону Word
- Проверка существования файла

---

### Этап 4: Модуль валидации (`mdlPaymentValidation.bas`)

#### 4.1. Основные функции

**`ValidatePaymentsWithoutPeriods()`**
- Главная функция валидации всех надбавок
- Проходит по всем строкам листа "Выплаты_Без_Периодов"
- Вызывает соответствующую функцию валидации для каждого типа выплаты
- Выводит отчет об ошибках

**`ValidateDriverSDE(ws As Worksheet, rowNum As Long) As Boolean`**
- Валидация надбавки водителям СдЕ
- Проверяет должность (должна быть "водитель" или "старший водитель", НЕ "механик-водитель")
- Проверяет наличие в основании: копия ВУ, справка ВАИ, марка машины, ГРЗ

**`ValidateCrew(ws As Worksheet, rowNum As Long) As Boolean`**
- Валидация надбавки за экипаж
- Получает ВУС и должность из листа "Штат" по личному номеру
- Проверяет пару (ВУС, должность) в справочнике "Справочник_ВУС_Экипаж"

**`ValidateFIZO(ws As Worksheet, rowNum As Long) As Boolean`**
- Валидация надбавки за ФИЗО
- Проверяет наличие минимум 2 ведомостей в основании
- Ищет подстроки "Ведомость" или "ведомость" в тексте основания

**`ValidateSecrecy(ws As Worksheet, rowNum As Long) As Boolean`**
- Валидация надбавки за секретность
- Проверяет формат основания: наличие "Форма:", "Номер:", "Дата:", "Пункт номенклатуры:"

**`GetEmployeeDataFromStaff(lichniyNomer As String) As Object`**
- Получение данных о военнослужащем из листа "Штат"
- Использует существующую функцию `mdlHelper.GetStaffData()`
- Возвращает Dictionary с данными

---

### Этап 5: Модуль универсального экспорта (`mdlUniversalPaymentExport.bas`)

#### 5.1. Основные функции

**`ExportPaymentsWithoutPeriods()`**
- Главная функция экспорта надбавок
- Собирает данные из листа "Выплаты_Без_Периодов"
- Группирует военнослужащих по типу выплаты
- Для каждого типа выплаты создает отдельный документ Word
- Использует соответствующий шаблон или формирует текст напрямую

**`CollectPaymentsData() As Collection`**
- Сбор всех данных о выплатах из листа "Выплаты_Без_Периодов"
- Для каждой строки получает данные о военнослужащем из листа "Штат"
- Возвращает коллекцию объектов PaymentWithoutPeriod

**`GroupPaymentsByType(payments As Collection) As Object`**
- Группировка выплат по типу
- Возвращает Dictionary, где ключ - тип выплаты, значение - коллекция выплат

**`GeneratePaymentOrder(paymentType As String, payments As Collection) As Boolean`**
- Генерация приказа Word для конкретного типа выплаты
- Логика выбора шаблона:
  1. Шаблон из справочника для типа выплаты
  2. Единый универсальный шаблон
  3. Формирование текста напрямую в Word
- Заполнение шаблона или создание текста приказа

**`FillPaymentTemplate(doc As Object, payment As PaymentWithoutPeriod) As Boolean`**
- Заполнение шаблона Word данными о выплате
- Замена плейсхолдеров: [ФИО], [ЗВАНИЕ], [ЛИЧНЫЙ_НОМЕР], [ДОЛЖНОСТЬ], [РАЗМЕР], [ОСНОВАНИЕ]

**`GeneratePaymentTextDirectly(doc As Object, payment As PaymentWithoutPeriod) As Boolean`**
- Формирование текста приказа напрямую в Word без шаблона
- Используется, если шаблоны отсутствуют
- Формат: "[Звание] [ФИО], личный номер [номер], [должность]. Размер: [размер]. Основание: [основание]"

---

### Этап 6: Обновление обработчиков ленты (`mdlRibbonHandlers.bas`)

#### 6.1. Новые обработчики

**`OnExportAllowancesClick(control As IRibbonControl)`**
- Обработчик кнопки "Экспорт надбавок"
- Вызывает `mdlUniversalPaymentExport.ExportPaymentsWithoutPeriods()`

**`OnValidateAllowancesClick(control As IRibbonControl)`**
- Обработчик кнопки "Проверить надбавки"
- Вызывает `mdlPaymentValidation.ValidatePaymentsWithoutPeriods()`

**`OnManageReferencesClick(control As IRibbonControl)`**
- Обработчик кнопки "Справочники"
- Открывает лист "Справочники" или создает его, если отсутствует
- Активирует лист для редактирования

---

### Этап 7: Обновление XML пользовательской ленты

#### 7.1. Добавление группы "Надбавки"
```xml
<group id="allowancesGroup" label="Надбавки">
  <button id="exportAllowances" 
          label="Экспорт надбавок" 
          onAction="OnExportAllowancesClick" 
          imageMso="ExportSharePointList"
          size="large"
          screentip="Экспорт надбавок"
          supertip="Создает приказы на выплату надбавок (водители СдЕ, экипаж, ФИЗО, секретность и др.)"/>
  
  <button id="validateAllowances" 
          label="Проверить надбавки" 
          onAction="OnValidateAllowancesClick" 
          imageMso="ReviewCheckDocument"
          size="normal"
          screentip="Валидация надбавок"
          supertip="Проверяет корректность данных в листе 'Выплаты_Без_Периодов'"/>
  
  <button id="manageReferences" 
          label="Справочники" 
          onAction="OnManageReferencesClick" 
          imageMso="DatabaseTable"
          size="normal"
          screentip="Управление справочниками"
          supertip="Открывает лист 'Справочники' для настройки ВУС, должностей и других справочных данных"/>
</group>
```

**Примечание:** XML ленты находится в самом Excel файле, не в папке modules. Пользователь должен обновить его вручную или мы предоставим инструкцию.

---

## 3. Последовательность реализации

### Шаг 1: Создание модулей (без функционала)
1. Создать `mdlReferenceData.bas` с заголовками функций
2. Создать `mdlPaymentTypes.bas` с типами данных
3. Создать `mdlPaymentValidation.bas` с заголовками функций
4. Создать `mdlUniversalPaymentExport.bas` с заголовками функций

### Шаг 2: Реализация работы со справочниками
1. Реализовать `mdlReferenceData.bas` полностью
2. Протестировать загрузку справочников

### Шаг 3: Реализация конфигурации типов выплат
1. Реализовать `mdlPaymentTypes.bas` полностью
2. Протестировать получение конфигураций

### Шаг 4: Реализация валидации
1. Реализовать `mdlPaymentValidation.bas` полностью
2. Протестировать валидацию для каждого типа выплаты

### Шаг 5: Реализация экспорта
1. Реализовать `mdlUniversalPaymentExport.bas` полностью
2. Протестировать экспорт с шаблонами и без

### Шаг 6: Интеграция с лентой
1. Добавить обработчики в `mdlRibbonHandlers.bas`
2. Обновить XML ленты (инструкция для пользователя)

### Шаг 7: Тестирование и отладка
1. Полное тестирование всех функций
2. Проверка обработки ошибок
3. Проверка производительности

---

## 4. Зависимости между модулями

```
mdlUniversalPaymentExport.bas
    ??? mdlPaymentTypes.bas (использует PaymentTypeConfig, PaymentWithoutPeriod)
    ??? mdlReferenceData.bas (использует GetPaymentTypeConfig)
    ??? mdlHelper.bas (использует GetStaffData, SklonitZvanie, SklonitFIO, SklonitDolzhnost)

mdlPaymentValidation.bas
    ??? mdlReferenceData.bas (использует CheckVUSPositionPair)
    ??? mdlHelper.bas (использует GetStaffData)

mdlRibbonHandlers.bas
    ??? mdlUniversalPaymentExport.bas (вызывает ExportPaymentsWithoutPeriods)
    ??? mdlPaymentValidation.bas (вызывает ValidatePaymentsWithoutPeriods)
```

---

## 5. Константы и настройки

### 5.1. Имена листов (константы)
```vba
Public Const SHEET_PAYMENTS_NO_PERIODS As String = "Выплаты_Без_Периодов"
Public Const SHEET_REF_VUS_CREW As String = "Справочник_ВУС_Экипаж"
Public Const SHEET_REF_PAYMENT_TYPES As String = "Справочник_Типы_Выплат"
Public Const SHEET_STAFF As String = "Штат"
```

### 5.2. Единый шаблон по умолчанию
```vba
Public Const DEFAULT_TEMPLATE As String = "Шаблон_Универсальный.docx"
```

### 5.3. Индексы колонок листа "Выплаты_Без_Периодов"
```vba
Public Const COL_NUMBER As Long = 1          ' A
Public Const COL_PAYMENT_TYPE As Long = 2    ' B
Public Const COL_FIO As Long = 3             ' C
Public Const COL_LICHNIY_NOMER As Long = 4  ' D
Public Const COL_AMOUNT As Long = 5          ' E
Public Const COL_FOUNDATION As Long = 6      ' F
```

---

## 6. Обработка ошибок

Все функции должны иметь:
- `On Error GoTo ErrorHandler` в начале
- Блок `ErrorHandler:` с обработкой ошибок
- Информативные сообщения для пользователя
- Очистку ресурсов (Set obj = Nothing)

---

## 7. Проверка типов

Перед реализацией проверить:
- Все вызовы функций из `mdlHelper.bas` на соответствие типов параметров
- Все вызовы функций из `mdlReferenceData.bas` на соответствие типов
- Все типы данных (PaymentTypeConfig, PaymentWithoutPeriod) определены корректно

---

## 8. Критерии готовности плана

- ? Все модули определены с функциями
- ? Структура данных описана
- ? Последовательность реализации определена
- ? Зависимости между модулями указаны
- ? Константы и настройки определены

---

**Статус плана:** Готов к утверждению

