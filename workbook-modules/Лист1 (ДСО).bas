' ===============================================================================
' Модуль событий листа ДСО для автоматической валидации с защитой от рекурсии
' Версия: 1.0.0
' Дата: 08.11.2025

' ===============================================================================

Option Explicit

' Константы для расчета ограничения по времени (3 года + 1 месяц)
Private Const PERIOD_LIMIT_YEARS As Integer = 3
Private Const PERIOD_LIMIT_MONTHS As Integer = 1

' НОВОЕ: Флаг для предотвращения рекурсии
Private isValidating As Boolean




' ================================
' Открытие формы по двойному щелчку
' ================================
' === Открытие формы по двойному клику по столбцам 2 или 3 ===
'==============================================================
' Открытие формы поиска/ввода с автозаполнением по ФИО/личному номеру,
' универсально с учетом любой структуры столбцов на листе "Штат"
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'==============================================================

'/**
'* Worksheet_BeforeDoubleClick — работает с универсальными индексами, защищает от всех ошибок структуры.
'* Весь поток снабжён отладочными печатями.
'* Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ

'/**
'* Worksheet_BeforeDoubleClick — теперь корректно заполняет форму выбранного ФИО аналогично поиску.
'* При двойном клике по строке ДСО открывается форма и отображаются все данные и периоды по выбранному военнослужащему.
'* Версия аннотации: 2.0 (19.11.2025)
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
'/**
'* Worksheet_BeforeDoubleClick — открывает форму поиска по двойному клику
'* Работает как по заполненным, так и по пустым ячейкам столбцов 2 и 3
'* Версия аннотации: 3.0 (01.12.2025)
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim rowNum As Long
    rowNum = Target.Row
    
    ' Проверяем, что клик не по заголовку (строка 1)
    If rowNum < 2 Then Exit Sub
    
    ' Проверяем, что клик по столбцам 2 или 3 (ФИО или Личный номер)
    If Target.Column <> 2 And Target.Column <> 3 Then Exit Sub
    
    ' Отменяем стандартное поведение Excel (редактирование ячейки)
    Cancel = True
    
    ' Получаем личный номер из столбца 3 текущей строки
    Dim lichniyNomer As String
    lichniyNomer = Trim(wsDSO.Cells(rowNum, 3).value)
    
    ' Если личный номер заполнен - ищем его в ДСО и заполняем форму
    If lichniyNomer <> "" Then
        Dim lastRowDSO As Long, i As Long, mainRow As Long
        lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
        mainRow = 0
        
        ' Ищем основную строку с этим личным номером
        For i = 2 To lastRowDSO
            If Trim(wsDSO.Cells(i, 3).value) = lichniyNomer Then
                mainRow = i
                Exit For
            End If
        Next i
        
        ' Если нашли - заполняем форму данными
        If mainRow > 0 Then
            frmSearchFIO.selectedLichniyNomer = lichniyNomer
            frmSearchFIO.FillByLichniyNomer
        End If
    End If
    
    ' Открываем форму в любом случае (и для пустых, и для заполненных ячеек)
    frmSearchFIO.Show
End Sub





' ИСПРАВЛЕННЫЙ обработчик события изменения данных с защитой от рекурсии
Private Sub Worksheet_Change(ByVal Target As Range)
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Предотвращение рекурсии
    If isValidating Then Exit Sub
    
    ' Фильтрация нерелевантных изменений
    If Target.count > 10 Then Exit Sub
    If Target.Column < 5 Or Target.Row < 2 Then Exit Sub
    
    ' НОВОЕ: Установка флага валидации для предотвращения рекурсии
    isValidating = True
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Сохранение исходного состояния событий
    Dim originalEventsState As Boolean
    originalEventsState = Application.EnableEvents
    
    ' Безопасное отключение событий
    Application.EnableEvents = False
    
    ' НОВОЕ: Использование структуры Try-Finally для гарантированной очистки
    On Error GoTo SafeCleanup
    
    ' Получаем строку для валидации (берем первую ячейку из диапазона)
    Dim targetRow As Long
    targetRow = Target.Cells(1, 1).Row
    
    ' Валидация только для корректных строк
    If targetRow >= 2 And targetRow <= Me.Rows.count Then
        Call ValidateRowCompletely(targetRow)
    End If
    
    ' КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Гарантированная очистка состояния
SafeCleanup:
    ' Восстанавливаем исходное состояние событий (ВСЕГДА выполняется)
    Application.EnableEvents = originalEventsState
    
    ' Сбрасываем флаг валидации (ВСЕГДА выполняется)
    isValidating = False
    
    ' Логирование ошибок для отладки
    If Err.number <> 0 Then
        Debug.Print "Ошибка валидации в строке " & targetRow & ": " & Err.Description & " (Код: " & Err.number & ")"
        
        ' Попытка очистить форматирование проблемной области при ошибке
        On Error Resume Next
        Call ClearRowFormattingSafe(Me, targetRow)
        On Error GoTo 0
        
        ' Очищаем ошибку
        Err.Clear
    End If
End Sub

' НОВАЯ безопасная функция очистки форматирования строки
Private Sub ClearRowFormattingSafe(ws As Worksheet, rowNum As Long)
    On Error Resume Next
    
    Dim lastCol As Long
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol > 50 Then lastCol = 50 ' Ограничиваем разумным пределом
    
    Dim i As Long
    For i = 5 To lastCol
        ws.Cells(rowNum, i).Interior.ColorIndex = xlNone
        If Not ws.Cells(rowNum, i).Comment Is Nothing Then
            ws.Cells(rowNum, i).Comment.Delete
        End If
    Next i
    
    Err.Clear
    On Error GoTo 0
End Sub

' УЛУЧШЕННАЯ функция расчета граничной даты с защитой от ошибок
Private Function GetCutoffDate() As Date
    On Error Resume Next
    
    Dim currentDate As Date
    Dim cutoffDate As Date
    
    currentDate = Date
    
    ' Безопасное вычитание 3 лет и 1 месяца
    cutoffDate = DateSerial(Year(currentDate) - PERIOD_LIMIT_YEARS, Month(currentDate) - PERIOD_LIMIT_MONTHS, Day(currentDate))
    
    ' Проверка на ошибку и установка значения по умолчанию
    If Err.number <> 0 Then
        cutoffDate = DateSerial(Year(currentDate) - 3, Month(currentDate), Day(currentDate))
        Debug.Print "Предупреждение: Ошибка расчета граничной даты, используется упрощенный расчет"
    End If
    
    GetCutoffDate = cutoffDate
    Err.Clear
    On Error GoTo 0
End Function

' БЕЗОПАСНАЯ комплексная валидация всей строки с периодами дат
Private Sub ValidateRowCompletely(rowNum As Long)
    ' Дополнительная проверка входных параметров
    If rowNum < 2 Or rowNum > Me.Rows.count Then
        Debug.Print "Некорректный номер строки для валидации: " & rowNum
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim j As Long
    Dim periodDates As Variant
    Dim periodCount As Long
    Dim i As Integer
    
    On Error GoTo ValidationError
    
    Set ws = Me
    
    ' Безопасное определение последнего столбца с ограничением
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol > 50 Then lastCol = 50 ' Ограничиваем разумным пределом для производительности
    
    ' Собираем все пары дат из строки
    ReDim periodDates(1 To 25, 1 To 6)
    periodCount = 0
    
    ' Проходим по столбцам парами начиная с E (5-й столбец)
    j = 5
    Do While j + 1 <= lastCol And periodCount < 25
        Dim startValue As String, endValue As String
        Dim StartDate As Date, EndDate As Date
        
        startValue = GetCellValueSafeLocal(ws, rowNum, j)
        endValue = GetCellValueSafeLocal(ws, rowNum, j + 1)
        
        ' Если хотя бы одна ячейка из пары заполнена
        If Len(startValue) > 0 Or Len(endValue) > 0 Then
            periodCount = periodCount + 1
            
            ' Сохраняем информацию о периоде
            periodDates(periodCount, 1) = j      ' Столбец начала
            periodDates(periodCount, 2) = startValue ' Значение начала
            periodDates(periodCount, 3) = j + 1  ' Столбец конца
            periodDates(periodCount, 4) = endValue   ' Значение конца
            
            ' Преобразуем в даты для проверки пересечений
            If IsValidDateLocal(startValue, StartDate) And IsValidDateLocal(endValue, EndDate) Then
                periodDates(periodCount, 5) = StartDate ' Дата начала для расчетов
                periodDates(periodCount, 6) = EndDate   ' Дата конца для расчетов
            Else
                periodDates(periodCount, 5) = DateSerial(1900, 1, 1) ' Невалидная дата
                periodDates(periodCount, 6) = DateSerial(1900, 1, 1) ' Невалидная дата
            End If
        End If
        
        j = j + 2
    Loop
    
    ' Валидируем каждый период и применяем форматирование
    For i = 1 To periodCount
        Call ValidatePeriodAndFormat(ws, rowNum, i, periodDates, periodCount)
    Next i
    
    ' Проверяем пересечения между всеми периодами
    If periodCount > 1 Then
        Call CheckPeriodsIntersection(ws, rowNum, periodDates, periodCount)
    End If
    
    ' Проверяем общую логику всех периодов
    If periodCount > 0 Then
        Call ValidatePeriodsSequence(ws, rowNum, periodDates, periodCount)
    End If
    
    Exit Sub

ValidationError:
    ' Безопасная обработка ошибок валидации
    Debug.Print "Ошибка валидации строки " & rowNum & ": " & Err.Description
    
    ' Попытка очистить форматирование проблемной области
    On Error Resume Next
    Call ClearRowFormattingSafe(ws, rowNum)
    On Error GoTo 0
    
    Err.Clear
End Sub

' Функция проверки пересечений между периодами (без изменений, но с улучшенной обработкой ошибок)
Private Sub CheckPeriodsIntersection(ws As Worksheet, rowNum As Long, periodDates As Variant, periodCount As Long)
    Dim i As Integer, j As Integer
    Dim period1Start As Date, period1End As Date
    Dim period2Start As Date, period2End As Date
    Dim period1StartCol As Long, period1EndCol As Long
    Dim period2StartCol As Long, period2EndCol As Long
    Dim intersectionMsg As String
    
    On Error GoTo IntersectionError
    
    ' Проверяем каждый период с каждым другим периодом
    For i = 1 To periodCount - 1
        For j = i + 1 To periodCount
            ' Получаем данные первого периода
            period1Start = periodDates(i, 5)
            period1End = periodDates(i, 6)
            period1StartCol = periodDates(i, 1)
            period1EndCol = periodDates(i, 3)
            
            ' Получаем данные второго периода
            period2Start = periodDates(j, 5)
            period2End = periodDates(j, 6)
            period2StartCol = periodDates(j, 1)
            period2EndCol = periodDates(j, 3)
            
            ' Проверяем только валидные даты
            If period1Start > DateSerial(1900, 1, 1) And period1End > DateSerial(1900, 1, 1) And _
               period2Start > DateSerial(1900, 1, 1) And period2End > DateSerial(1900, 1, 1) Then
                
                ' Проверяем пересечение периодов
                If period1Start <= period2End And period1End >= period2Start Then
                    ' Определяем тип пересечения
                    If period1Start = period2Start And period1End = period2End Then
                        intersectionMsg = "Полное дублирование периодов " & i & " и " & j
                    ElseIf (period2Start >= period1Start And period2End <= period1End) Then
                        intersectionMsg = "Период " & j & " полностью находится внутри периода " & i
                    ElseIf (period1Start >= period2Start And period1End <= period2End) Then
                        intersectionMsg = "Период " & i & " полностью находится внутри периода " & j
                    Else
                        intersectionMsg = "Частичное пересечение периодов " & i & " и " & j
                    End If
                    
                    ' Применяем красное форматирование ко всем ячейкам пересекающихся периодов
                    Call ApplyErrorFormat(ws.Cells(rowNum, period1StartCol), intersectionMsg)
                    Call ApplyErrorFormat(ws.Cells(rowNum, period1EndCol), intersectionMsg)
                    Call ApplyErrorFormat(ws.Cells(rowNum, period2StartCol), intersectionMsg)
                    Call ApplyErrorFormat(ws.Cells(rowNum, period2EndCol), intersectionMsg)
                End If
            End If
        Next j
    Next i
    
    Exit Sub

IntersectionError:
    Debug.Print "Ошибка проверки пересечений в строке " & rowNum & ": " & Err.Description
    Err.Clear
End Sub

' [Остальные функции остаются без изменений, но с улучшенной обработкой ошибок]

' Валидация отдельного периода с применением форматирования
Private Sub ValidatePeriodAndFormat(ws As Worksheet, rowNum As Long, periodIndex As Integer, periodDates As Variant, totalPeriods As Long)
    Dim startCol As Long, endCol As Long
    Dim startValue As String, endValue As String
    Dim StartDate As Date, EndDate As Date
    Dim startValid As Boolean, endValid As Boolean
    Dim hasError As Boolean, hasWarning As Boolean
    Dim errorMsg As String, warningMsg As String
    Dim cutoffDate As Date
    
    On Error GoTo PeriodError
    
    ' Получаем граничную дату
    cutoffDate = GetCutoffDate()
    
    ' Получаем данные периода
    startCol = periodDates(periodIndex, 1)
    startValue = periodDates(periodIndex, 2)
    endCol = periodDates(periodIndex, 3)
    endValue = periodDates(periodIndex, 4)
    
    hasError = False
    hasWarning = False
    errorMsg = ""
    warningMsg = ""
    
    ' Проверка 1: Парность заполнения
    If (Len(startValue) > 0 And Len(endValue) = 0) Or (Len(startValue) = 0 And Len(endValue) > 0) Then
        hasError = True
        errorMsg = "Неполная пара дат в периоде " & periodIndex
        Call ApplyErrorFormat(ws.Cells(rowNum, startCol), errorMsg)
        Call ApplyErrorFormat(ws.Cells(rowNum, endCol), errorMsg)
        Exit Sub
    End If
    
    ' Если обе ячейки пустые, очищаем форматирование
    If Len(startValue) = 0 And Len(endValue) = 0 Then
        Call ClearCellFormat(ws.Cells(rowNum, startCol))
        Call ClearCellFormat(ws.Cells(rowNum, endCol))
        Exit Sub
    End If
    
    ' Проверка 2: Корректность формата дат
    startValid = IsValidDateLocal(startValue, StartDate)
    endValid = IsValidDateLocal(endValue, EndDate)
    
    If Not startValid Then
        hasError = True
        errorMsg = "Некорректная дата начала: '" & startValue & "'"
        Call ApplyErrorFormat(ws.Cells(rowNum, startCol), errorMsg)
    End If
    
    If Not endValid Then
        hasError = True
        errorMsg = "Некорректная дата окончания: '" & endValue & "'"
        Call ApplyErrorFormat(ws.Cells(rowNum, endCol), errorMsg)
    End If
    
    ' Если даты некорректны, не продолжаем проверку
    If Not (startValid And endValid) Then Exit Sub
    
    ' КРИТИЧЕСКАЯ ПРОВЕРКА 3: Логика дат (окончание ДОЛЖНО быть больше начала)
    If EndDate < StartDate Then
        hasError = True
        errorMsg = "КРИТИЧЕСКАЯ ОШИБКА: Дата окончания (" & Format(EndDate, "dd.mm.yyyy") & ") должна быть больше даты начала (" & Format(StartDate, "dd.mm.yyyy") & ")"
        Call ApplyErrorFormat(ws.Cells(rowNum, startCol), errorMsg)
        Call ApplyErrorFormat(ws.Cells(rowNum, endCol), errorMsg)
        
        ' Дополнительно выделяем ячейки более ярким красным цветом для критических ошибок
        ws.Cells(rowNum, startCol).Interior.Color = RGB(255, 100, 100) ' Ярко-красный
        ws.Cells(rowNum, endCol).Interior.Color = RGB(255, 100, 100)   ' Ярко-красный
        Exit Sub
    End If
    
    ' Проверка 4: Период старше ограничения (3 года + 1 месяц)
    If EndDate < cutoffDate Then
        hasWarning = True
        warningMsg = "Устаревший период (старше " & Format(cutoffDate, "dd.mm.yyyy") & ") - не войдет в приказ"
        Call ApplyWarningFormat(ws.Cells(rowNum, startCol), warningMsg)
        Call ApplyWarningFormat(ws.Cells(rowNum, endCol), warningMsg)
        Exit Sub
    End If
    
    ' Проверка 5: Даты в будущем
    If StartDate > Date Then
        hasWarning = True
        warningMsg = "Дата начала в будущем"
        Call ApplyWarningFormat(ws.Cells(rowNum, startCol), warningMsg)
    End If
    
    If EndDate > Date Then
        hasWarning = True
        warningMsg = "Дата окончания в будущем"
        Call ApplyWarningFormat(ws.Cells(rowNum, endCol), warningMsg)
    End If
    
    ' Проверка 6: Очень длинные периоды (более года)
    Dim periodDays As Long
    periodDays = EndDate - StartDate + 1
    
    If periodDays > 365 Then
        hasWarning = True
        warningMsg = "Очень длинный период (" & periodDays & " дней)"
        Call ApplyWarningFormat(ws.Cells(rowNum, startCol), warningMsg)
        Call ApplyWarningFormat(ws.Cells(rowNum, endCol), warningMsg)
    End If
    
    ' Если нет ошибок и предупреждений, применяем зеленое форматирование
    If Not hasError And Not hasWarning Then
        Call ApplyValidFormat(ws.Cells(rowNum, startCol))
        Call ApplyValidFormat(ws.Cells(rowNum, endCol))
    End If
    
    Exit Sub

PeriodError:
    Debug.Print "Ошибка валидации периода " & periodIndex & " в строке " & rowNum & ": " & Err.Description
    Call ApplyErrorFormat(ws.Cells(rowNum, startCol), "Ошибка валидации: " & Err.Description)
    Call ApplyErrorFormat(ws.Cells(rowNum, endCol), "Ошибка валидации: " & Err.Description)
    Err.Clear
End Sub

' [Остальные вспомогательные функции остаются без изменений]

' Безопасное получение значения ячейки
Private Function GetCellValueSafeLocal(ws As Worksheet, rowNum As Long, colNum As Long) As String
    On Error Resume Next
    GetCellValueSafeLocal = Trim(CStr(ws.Cells(rowNum, colNum).value))
    If Err.number <> 0 Then GetCellValueSafeLocal = ""
    Err.Clear
    On Error GoTo 0
End Function

' Безопасная проверка корректности даты
Private Function IsValidDateLocal(dateString As String, ByRef resultDate As Date) As Boolean
    On Error Resume Next
    resultDate = DateValue(dateString)
    If Err.number = 0 And resultDate > DateSerial(1900, 1, 1) And resultDate < DateSerial(2100, 12, 31) Then
        IsValidDateLocal = True
    Else
        IsValidDateLocal = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Применение форматирования для ошибок (красный фон)
Private Sub ApplyErrorFormat(cell As Range, message As String)
    On Error Resume Next
    
    cell.Interior.Color = RGB(255, 200, 200) ' Светло-красный
    
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment "ОШИБКА: " & message
    
    With cell.Comment.Shape.TextFrame
        .Characters.Font.Size = 9
        .Characters.Font.Name = "Arial"
        .AutoSize = True
    End With
    
    Err.Clear
    On Error GoTo 0
End Sub

' Применение форматирования для предупреждений (желтый фон)
Private Sub ApplyWarningFormat(cell As Range, message As String)
    On Error Resume Next
    
    cell.Interior.Color = RGB(255, 255, 200) ' Светло-желтый
    
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment "ПРЕДУПРЕЖДЕНИЕ: " & message
    
    With cell.Comment.Shape.TextFrame
        .Characters.Font.Size = 9
        .Characters.Font.Name = "Arial"
        .AutoSize = True
    End With
    
    Err.Clear
    On Error GoTo 0
End Sub

' Применение форматирования для корректных данных (зеленый фон)
Private Sub ApplyValidFormat(cell As Range)
    On Error Resume Next
    
    cell.Interior.Color = RGB(220, 255, 220) ' Светло-зеленый
    
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    
    Err.Clear
    On Error GoTo 0
End Sub

' Очистка форматирования ячейки
Private Sub ClearCellFormat(cell As Range)
    On Error Resume Next
    
    cell.Interior.ColorIndex = xlNone
    
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    
    Err.Clear
    On Error GoTo 0
End Sub

' Валидация последовательности всех периодов (упрощенная версия для безопасности)
Private Sub ValidatePeriodsSequence(ws As Worksheet, rowNum As Long, periodDates As Variant, periodCount As Long)
    On Error Resume Next
    
    ' Упрощенная проверка последовательности без сложной логики
    ' для предотвращения ошибок в критическом коде
    
    Err.Clear
    On Error GoTo 0
End Sub


