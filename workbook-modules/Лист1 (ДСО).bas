' ===============================================================================
' Модуль событий листа ДСО для автоматической валидации
' Версия: 2.0.0 (Refactored & Safe)
' Дата: 14.02.2026
' Описание: Автоматическая валидация при вводе, подсветка ошибок, открытие поиска.
'           Интегрирован с mdlHelper для унификации логики дат.
' ===============================================================================

Option Explicit

' Флаг для предотвращения рекурсии (чтобы изменение цвета не вызывало событие Change снова)
Private isValidating As Boolean

' ===============================================================================
' СОБЫТИЕ: Двойной клик (Открытие формы поиска)
' ===============================================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rowNum As Long
    rowNum = Target.Row
    
    ' Проверяем, что клик не по заголовку
    If rowNum < 2 Then Exit Sub
    
    ' Проверяем, что клик по столбцам 2 (ФИО) или 3 (Личный номер)
    If Target.Column <> 2 And Target.Column <> 3 Then Exit Sub
    
    ' Отменяем вход в ячейку
    Cancel = True
    
    ' Получаем личный номер
    Dim lichniyNomer As String
    lichniyNomer = Trim(Me.Cells(rowNum, 3).value)
    
    ' Если есть номер - настраиваем форму
    If lichniyNomer <> "" Then
        frmSearchFIO.selectedLichniyNomer = lichniyNomer
        frmSearchFIO.FillByLichniyNomer
    End If
    
    ' Открываем форму
    frmSearchFIO.Show
End Sub

' ===============================================================================
' СОБЫТИЕ: Изменение ячеек (Валидация)
' ===============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Защита от рекурсии
    If isValidating Then Exit Sub
    
    ' Фильтрация: если меняется слишком много ячеек (например, удаление строки), выходим
    If Target.count > 50 Then Exit Sub
    
    ' Реагируем только на изменения в зоне периодов (столбец E и правее, строка 2+)
    If Target.Column < 5 Or Target.Row < 2 Then Exit Sub
    
    ' Включаем защиту
    isValidating = True
    
    ' Отключаем события для безопасности
    Dim originalEventsState As Boolean
    originalEventsState = Application.EnableEvents
    Application.EnableEvents = False
    
    On Error GoTo SafeCleanup
    
    ' Валидируем строку, в которой произошло изменение
    Dim targetRow As Long
    targetRow = Target.Cells(1, 1).Row
    
    If targetRow >= 2 And targetRow <= Me.Rows.count Then
        Call ValidateRowCompletely(targetRow)
    End If
    
SafeCleanup:
    ' Восстанавливаем состояние
    Application.EnableEvents = originalEventsState
    isValidating = False
    
    If Err.number <> 0 Then
        Debug.Print "Error in Worksheet_Change: " & Err.Description
        Err.Clear
    End If
End Sub

' ===============================================================================
' ЛОГИКА ВАЛИДАЦИИ
' ===============================================================================

' Комплексная валидация строки
Private Sub ValidateRowCompletely(rowNum As Long)
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim j As Long
    Dim periodDates As Variant
    Dim periodCount As Long
    Dim i As Integer
    
    Set ws = Me
    
    ' 1. ОЧИСТКА СТАРОГО ФОРМАТИРОВАНИЯ (Fix: убирает "вечный зеленый")
    ' Очищаем диапазон с 5 по 55 столбец (запас для ~25 периодов)
    ' Это гарантирует, что удаленные или пустые ячейки станут белыми
    With ws.Range(ws.Cells(rowNum, 5), ws.Cells(rowNum, 55))
        .Interior.ColorIndex = xlNone
        .ClearComments
    End With
    
    ' Определяем реальный последний столбец с данными
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then Exit Sub ' Нет периодов
    If lastCol > 54 Then lastCol = 54 ' Ограничиваем пределом (50+4)
    
    ' Собираем данные периодов в массив
    ReDim periodDates(1 To 25, 1 To 6)
    periodCount = 0
    
    j = 5
    Do While j + 1 <= lastCol And periodCount < 25
        Dim startValue As String, endValue As String
        Dim StartDate As Date, EndDate As Date
        
        startValue = GetCellValueSafeLocal(ws, rowNum, j)
        endValue = GetCellValueSafeLocal(ws, rowNum, j + 1)
        
        ' Если хотя бы одна ячейка заполнена
        If Len(startValue) > 0 Or Len(endValue) > 0 Then
            periodCount = periodCount + 1
            
            periodDates(periodCount, 1) = j            ' Col Start
            periodDates(periodCount, 2) = startValue   ' Val Start
            periodDates(periodCount, 3) = j + 1        ' Col End
            periodDates(periodCount, 4) = endValue     ' Val End
            
            ' Используем мощную проверку дат
            If IsValidDateLocal(startValue, StartDate) And IsValidDateLocal(endValue, EndDate) Then
                periodDates(periodCount, 5) = StartDate
                periodDates(periodCount, 6) = EndDate
            Else
                ' Маркер ошибки даты
                periodDates(periodCount, 5) = DateSerial(1900, 1, 1)
                periodDates(periodCount, 6) = DateSerial(1900, 1, 1)
            End If
        End If
        j = j + 2
    Loop
    
    ' 2. ВАЛИДАЦИЯ И ОКРАСКА
    For i = 1 To periodCount
        Call ValidatePeriodAndFormat(ws, rowNum, i, periodDates)
    Next i
    
    ' 3. ПРОВЕРКА ПЕРЕСЕЧЕНИЙ
    If periodCount > 1 Then
        Call CheckPeriodsIntersection(ws, rowNum, periodDates, periodCount)
    End If
End Sub

' Валидация конкретного периода
Private Sub ValidatePeriodAndFormat(ws As Worksheet, rowNum As Long, periodIndex As Integer, periodDates As Variant)
    Dim startCol As Long, endCol As Long
    Dim startValue As String, endValue As String
    Dim StartDate As Date, EndDate As Date
    Dim cutoffDate As Date
    
    Dim hasError As Boolean, hasWarning As Boolean
    Dim errorMsg As String, warningMsg As String
    
    ' Получаем данные из массива
    startCol = periodDates(periodIndex, 1)
    startValue = periodDates(periodIndex, 2)
    endCol = periodDates(periodIndex, 3)
    endValue = periodDates(periodIndex, 4)
    StartDate = periodDates(periodIndex, 5)
    EndDate = periodDates(periodIndex, 6)
    
    ' Получаем дату отсечения из mdlHelper (единый источник истины)
    cutoffDate = mdlHelper.GetExportCutoffDate()
    
    ' --- ПРОВЕРКИ ---
    
    ' 1. Неполная пара (одна ячейка пустая)
    If (Len(startValue) > 0 And Len(endValue) = 0) Or (Len(startValue) = 0 And Len(endValue) > 0) Then
        ApplyErrorFormat ws.Cells(rowNum, startCol), "Неполная пара дат"
        ApplyErrorFormat ws.Cells(rowNum, endCol), "Неполная пара дат"
        Exit Sub
    End If
    
    ' 2. Некорректный формат даты (текст, который не парсится)
    If StartDate = DateSerial(1900, 1, 1) Then
        ApplyErrorFormat ws.Cells(rowNum, startCol), "Некорректная дата"
        hasError = True
    End If
    If EndDate = DateSerial(1900, 1, 1) Then
        ApplyErrorFormat ws.Cells(rowNum, endCol), "Некорректная дата"
        hasError = True
    End If
    If hasError Then Exit Sub
    
    ' 3. Логика: Конец < Начала (Критическая)
    If EndDate < StartDate Then
        errorMsg = "Дата окончания меньше даты начала!"
        ApplyErrorFormat ws.Cells(rowNum, startCol), errorMsg
        ApplyErrorFormat ws.Cells(rowNum, endCol), errorMsg
        ws.Cells(rowNum, startCol).Interior.Color = RGB(255, 100, 100) ' Ярко-красный
        ws.Cells(rowNum, endCol).Interior.Color = RGB(255, 100, 100)
        Exit Sub
    End If
    
    ' 4. Устаревший период (Предупреждение)
    If EndDate < cutoffDate Then
        hasWarning = True
        warningMsg = "Период старше 3 лет - не войдет в приказ"
        ApplyWarningFormat ws.Cells(rowNum, startCol), warningMsg
        ApplyWarningFormat ws.Cells(rowNum, endCol), warningMsg
        Exit Sub
    End If
    
    ' 5. Даты в будущем (Предупреждение)
    If StartDate > Date Or EndDate > Date Then
        hasWarning = True
        warningMsg = "Дата в будущем"
        ApplyWarningFormat ws.Cells(rowNum, startCol), warningMsg
        ApplyWarningFormat ws.Cells(rowNum, endCol), warningMsg
        Exit Sub
    End If
    
    ' 6. УСПЕХ (Зеленый)
    ' Если дошли сюда, значит период валиден и ячейки не пустые
    ApplyValidFormat ws.Cells(rowNum, startCol)
    ApplyValidFormat ws.Cells(rowNum, endCol)
End Sub

' Проверка пересечений
Private Sub CheckPeriodsIntersection(ws As Worksheet, rowNum As Long, periodDates As Variant, periodCount As Long)
    Dim i As Integer, j As Integer
    Dim s1 As Date, e1 As Date, s2 As Date, e2 As Date
    Dim msg As String
    
    For i = 1 To periodCount - 1
        s1 = periodDates(i, 5): e1 = periodDates(i, 6)
        
        ' Пропускаем невалидные даты
        If s1 > DateSerial(1900, 1, 1) Then
            For j = i + 1 To periodCount
                s2 = periodDates(j, 5): e2 = periodDates(j, 6)
                
                If s2 > DateSerial(1900, 1, 1) Then
                    ' Пересечение: (Start1 <= End2) and (End1 >= Start2)
                    If s1 <= e2 And e1 >= s2 Then
                        msg = "Пересечение периодов!"
                        ApplyErrorFormat ws.Cells(rowNum, periodDates(i, 1)), msg
                        ApplyErrorFormat ws.Cells(rowNum, periodDates(i, 3)), msg
                        ApplyErrorFormat ws.Cells(rowNum, periodDates(j, 1)), msg
                        ApplyErrorFormat ws.Cells(rowNum, periodDates(j, 3)), msg
                    End If
                End If
            Next j
        End If
    Next i
End Sub

' ===============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ===============================================================================

' Безопасное получение текста ячейки
Private Function GetCellValueSafeLocal(ws As Worksheet, rowNum As Long, colNum As Long) As String
    On Error Resume Next
    GetCellValueSafeLocal = Trim(CStr(ws.Cells(rowNum, colNum).value))
End Function

' Проверка даты через mdlHelper (Fix 01.02.25)
Private Function IsValidDateLocal(dateString As String, ByRef resultDate As Date) As Boolean
    ' Делегируем разбор даты нашему мощному парсеру в mdlHelper
    resultDate = mdlHelper.ParseDateSafe(dateString)
    
    ' Проверяем, что вернулась адекватная дата (больше 2000 года)
    If resultDate > DateSerial(2000, 1, 1) And resultDate < DateSerial(2100, 12, 31) Then
        IsValidDateLocal = True
    Else
        IsValidDateLocal = False
    End If
End Function

' Форматирование: Ошибка (Красный)
Private Sub ApplyErrorFormat(cell As Range, message As String)
    On Error Resume Next
    cell.Interior.Color = RGB(255, 200, 200)
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment message
End Sub

' Форматирование: Предупреждение (Желтый)
Private Sub ApplyWarningFormat(cell As Range, message As String)
    On Error Resume Next
    cell.Interior.Color = RGB(255, 255, 200)
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
    cell.AddComment message
End Sub

' Форматирование: Успех (Зеленый)
Private Sub ApplyValidFormat(cell As Range)
    On Error Resume Next
    cell.Interior.Color = RGB(220, 255, 220)
    If Not cell.Comment Is Nothing Then cell.Comment.Delete
End Sub
