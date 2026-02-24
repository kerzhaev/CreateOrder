Attribute VB_Name = "mdlDataValidation"
' ===============================================================================
' Module: mdlDataValidation
' Version: 5.0.0 (Refactored)
' Date: 14.02.2026
' Description: Bulk validation of the main sheet (DSO).
'              Now uses mdlHelper.ParseDateSafe for robust date checking.
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' /**
'  * Main entry point for the "Validate Data" ribbon button.
'  * Scans the entire DSO sheet and highlights errors.
'  */
' /**
'  * Main entry point for the "Validate Data" ribbon button.
'  * Scans the entire DSO sheet and highlights errors.
'  */
Public Sub ValidateMainSheetData(Optional isSilent As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorCount As Long
    Dim warningCount As Long
    Dim processedRows As Long
    Dim reportText As String

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Starting validation..."

    ' 1. Get Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ДСО")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        If Not isSilent Then MsgBox "Лист 'ДСО' не найден!", vbCritical
        GoTo CleanUp
    End If

    ' 2. Determine range
    lastRow = mdlHelper.GetLastRow(ws, "C")
    If lastRow < 2 Then
        If Not isSilent Then MsgBox "Нет данных для проверки (строки 2+).", vbInformation
        GoTo CleanUp
    End If

    errorCount = 0
    warningCount = 0
    processedRows = 0
    reportText = "====== ОТЧЁТ О ВАЛИДАЦИИ ======" & vbCrLf & vbCrLf
    reportText = reportText & "Дата: " & Format(Now, "dd.mm.yyyy hh:mm") & vbCrLf
    reportText = reportText & "Строк: " & (lastRow - 1) & vbCrLf & vbCrLf

    ' 3. Loop through rows
    For i = 2 To lastRow
        Application.StatusBar = "Проверка строки " & i & " из " & lastRow
        Call ValidateRowLogic(ws, i, errorCount, warningCount)
        processedRows = processedRows + 1
    Next i

    ' 4. Final Report
    Application.StatusBar = False
    
    If Not isSilent Then
        If errorCount = 0 And warningCount = 0 Then
            reportText = reportText & "ОШИБОК НЕ ОБНАРУЖЕНО." & vbCrLf & "Все даты корректны."
            MsgBox reportText, vbInformation, "Успех"
        Else
            reportText = reportText & "Найдено ошибок (пересечения/опечатки): " & errorCount & vbCrLf
            reportText = reportText & "Предупреждений: " & warningCount & vbCrLf
            reportText = reportText & "Проверьте ячейки, выделенные красным и желтым."
            MsgBox reportText, vbExclamation, "Результаты"
        End If
    End If

    GoTo CleanUp

ErrorHandler:
    If Not isSilent Then MsgBox "Ошибка валидации: " & Err.Description, vbCritical
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

' /**
'  * Validates a single row (Columns E onwards).
'  * Applies formatting (Red/Yellow/Green) based on logic.
'  */
Private Sub ValidateRowLogic(ws As Worksheet, rowNum As Long, ByRef errCnt As Long, ByRef warnCnt As Long)
    Dim lastCol As Long, j As Long
    Dim startVal As Variant, endVal As Variant
    Dim dStart As Date, dEnd As Date
    Dim cutoffDate As Date
    Dim hasLocalError As Boolean
    
    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5
    If lastCol > 60 Then lastCol = 60
    
    ' 1. Очищаем старое форматирование
    With ws.Range(ws.Cells(rowNum, 5), ws.Cells(rowNum, 60))
        .Interior.ColorIndex = xlNone
        .ClearComments
    End With
    
    Dim validPeriods() As Variant
    Dim pCount As Long
    pCount = 0
    ReDim validPeriods(1 To 30, 1 To 4) ' Хранилище: ColStart, ColEnd, DateStart, DateEnd
    
    ' 2. Проверяем каждую пару индивидуально
    For j = 5 To lastCol Step 2
        startVal = ws.Cells(rowNum, j).value
        endVal = ws.Cells(rowNum, j + 1).value
        hasLocalError = False
        
        ' Пропуск пустых
        If (IsEmpty(startVal) Or Trim(CStr(startVal)) = "") And (IsEmpty(endVal) Or Trim(CStr(endVal)) = "") Then
            GoTo NextPair
        End If
        
        ' Неполная пара
        If (Trim(CStr(startVal)) <> "" And Trim(CStr(endVal)) = "") Or _
           (Trim(CStr(startVal)) = "" And Trim(CStr(endVal)) <> "") Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            GoTo NextPair
        End If
        
        ' Парсинг дат
        dStart = mdlHelper.ParseDateSafe(startVal)
        dEnd = mdlHelper.ParseDateSafe(endVal)
        
        If dStart = 0 Or dEnd = 0 Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            hasLocalError = True
        ElseIf dEnd < dStart Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            hasLocalError = True
        End If
        
        If hasLocalError Then GoTo NextPair
        
        ' Сохраняем валидную пару для перекрестной проверки
        pCount = pCount + 1
        validPeriods(pCount, 1) = j
        validPeriods(pCount, 2) = j + 1
        validPeriods(pCount, 3) = dStart
        validPeriods(pCount, 4) = dEnd
        
        ' Проверка на предупреждения (устаревшие или будущие даты)
        If dStart > Date Or dEnd > Date Or dEnd < cutoffDate Then
            ApplyFormat ws.Cells(rowNum, j), 3 ' Yellow
            ApplyFormat ws.Cells(rowNum, j + 1), 3
            warnCnt = warnCnt + 1
        Else
            ApplyFormat ws.Cells(rowNum, j), 1 ' Green
            ApplyFormat ws.Cells(rowNum, j + 1), 1
        End If
        
NextPair:
    Next j
    
    ' 3. ПРОВЕРКА ПЕРЕСЕЧЕНИЙ И ДУБЛИКАТОВ (Нахлест дат)
    If pCount > 1 Then
        Dim k As Long, m As Long
        Dim s1 As Date, e1 As Date, s2 As Date, e2 As Date
        
        For k = 1 To pCount - 1
            s1 = validPeriods(k, 3)
            e1 = validPeriods(k, 4)
            For m = k + 1 To pCount
                s2 = validPeriods(m, 3)
                e2 = validPeriods(m, 4)
                
                ' Жесткая математическая проверка пересечения дат (нахлеста)
                If s1 <= e2 And e1 >= s2 Then
                    ' Перекрашиваем обе конфликтующие ячейки в красный (перезаписывая зеленый)
                    ApplyFormat ws.Cells(rowNum, validPeriods(k, 1)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(k, 2)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(m, 1)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(m, 2)), 2
                    
                    On Error Resume Next
                    ws.Cells(rowNum, validPeriods(k, 1)).AddComment "Дубликат/Пересечение!"
                    ws.Cells(rowNum, validPeriods(m, 1)).AddComment "Дубликат/Пересечение!"
                    On Error GoTo 0
                    
                    errCnt = errCnt + 1
                End If
            Next m
        Next k
    End If
End Sub

' /**
'  * Helper to apply color.
'  * Mode: 1=Green, 2=Red, 3=Yellow
'  */
Private Sub ApplyFormat(rng As Range, mode As Integer)
    Select Case mode
        Case 1: rng.Interior.Color = RGB(220, 255, 220) ' Light Green
        Case 2: rng.Interior.Color = RGB(255, 100, 100) ' Red
        Case 3: rng.Interior.Color = RGB(255, 255, 200) ' Yellow
        Case Else: rng.Interior.ColorIndex = xlNone
    End Select
End Sub

' /**
'  * Diagnostics (Ribbon button).
'  * Checks if sheets exist.
'  */
' Handler for structure diagnostics (Обновленный с вызовом Self-Healing)
Public Sub DiagnoseWorkbookStructure()
    Dim msg As String
    Dim isOk As Boolean
    isOk = True
    
    msg = "=== РЕЗУЛЬТАТЫ ДИАГНОСТИКИ ===" & vbCrLf & vbCrLf
    
    ' Проверка листа Штат
    If Not Evaluate("ISREF('Штат'!A1)") Then
        msg = msg & "Лист 'Штат' не найден!" & vbCrLf
        isOk = False
    Else
        msg = msg & "Лист 'Штат' найден." & vbCrLf
    End If
    
    ' Проверка ДСО
    If Not Evaluate("ISREF('ДСО'!A1)") Then
        msg = msg & "Лист 'ДСО' не найден!" & vbCrLf
        isOk = False
    Else
        msg = msg & "Лист 'ДСО' найден." & vbCrLf
    End If
    
    ' Проверка Надбавок
    If Not Evaluate("ISREF('Надбавки без периодов'!A1)") Then
        msg = msg & "Лист 'Надбавки без периодов' не найден!" & vbCrLf
        isOk = False
    Else
        msg = msg & "Лист 'Надбавки без периодов' найден." & vbCrLf
    End If
    
    If Not isOk Then
        msg = msg & vbCrLf & "ВНИМАНИЕ: Обнаружены критические ошибки структуры." & vbCrLf & _
              "Хотите запустить автоматическое восстановление поврежденных/удаленных листов?"
              
        If MsgBox(msg, vbYesNo + vbCritical, "Диагностика: Ошибка") = vbYes Then
            Call HealWorkbookStructure
        End If
    Else
        msg = msg & vbCrLf & "Структура файла в идеальном состоянии."
        MsgBox msg, vbInformation, "Диагностика: ОК"
    End If
End Sub
' ===============================================================================
' БЛОК САМОВОССТАНОВЛЕНИЯ СТРУКТУРЫ (SELF-HEALING ARCHITECTURE)
' ===============================================================================

' Тихая проверка (Идеально для запуска при открытии файла)
Public Sub SilentCheckStructure()
    Dim isOk As Boolean
    isOk = True
    
    ' Надежная объектная проверка вместо Evaluate
    If Not SheetExists("Штат") Then isOk = False
    If Not SheetExists("ДСО") Then isOk = False
    If Not SheetExists(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS) Then isOk = False
    
    If Not isOk Then
        If MsgBox("ВНИМАНИЕ: Обнаружено отсутствие обязательных листов системы." & vbCrLf & _
                  "Запустить автоматическое восстановление структуры?", _
                  vbYesNo + vbCritical, "Повреждение структуры файла") = vbYes Then
            Call HealWorkbookStructure
        End If
    End If
End Sub

' Вспомогательная пуленепробиваемая функция проверки существования листа
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Trim(sheetName)) ' Trim убирает случайные пробелы
    On Error GoTo 0
    SheetExists = Not (ws Is Nothing)
End Function

' Главная функция восстановления
Public Sub HealWorkbookStructure()
    Application.ScreenUpdating = False
    Call RestoreSheetDSO
    Call RestoreSheetPayments
    Application.ScreenUpdating = True
    MsgBox "Структура книги успешно проверена и восстановлена!", vbInformation, "Самовосстановление"
End Sub

' Восстановление листа ДСО
Private Sub RestoreSheetDSO()
    Dim ws As Worksheet, sheetName As String, i As Integer, colIndex As Integer
    sheetName = "ДСО"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear ' Полностью очищаем лист для чистоты формата
    
    ' 1. Базовые заголовки
    ws.Cells(1, 1).value = "№ п/п"
    ws.Cells(1, 2).value = "ФИО"
    ws.Cells(1, 3).value = "Личный номер"
    ws.Cells(1, 4).value = "Основание"
    
    ' 2. Генерация периодов
    colIndex = 5
    For i = 1 To 24
        ws.Cells(1, colIndex).value = "Начало" & i
        ws.Cells(1, colIndex + 1).value = "Конец" & i
        colIndex = colIndex + 2
    Next i
    
    ' 3. Применяем форматирование шапки
    Call FormatHeaderRow(ws, colIndex - 1)
    
    ' 4. Жестко задаем КРАСИВУЮ ширину столбцов
    ws.Columns(1).ColumnWidth = 6       ' № п/п
    ws.Columns(2).ColumnWidth = 35      ' ФИО (широкое поле)
    ws.Columns(3).ColumnWidth = 15      ' Личный номер
    ws.Columns(4).ColumnWidth = 25      ' Основание
    
    ' Для дат делаем аккуратную ширину
    For i = 5 To colIndex - 1
        ws.Columns(i).ColumnWidth = 11.5
    Next i
End Sub

' Восстановление листа Надбавок
Private Sub RestoreSheetPayments()
    Dim ws As Worksheet, sheetName As String, headers As Variant, i As Integer
    
    On Error Resume Next
    sheetName = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS
    If sheetName = "" Then sheetName = "Надбавки без периодов"
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    headers = Array("№", "Тип выплаты", "ФИО", "Личный номер", "Размер выплаты", "Основание")
    
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i
    
    ' Применяем форматирование шапки
    Call FormatHeaderRow(ws, UBound(headers) + 1)
    
    ' Жестко задаем КРАСИВУЮ ширину столбцов
    ws.Columns(1).ColumnWidth = 5       ' №
    ws.Columns(2).ColumnWidth = 30      ' Тип выплаты
    ws.Columns(3).ColumnWidth = 35      ' ФИО
    ws.Columns(4).ColumnWidth = 15      ' Личный номер
    ws.Columns(5).ColumnWidth = 20      ' Размер выплаты
    ws.Columns(6).ColumnWidth = 35      ' Основание
End Sub

' Универсальная функция наведения красоты на шапку
Private Sub FormatHeaderRow(ws As Worksheet, lastCol As Integer)
    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    With headerRange
        ' Шрифт (Классический для документов)
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .Font.Bold = True
        
        ' Заливка (Приятный светло-синий/серый оттенок как в таблицах стиля Office)
        .Interior.Color = RGB(217, 225, 242)
        
        ' Выравнивание
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' Добавляем четкие границы (рамки) для шапки
    With headerRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' Высота шапки с запасом
    ws.Rows(1).RowHeight = 35
    
    ' Включаем автофильтр
    If Not ws.AutoFilterMode Then headerRange.AutoFilter
    
    ' Закрепление первой строки (Freeze Panes)
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False ' Чтобы не мигал экран при переключении
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    
    currentSheet.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

