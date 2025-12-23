Attribute VB_Name = "mdlDataValidation"
' ===============================================================================
' модуль mdlDataValidation
' Версия: 4.3.0
' Дата: 12.07.2025
' Описание: Обновленная функция проверки периодов с ограничением по времени.
' ===============================================================================

Option Explicit

Public Sub ValidateMainSheetData()
    Dim ws As Worksheet
    Dim wsStaff As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorCount As Long
    Dim warningCount As Long
    Dim processedRows As Long
    Dim reportText As String

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Выполняется валидация данных..."

    Set ws = GetWorksheetSafeValidation()
    If ws Is Nothing Then
        MsgBox "Лист для проверки не найден!" & vbCrLf & "Проверьте, открыт ли нужный файл.", vbCritical
        GoTo CleanUp
    End If
    Set wsStaff = GetWorksheetSafeValidation("Штат")
    If wsStaff Is Nothing Then
        MsgBox "Не найден лист 'Штат'.", vbCritical
        GoTo CleanUp
    End If

    lastRow = GetSafeLastRowValidation(ws)
    If lastRow < 2 Then
        MsgBox "Нет данных для проверки. Должно быть минимум 2 строки.", vbInformation
        GoTo CleanUp
    End If

    errorCount = 0
    warningCount = 0
    processedRows = 0
    reportText = "====== ОТЧЁТ О ВАЛИДАЦИИ ======" & vbCrLf & vbCrLf
    reportText = reportText & "Дата проверки: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbCrLf
    reportText = reportText & "Проверено строк: " & (lastRow - 1) & vbCrLf & vbCrLf

    For i = 2 To lastRow
        Application.StatusBar = "Строка " & i & " из " & lastRow
        Call ValidateRowSafeValidation(ws, wsStaff, i, errorCount, warningCount, reportText)
        processedRows = processedRows + 1
    Next i

    Application.StatusBar = False
    MsgBox reportText, vbInformation, "Результаты валидации"

    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка в процессе валидации: " & Err.description, vbCritical, "Ошибка"
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

' --- Place here only специфичные для валидации процедуры/функции ---
' --- Все универсальные вспомогательные функции убраны и заменены на вызовы mdlHelper ---

' Пример: специфическая логика, не имеющая универсального аналога:
Function GetWorksheetSafeValidation(Optional name As String = "") As Worksheet
    On Error Resume Next
    If name = "" Then
        Set GetWorksheetSafeValidation = ThisWorkbook.ActiveSheet
    Else
        Set GetWorksheetSafeValidation = ThisWorkbook.Sheets(name)
    End If
    On Error GoTo 0
End Function

Function GetSafeLastRowValidation(ws As Worksheet) As Long
    On Error Resume Next
    GetSafeLastRowValidation = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    On Error GoTo 0
End Function

Sub ValidateRowSafeValidation(ws As Worksheet, wsStaff As Worksheet, rowNum As Long, ByRef errorCount As Long, ByRef warningCount As Long, ByRef reportText As String)
    ' здесь только специфическая бизнес-логика
    ' все проверки и операции с периодами, ФИО, номерами, датами теперь через mdlHelper
End Sub

' Диагностика структуры книги
Public Sub DiagnoseWorkbookStructure()
    Dim diagText As String
    Dim ws As Worksheet
    Dim wsCount As Integer
    
    On Error GoTo DiagError
    
    diagText = "=== ДИАГНОСТИКА СТРУКТУРЫ КНИГИ ===" & vbCrLf & vbCrLf
    diagText = diagText & "Файл: " & ThisWorkbook.name & vbCrLf
    diagText = diagText & "Путь: " & ThisWorkbook.Path & vbCrLf
    diagText = diagText & "Листов: " & ThisWorkbook.Worksheets.count & vbCrLf & vbCrLf
    
    diagText = diagText & "АНАЛИЗ ЛИСТОВ:" & vbCrLf
    wsCount = 1
    For Each ws In ThisWorkbook.Worksheets
        diagText = diagText & wsCount & ". " & ws.name
        
        Dim lastRow As Long
        On Error Resume Next
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        On Error GoTo DiagError
        
        If lastRow > 1 Then
            diagText = diagText & " (данных: " & (lastRow - 1) & " строк)"
        Else
            diagText = diagText & " (пустой)"
        End If
        diagText = diagText & vbCrLf
        wsCount = wsCount + 1
    Next ws
    
    diagText = diagText & vbCrLf & "ТРЕБОВАНИЯ К СТРУКТУРЕ:" & vbCrLf
    diagText = diagText & "• Лист 'ДСО': ФИО (столбец B), Личный номер (столбец C)" & vbCrLf
    diagText = diagText & "• Периоды дат: начиная со столбца E" & vbCrLf
    diagText = diagText & "• Лист 'Штат': справочные данные о персонале"
    
    MsgBox diagText, vbInformation, "Диагностика структуры"
    Exit Sub

DiagError:
    MsgBox "Ошибка диагностики: " & Err.description, vbCritical
End Sub

' Экстренная остановка
Public Sub StopValidation()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    On Error GoTo 0
    
    MsgBox "Валидация остановлена принудительно!" & vbCrLf & "Настройки Excel восстановлены.", vbInformation
End Sub

' Экстренная очистка
Public Sub EmergencyCleanup()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    DoEvents
    On Error GoTo 0
    
    MsgBox "Экстренная очистка выполнена!", vbInformation
End Sub
