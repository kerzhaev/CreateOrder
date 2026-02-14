Attribute VB_Name = "mdlDataValidation"
' ==============================================================================
' Module: mdlDataValidation
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 14.02.2026
' Description: Validates sheet data using high-performance array processing.
'              Checks for empty fields and invalid date periods.
' ==============================================================================

Option Explicit

' /**
'  * Main validation procedure.
'  * Reads data into memory (Array), checks logic, and highlights errors in Red.
'  */
Public Sub ValidateMainSheetData()
    Dim ws As Worksheet
    Dim wsStaff As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim vData As Variant ' Data Array
    Dim errorCount As Long
    Dim strFIO As String, strID As String
    Dim dStart As Variant, dEnd As Variant
    Dim tStart As Double

    On Error GoTo ErrorHandler

    ' 1. Setup Environment
    tStart = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Подготовка данных..."

    ' 2. Validate Sheets existence
    Set ws = GetWorksheetSafeValidation() ' Active or Target
    If ws Is Nothing Then
        MsgBox "Ошибка: Лист для проверки не найден!", vbCritical, "Ошибка"
        GoTo CleanUp
    End If
    
    Set wsStaff = GetWorksheetSafeValidation("Штат")
    If wsStaff Is Nothing Then
        MsgBox "Ошибка: Не найден лист 'Штат'!", vbCritical, "Ошибка"
        GoTo CleanUp
    End If

    ' 3. Load Data into Memory (Array)
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row ' Column B (Name)
    If lastRow < 2 Then
        MsgBox "Нет данных для проверки (строк < 2).", vbInformation, "Инфо"
        GoTo CleanUp
    End If
    
    ' Determine last column (at least up to column 20 or actual data)
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 20 Then lastCol = 20

    ' Optimization: Read entire range to Array
    ' Value2 is faster and handles dates as doubles
    vData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value2

    ' Clear previous error highlights
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Interior.Pattern = xlNone

    errorCount = 0
    Application.StatusBar = "Анализ строк..."

    ' 4. Validation Loop (In Memory)
    For i = 2 To UBound(vData, 1)
        
        ' Check FIO (Column 2 / B)
        strFIO = CStr(vData(i, 2))
        If Len(Trim(strFIO)) = 0 Then
            MarkCellError ws, i, 2
            errorCount = errorCount + 1
        End If

        ' Check Personal ID (Column 3 / C)
        strID = CStr(vData(i, 3))
        If Len(Trim(strID)) = 0 Then
            MarkCellError ws, i, 3
            errorCount = errorCount + 1
        End If

        ' Check Periods (Start from Column 5 / E, Step 2)
        ' Expected pair: StartDate, EndDate
        For j = 5 To UBound(vData, 2) - 1 Step 2
            dStart = vData(i, j)
            dEnd = vData(i, j + 1)
            
            ' If at least one cell in pair is not empty
            If Not IsEmpty(dStart) Or Not IsEmpty(dEnd) Then
                ' 4.1 Check completeness
                If IsEmpty(dStart) Then
                    MarkCellError ws, i, j
                    errorCount = errorCount + 1
                ElseIf IsEmpty(dEnd) Then
                    MarkCellError ws, i, j + 1
                    errorCount = errorCount + 1
                Else
                    ' 4.2 Check Date Logic (Start <= End)
                    If IsNumeric(dStart) And IsNumeric(dEnd) Then
                        If CDbl(dStart) > CDbl(dEnd) Then
                            MarkCellError ws, i, j
                            MarkCellError ws, i, j + 1
                            errorCount = errorCount + 1
                        End If
                    Else
                        ' Not valid dates
                        MarkCellError ws, i, j
                        errorCount = errorCount + 1
                    End If
                End If
            End If
        Next j
        
        ' UI Feedback (every 100 rows)
        If i Mod 100 = 0 Then Application.StatusBar = "Проверено строк: " & i & " из " & lastRow
    Next i

    ' 5. Results
    Application.StatusBar = False
    Dim timeTaken As Double
    timeTaken = Round(Timer - tStart, 2)

    If errorCount > 0 Then
        MsgBox "Проверка завершена за " & timeTaken & " сек." & vbCrLf & _
               "Найдено ошибок: " & errorCount & "." & vbCrLf & _
               "Ошибочные ячейки выделены красным.", vbExclamation, "Результат"
    Else
        MsgBox "Ошибок не обнаружено!" & vbCrLf & _
               "Проверено строк: " & (lastRow - 1) & vbCrLf & _
               "Время: " & timeTaken & " сек.", vbInformation, "Успех"
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    MsgBox "Критическая ошибка валидации: " & Err.Description, vbCritical, "Ошибка " & Err.number
    Resume CleanUp
End Sub

' /**
'  * Helper: Colors a cell red.
'  * Accessed directly via Cells() as visual updates cannot be done via Array.
'  */
Private Sub MarkCellError(ws As Worksheet, r As Long, c As Long)
    On Error Resume Next
    ws.Cells(r, c).Interior.Color = vbRed
    On Error GoTo 0
End Sub

' /**
'  * Helper: Safely gets a worksheet by name or returns ActiveSheet.
'  */
Function GetWorksheetSafeValidation(Optional Name As String = "") As Worksheet
    On Error Resume Next
    If Name = "" Then
        Set GetWorksheetSafeValidation = ThisWorkbook.ActiveSheet
    Else
        Set GetWorksheetSafeValidation = ThisWorkbook.Sheets(Name)
    End If
    On Error GoTo 0
End Function

' /**
'  * Diagnostics: Analyzes workbook structure and reports to user.
'  */
Public Sub DiagnoseWorkbookStructure()
    Dim diagText As String
    Dim ws As Worksheet
    Dim wsCount As Integer
    
    On Error GoTo DiagError
    
    diagText = "=== ДИАГНОСТИКА СТРУКТУРЫ ===" & vbCrLf & vbCrLf
    diagText = diagText & "Файл: " & ThisWorkbook.Name & vbCrLf
    diagText = diagText & "Листов: " & ThisWorkbook.Worksheets.count & vbCrLf & vbCrLf
    
    diagText = diagText & "АНАЛИЗ ЛИСТОВ:" & vbCrLf
    wsCount = 1
    For Each ws In ThisWorkbook.Worksheets
        diagText = diagText & wsCount & ". " & ws.Name
        Dim lr As Long
        On Error Resume Next
        lr = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
        If lr > 1 Then
            diagText = diagText & " (строк: " & (lr - 1) & ")"
        Else
            diagText = diagText & " (пустой)"
        End If
        diagText = diagText & vbCrLf
        wsCount = wsCount + 1
    Next ws
    
    MsgBox diagText, vbInformation, "Диагностика"
    Exit Sub

DiagError:
    MsgBox "Ошибка диагностики: " & Err.Description, vbCritical
End Sub

' /**
'  * Emergency Stop: Resets Excel application state.
'  */
Public Sub StopValidation()
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Валидация остановлена.", vbInformation
End Sub

' /**
'  * Emergency Cleanup: Hard reset of environment variables.
'  */
Public Sub EmergencyCleanup()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Настройки Excel восстановлены.", vbInformation
End Sub

