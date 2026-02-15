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
Public Sub ValidateMainSheetData()
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
        MsgBox "Лист 'ДСО' не найден!", vbCritical
        GoTo CleanUp
    End If

    ' 2. Determine range
    lastRow = mdlHelper.GetLastRow(ws, "C")
    If lastRow < 2 Then
        MsgBox "Нет данных для проверки (строки 2+).", vbInformation
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
    
    If errorCount = 0 And warningCount = 0 Then
        reportText = reportText & "ОШИБОК НЕ ОБНАРУЖЕНО." & vbCrLf & "Все даты корректны."
        MsgBox reportText, vbInformation, "Успех"
    Else
        reportText = reportText & "Найдено ошибок: " & errorCount & vbCrLf
        reportText = reportText & "Предупреждений: " & warningCount & vbCrLf
        reportText = reportText & "Проверьте ячейки, выделенные красным и желтым."
        MsgBox reportText, vbExclamation, "Результаты"
    End If

    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка валидации: " & Err.Description, vbCritical
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
    Dim lastCol As Long
    Dim j As Long
    Dim startVal As Variant, endVal As Variant
    Dim dStart As Date, dEnd As Date
    Dim cutoffDate As Date
    Dim isError As Boolean, isWarning As Boolean
    
    ' Get cutoff date from Helper (single source of truth)
    cutoffDate = mdlHelper.GetExportCutoffDate()
    
    ' Determine last column in this row
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5
    ' Limit check to reasonable amount of columns to save performance
    If lastCol > 60 Then lastCol = 60
    
    ' Clear old formatting for the whole row stripe (columns 5 to 60)
    With ws.Range(ws.Cells(rowNum, 5), ws.Cells(rowNum, 60))
        .Interior.ColorIndex = xlNone
        .ClearComments
    End With
    
    ' Loop through period pairs
    For j = 5 To lastCol Step 2
        startVal = ws.Cells(rowNum, j).value
        endVal = ws.Cells(rowNum, j + 1).value
        
        isError = False
        isWarning = False
        
        ' Skip if both empty
        If (IsEmpty(startVal) Or Trim(CStr(startVal)) = "") And _
           (IsEmpty(endVal) Or Trim(CStr(endVal)) = "") Then
            ' Ensure it is white
            ws.Cells(rowNum, j).Interior.ColorIndex = xlNone
            ws.Cells(rowNum, j + 1).Interior.ColorIndex = xlNone
            GoTo NextPair
        End If
        
        ' Check 1: Incomplete pair
        If (Trim(CStr(startVal)) <> "" And Trim(CStr(endVal)) = "") Or _
           (Trim(CStr(startVal)) = "" And Trim(CStr(endVal)) <> "") Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            GoTo NextPair
        End If
        
        ' Check 2: Parse Dates using Helper (The Fix for 01.02.25)
        dStart = mdlHelper.ParseDateSafe(startVal)
        dEnd = mdlHelper.ParseDateSafe(endVal)
        
        ' If 0, parsing failed (or date is too old)
        If dStart = 0 Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            errCnt = errCnt + 1
            isError = True
        End If
        If dEnd = 0 Then
            ApplyFormat ws.Cells(rowNum, j + 1), 2 ' Red
            errCnt = errCnt + 1
            isError = True
        End If
        
        If isError Then GoTo NextPair
        
        ' Check 3: Logic (End < Start)
        If dEnd < dStart Then
            ApplyFormat ws.Cells(rowNum, j), 2
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            GoTo NextPair
        End If
        
        ' Check 4: Future dates
        If dStart > Date Or dEnd > Date Then
            ApplyFormat ws.Cells(rowNum, j), 3 ' Yellow
            ApplyFormat ws.Cells(rowNum, j + 1), 3
            warnCnt = warnCnt + 1
            isWarning = True
        End If
        
        ' Check 5: Old dates (Cutoff)
        If dEnd < cutoffDate Then
            ApplyFormat ws.Cells(rowNum, j), 3 ' Yellow
            ApplyFormat ws.Cells(rowNum, j + 1), 3
            warnCnt = warnCnt + 1
            isWarning = True
        End If
        
        ' Success: Green (only if no error/warning)
        If Not isWarning Then
            ApplyFormat ws.Cells(rowNum, j), 1 ' Green
            ApplyFormat ws.Cells(rowNum, j + 1), 1
        End If
        
NextPair:
    Next j
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
Public Sub DiagnoseWorkbookStructure()
    Dim msg As String
    msg = "Диагностика:" & vbCrLf
    
    On Error Resume Next
    If Not ThisWorkbook.Sheets("ДСО") Is Nothing Then
        msg = msg & "[OK] Лист ДСО найден." & vbCrLf
    Else
        msg = msg & "[FAIL] Лист ДСО отсутствует!" & vbCrLf
    End If
    
    If Not ThisWorkbook.Sheets("Штат") Is Nothing Then
        msg = msg & "[OK] Лист Штат найден." & vbCrLf
    Else
        msg = msg & "[FAIL] Лист Штат отсутствует!" & vbCrLf
    End If
    
    MsgBox msg, vbInformation
End Sub

