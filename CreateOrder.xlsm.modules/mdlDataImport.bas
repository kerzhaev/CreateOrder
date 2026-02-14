Attribute VB_Name = "mdlDataImport"
' ==============================================================================
' Module: mdlDataImport
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 14.02.2026
' Description: Imports data from an external Excel file into the "Staff" (Shtat) sheet.
'              Replaces all existing data on the target sheet.
' ==============================================================================

Option Explicit

' /**
'  * Main procedure to import data to "Staff" sheet.
'  * Opens a file dialog, reads source data, and overwrites the "Staff" sheet.
'  */
Sub ImportDataToStaff()
    Dim sourceFile As Variant
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetWorksheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim selectedSheetIndex As Integer
    
    On Error GoTo ErrorHandler
    
    ' Optimization: Turn off updates
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' 1. Select File
    sourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Выберите файл для импорта данных на лист 'Штат'")
    
    If sourceFile = False Then
        MsgBox "Импорт отменен пользователем.", vbInformation, "Импорт данных"
        GoTo CleanUp
    End If
    
    ' 2. Open Source File
    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=True)
    
    ' 3. Select Worksheet
    If sourceWorkbook.Worksheets.count > 1 Then
        selectedSheetIndex = SelectWorksheetFromFile(sourceWorkbook)
        If selectedSheetIndex = 0 Then
            sourceWorkbook.Close False
            MsgBox "Импорт отменен - лист не выбран.", vbInformation, "Импорт данных"
            GoTo CleanUp
        End If
        Set sourceWorksheet = sourceWorkbook.Worksheets(selectedSheetIndex)
    Else
        Set sourceWorksheet = sourceWorkbook.Worksheets(1)
    End If
    
    ' 4. Prepare Target Sheet
    Set targetWorksheet = GetOrCreateStaffWorksheet()
    
    ' Determine data range
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.count, 1).End(xlUp).Row
    lastCol = sourceWorksheet.Cells(1, sourceWorksheet.Columns.count).End(xlToLeft).Column
    
    ' Check if empty
    If lastRow < 1 Or (lastRow = 1 And sourceWorksheet.Cells(1, 1).value = "") Then
        sourceWorkbook.Close False
        MsgBox "Выбранный лист пуст - нет данных для импорта.", vbExclamation, "Импорт данных"
        GoTo CleanUp
    End If
    
    ' 5. Perform Import (Copy/Paste to preserve formatting)
    ' Clear target first
    targetWorksheet.Cells.Clear
    
    ' Copy specific range (more efficient than copying whole columns)
    sourceWorksheet.Range(sourceWorksheet.Cells(1, 1), sourceWorksheet.Cells(lastRow, lastCol)).Copy
    targetWorksheet.Cells(1, 1).PasteSpecial xlPasteAll
    
    ' Cleanup UI
    targetWorksheet.Range("A1").Select
    Application.CutCopyMode = False
    
    ' Close Source
    sourceWorkbook.Close False
    
    ' 6. Finalize
    MsgBox "Данные успешно импортированы на лист 'Штат'!" & vbCrLf & _
           "Обработано строк: " & lastRow & vbCrLf & _
           "Столбцов: " & lastCol, vbInformation, "Успех"
    
    ' Re-initialize column indexes (critical step!)
    Call mdlHelper.InitStaffColumnIndexes
    
    GoTo CleanUp
    
ErrorHandler:
    If Not sourceWorkbook Is Nothing Then
        ' Close without saving if error occurred while open
        On Error Resume Next
        sourceWorkbook.Close False
        On Error GoTo 0
    End If
    MsgBox "Критическая ошибка при импорте: " & Err.Description, vbCritical, "Ошибка " & Err.number
    
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
    
    Set sourceWorkbook = Nothing
    Set sourceWorksheet = Nothing
    Set targetWorksheet = Nothing
End Sub

' /**
'  * Helper: Dialog to select a specific worksheet if multiple exist.
'  * Returns 0 if cancelled, or Index (1..N) if selected.
'  */
Function SelectWorksheetFromFile(wb As Workbook) As Integer
    Dim sheetList As String
    Dim ws As Worksheet
    Dim i As Integer
    Dim userInput As String
    
    ' Build list
    sheetList = "В файле найдено несколько листов:" & vbCrLf & vbCrLf
    i = 1
    For Each ws In wb.Worksheets
        ' Limit list to 20 sheets to prevent msgbox overflow
        If i <= 20 Then
            sheetList = sheetList & i & ". " & ws.Name & vbCrLf
        End If
        i = i + 1
    Next ws
    
    If i > 21 Then sheetList = sheetList & "... и еще " & (i - 21) & " листов."
    
    ' Prompt
    userInput = InputBox(sheetList & vbCrLf & "Введите НОМЕР листа для импорта (1-" & wb.Worksheets.count & "):", _
                        "Выбор источника", "1")
    
    ' Validation
    If userInput = "" Or Not IsNumeric(userInput) Then
        SelectWorksheetFromFile = 0
        Exit Function
    End If
    
    Dim val As Integer
    val = CInt(userInput)
    
    If val >= 1 And val <= wb.Worksheets.count Then
        SelectWorksheetFromFile = val
    Else
        MsgBox "Неверный номер листа!", vbExclamation
        SelectWorksheetFromFile = 0
    End If
End Function

' /**
'  * Helper: Gets the "Staff" (Shtat) worksheet or creates it if missing.
'  */
Function GetOrCreateStaffWorksheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Штат")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = "Штат"
        MsgBox "Лист 'Штат' не был найден и создан автоматически.", vbInformation, "Система"
    End If
    
    Set GetOrCreateStaffWorksheet = ws
End Function

' /**
'  * Helper: Preview the first 5 rows of a selected file without importing.
'  */
Sub PreviewImportData()
    Dim sourceFile As Variant
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim previewText As String
    Dim i As Long, j As Long
    Dim maxRows As Long, maxCols As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    sourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Выберите файл для предварительного просмотра")
    
    If sourceFile = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=True)
    Set sourceWorksheet = sourceWorkbook.Worksheets(1)
    
    ' Preview bounds
    Dim lr As Long, lc As Long
    lr = sourceWorksheet.Cells(sourceWorksheet.Rows.count, 1).End(xlUp).Row
    lc = sourceWorksheet.Cells(1, sourceWorksheet.Columns.count).End(xlToLeft).Column
    
    maxRows = IIf(lr > 5, 5, lr)
    maxCols = IIf(lc > 5, 5, lc)
    
    previewText = "=== ПРЕДПРОСМОТР (Первые 5 строк) ===" & vbCrLf & vbCrLf
    previewText = previewText & "Файл: " & sourceWorkbook.Name & vbCrLf
    previewText = previewText & "Лист: " & sourceWorksheet.Name & vbCrLf & vbCrLf
    
    For i = 1 To maxRows
        For j = 1 To maxCols
            previewText = previewText & Left(CStr(sourceWorksheet.Cells(i, j).value), 15) & vbTab
        Next j
        previewText = previewText & vbCrLf
    Next i
    
    sourceWorkbook.Close False
    
    MsgBox previewText, vbInformation, "Данные файла"
    
    GoTo CleanUp
    
ErrorHandler:
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close False
    MsgBox "Ошибка чтения файла: " & Err.Description, vbCritical
    
CleanUp:
    Application.ScreenUpdating = True
End Sub

