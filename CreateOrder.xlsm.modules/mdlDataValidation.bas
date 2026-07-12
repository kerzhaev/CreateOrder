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

Private Const LONG_PERIOD_WARNING_DAYS As Long = 120

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
    Set ws = FindDsoWorksheet()
    If ws Is Nothing Then
        If Not isSilent Then MsgBox "DSO sheet not found.", vbCritical, "Validation"
        GoTo CleanUp
    End If

    ' 2. Determine range
    lastRow = mdlHelper.GetLastRow(ws, "C")
    If lastRow < 2 Then
        If Not isSilent Then MsgBox "No data to validate (rows 2+ are empty).", vbInformation, "Validation"
        GoTo CleanUp
    End If

    errorCount = 0
    warningCount = 0
    processedRows = 0
    reportText = "====== Validation Report ======" & vbCrLf & vbCrLf
    reportText = reportText & "Date: " & Format(Now, "dd.mm.yyyy hh:mm") & vbCrLf
    reportText = reportText & "Rows: " & (lastRow - 1) & vbCrLf & vbCrLf

    ' 3. Loop through rows
    For i = 2 To lastRow
        Application.StatusBar = "Validating row " & i & " of " & lastRow
        Call ValidateRowLogic(ws, i, errorCount, warningCount)
        processedRows = processedRows + 1
    Next i

    ' 4. Final Report
    Application.StatusBar = False
    
    If Not isSilent Then
        If errorCount = 0 And warningCount = 0 Then
            reportText = reportText & "No issues found." & vbCrLf & "All rows are valid."
            MsgBox reportText, vbInformation, "Validation"
        Else
            reportText = reportText & "Errors highlighted: " & errorCount & vbCrLf
            reportText = reportText & "Warnings: " & warningCount & vbCrLf
            reportText = reportText & "Check highlighted cells and comments."
            MsgBox reportText, vbExclamation, "Validation Result"
        End If
    End If

    GoTo CleanUp

ErrorHandler:
    If Not isSilent Then MsgBox "Validation error: " & Err.Description, vbCritical, "Validation"
CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub

' /**
'  * Validates a single row (Columns E onwards).
'  * пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ, пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ, пїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ!
'  */
Private Sub ValidateRowLogic(ws As Worksheet, rowNum As Long, ByRef errCnt As Long, ByRef warnCnt As Long)
    Dim lastCol As Long, j As Long
    Dim startVal As Variant, endVal As Variant
    Dim dStart As Date, dEnd As Date
    Dim cutoffDate As Date
    Dim hasLocalError As Boolean
    Dim periodLengthDays As Long
    
    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 5 Then lastCol = 5
    If lastCol > 60 Then lastCol = 60
    
    ' 1. пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    With ws.Range(ws.Cells(rowNum, 5), ws.Cells(rowNum, 60))
        .Interior.ColorIndex = xlNone
        .ClearComments
    End With
    
    ' 2. пїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅ-пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ)
    Dim rawPeriods() As Variant
    Dim pCount As Long
    pCount = 0
    ReDim rawPeriods(1 To 30, 1 To 3) ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ: StartText, EndText, StartDate (пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
    
    For j = 5 To lastCol Step 2
        startVal = Trim(CStr(ws.Cells(rowNum, j).value))
        endVal = Trim(CStr(ws.Cells(rowNum, j + 1).value))
        
        If startVal <> "" Or endVal <> "" Then
            pCount = pCount + 1
            rawPeriods(pCount, 1) = startVal
            rawPeriods(pCount, 2) = endVal
            rawPeriods(pCount, 3) = mdlHelper.ParseDateSafe(startVal)
        End If
    Next j
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (Bubble Sort)
    If pCount > 1 Then
        Dim i As Long, k As Long
        Dim t1 As String, t2 As String, t3 As Date
        For i = 1 To pCount - 1
            For k = i + 1 To pCount
                If rawPeriods(i, 3) > rawPeriods(k, 3) Then
                    t1 = rawPeriods(i, 1): t2 = rawPeriods(i, 2): t3 = rawPeriods(i, 3)
                    rawPeriods(i, 1) = rawPeriods(k, 1): rawPeriods(i, 2) = rawPeriods(k, 2): rawPeriods(i, 3) = rawPeriods(k, 3)
                    rawPeriods(k, 1) = t1: rawPeriods(k, 2) = t2: rawPeriods(k, 3) = t3
                End If
            Next k
        Next i
    End If
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
    Dim colIdx As Long
    colIdx = 5
    For i = 1 To pCount
        ws.Cells(rowNum, colIdx).value = rawPeriods(i, 1)
        ws.Cells(rowNum, colIdx + 1).value = rawPeriods(i, 2)
        colIdx = colIdx + 2
    Next i
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
    If colIdx <= lastCol Then
        ws.Range(ws.Cells(rowNum, colIdx), ws.Cells(rowNum, 60)).ClearContents
    End If
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ lastCol пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    lastCol = colIdx - 1
    If lastCol < 5 Then Exit Sub
    
    ' 3. пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    Dim validPeriods() As Variant
    Dim vCount As Long
    vCount = 0
    ReDim validPeriods(1 To 30, 1 To 4) ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ: ColStart, ColEnd, DateStart, DateEnd
    
    For j = 5 To lastCol Step 2
        startVal = ws.Cells(rowNum, j).value
        endVal = ws.Cells(rowNum, j + 1).value
        hasLocalError = False
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ
        If (Trim(CStr(startVal)) <> "" And Trim(CStr(endVal)) = "") Or _
           (Trim(CStr(startVal)) = "" And Trim(CStr(endVal)) <> "") Then
            ApplyFormat ws.Cells(rowNum, j), 2 ' Red
            ApplyFormat ws.Cells(rowNum, j + 1), 2
            errCnt = errCnt + 1
            GoTo NextPair
        End If
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ
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
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
        vCount = vCount + 1
        validPeriods(vCount, 1) = j
        validPeriods(vCount, 2) = j + 1
        validPeriods(vCount, 3) = dStart
        validPeriods(vCount, 4) = dEnd
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ)
        If dStart > Date Or dEnd > Date Or dEnd < cutoffDate Then
            ApplyFormat ws.Cells(rowNum, j), 3 ' Yellow
            ApplyFormat ws.Cells(rowNum, j + 1), 3
            warnCnt = warnCnt + 1
        Else
            periodLengthDays = DateDiff("d", dStart, dEnd) + 1
            If periodLengthDays > LONG_PERIOD_WARNING_DAYS Then
                ApplyFormat ws.Cells(rowNum, j), 4
                ApplyFormat ws.Cells(rowNum, j + 1), 4
                On Error Resume Next
                ws.Cells(rowNum, j).AddComment LocalizeValidationText("validation.long_period_comment", GetLongPeriodCommentText())
                ResizeCommentBox ws.Cells(rowNum, j), 220, 95
                On Error GoTo 0
                warnCnt = warnCnt + 1
            Else
                ApplyFormat ws.Cells(rowNum, j), 1 ' Green
                ApplyFormat ws.Cells(rowNum, j + 1), 1
            End If
        End If
        
NextPair:
    Next j
    
    ' 4. пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ)
    If vCount > 1 Then
        Dim s1 As Date, e1 As Date, s2 As Date, e2 As Date
        
        For i = 1 To vCount - 1
            s1 = validPeriods(i, 3)
            e1 = validPeriods(i, 4)
            For k = i + 1 To vCount
                s2 = validPeriods(k, 3)
                e2 = validPeriods(k, 4)
                
                ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ
                If s1 <= e2 And e1 >= s2 Then
                    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
                    ApplyFormat ws.Cells(rowNum, validPeriods(i, 1)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(i, 2)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(k, 1)), 2
                    ApplyFormat ws.Cells(rowNum, validPeriods(k, 2)), 2
                    
                    On Error Resume Next
                    ws.Cells(rowNum, validPeriods(i, 1)).AddComment GetPeriodsOverlapCommentText()
                    ws.Cells(rowNum, validPeriods(k, 1)).AddComment GetPeriodsOverlapCommentText()
                    ResizeCommentBox ws.Cells(rowNum, validPeriods(i, 1)), 180, 55
                    ResizeCommentBox ws.Cells(rowNum, validPeriods(k, 1)), 180, 55
                    On Error GoTo 0
                    
                    errCnt = errCnt + 1
                End If
            Next k
        Next i
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
        Case 4: rng.Interior.Color = RGB(189, 215, 238) ' Light Blue
        Case Else: rng.Interior.ColorIndex = xlNone
    End Select
End Sub

Private Function LocalizeValidationText(ByVal localizationKey As String, ByVal fallback As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim languageCode As String

    On Error GoTo SafeFallback

    Set ws = ThisWorkbook.Worksheets("Localization")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For colIndex = 2 To lastCol
        languageCode = LCase$(Trim$(CStr(ws.Cells(1, colIndex).Value)))
        If languageCode = "ru" Then
            For rowIndex = 2 To lastRow
                If LCase$(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = LCase$(Trim$(localizationKey)) Then
                    If Len(CStr(ws.Cells(rowIndex, colIndex).Value)) > 0 Then
                        LocalizeValidationText = CStr(ws.Cells(rowIndex, colIndex).Value)
                        If Not IsLikelyBrokenLocalization(LocalizeValidationText) Then
                            Exit Function
                        End If
                        LocalizeValidationText = vbNullString
                    End If
                End If
            Next rowIndex
            Exit For
        End If
    Next colIndex

SafeFallback:
    If Len(LocalizeValidationText) = 0 Then LocalizeValidationText = fallback
End Function

Private Function IsLikelyBrokenLocalization(ByVal value As String) As Boolean
    Dim sample As String

    sample = Trim$(value)
    If Len(sample) = 0 Then Exit Function

    IsLikelyBrokenLocalization = _
        InStr(1, sample, "РїС—", vbTextCompare) > 0 Or _
        InStr(1, sample, "Гђ", vbTextCompare) > 0 Or _
        InStr(1, sample, "Г‘", vbTextCompare) > 0 Or _
        InStr(1, sample, "пїЅпїЅпїЅпїЅ", vbTextCompare) > 0
End Function

Private Function GetLongPeriodCommentText() As String
    GetLongPeriodCommentText = BuildUnicodeText(1044, 1083, 1080, 1085, 1085, 1099, 1081, 32, 1085, 1077, 1087, 1088, 1077, 1088, 1099, 1074, 1085, 1099, 1081, 32, 1087, 1077, 1088, 1080, 1086, 1076, 46, 32, 1055, 1088, 1086, 1074, 1077, 1088, 1100, 1090, 1077, 32, 1084, 1077, 1089, 1103, 1094, 32, 1080, 32, 1076, 1072, 1090, 1099, 46, 32, 1045, 1089, 1083, 1080, 32, 1086, 1076, 1085, 1072, 32, 1087, 1072, 1088, 1072, 32, 1076, 1072, 1090, 32, 1091, 1082, 1072, 1079, 1072, 1085, 1072, 32, 1086, 1089, 1086, 1079, 1085, 1072, 1085, 1085, 1086, 44, 32, 1080, 1089, 1087, 1088, 1072, 1074, 1083, 1077, 1085, 1080, 1077, 32, 1085, 1077, 32, 1090, 1088, 1077, 1073, 1091, 1077, 1090, 1089, 1103, 46)
End Function

Private Function GetPeriodsOverlapCommentText() As String
    GetPeriodsOverlapCommentText = BuildUnicodeText(1055, 1077, 1088, 1080, 1086, 1076, 1099, 32, 1087, 1077, 1088, 1077, 1089, 1077, 1082, 1072, 1102, 1090, 1089, 1103, 33)
End Function

Private Function BuildUnicodeText(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(codePoints) To UBound(codePoints)
        result = result & ChrW$(CLng(codePoints(i)))
    Next i

    BuildUnicodeText = result
End Function

Private Sub ResizeCommentBox(ByVal targetCell As Range, ByVal commentWidth As Double, ByVal commentHeight As Double)
    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then
        targetCell.Comment.Shape.Width = commentWidth
        targetCell.Comment.Shape.Height = commentHeight
    End If
    On Error GoTo 0
End Sub

' /**
'  * Diagnostics (Ribbon button).
'  * Checks if sheets exist.
'  */
' Handler for structure diagnostics (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ Self-Healing)
Public Sub DiagnoseWorkbookStructure()
    Dim msg As String
    Dim hasStaffSheet As Boolean
    Dim hasDsoSheet As Boolean
    Dim hasPaymentsSheet As Boolean
    Dim wsDso As Worksheet

    hasStaffSheet = SheetExistsSafe(mdlReferenceData.SHEET_STAFF)
    Set wsDso = FindDsoWorksheet()
    hasDsoSheet = Not wsDso Is Nothing
    hasPaymentsSheet = SheetExistsSafe(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS)

    msg = "=== Workbook Structure Diagnostics ===" & vbCrLf & vbCrLf
    msg = msg & IIf(hasStaffSheet, "Staff sheet: OK", "Staff sheet: missing") & vbCrLf
    msg = msg & IIf(hasDsoSheet, "DSO sheet: OK", "DSO sheet: missing") & vbCrLf
    msg = msg & IIf(hasPaymentsSheet, "Payments without periods sheet: OK", "Payments without periods sheet: missing") & vbCrLf

    If Not (hasStaffSheet And hasDsoSheet And hasPaymentsSheet) Then
        msg = msg & vbCrLf & "Some required sheets are missing." & vbCrLf & _
              "Run automatic recovery now?"

        If MsgBox(msg, vbYesNo + vbCritical, "Diagnostics: Error") = vbYes Then
            Call HealWorkbookStructure
        End If
    Else
        msg = msg & vbCrLf & "Workbook structure looks correct."
        MsgBox msg, vbInformation, "Diagnostics: OK"
    End If
End Sub
' ===============================================================================
' пїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (SELF-HEALING ARCHITECTURE)
' ===============================================================================

' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ)
Public Sub SilentCheckStructure()
    Dim hasStaffSheet As Boolean
    Dim hasDsoSheet As Boolean
    Dim hasPaymentsSheet As Boolean
    Dim hasEnrollmentSheet As Boolean
    Dim missingList As String
    Dim wsDso As Worksheet

    hasStaffSheet = SheetExistsSafe(mdlReferenceData.SHEET_STAFF)
    Set wsDso = FindDsoWorksheet()
    hasDsoSheet = Not wsDso Is Nothing
    hasPaymentsSheet = SheetExistsSafe(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS)
    hasEnrollmentSheet = SheetExistsSafe(mdlReferenceData.SHEET_ENROLLMENT)

    If hasStaffSheet And hasDsoSheet And hasPaymentsSheet And hasEnrollmentSheet Then
        If wsDso.Name <> GetDsoSheetName() Then
            On Error Resume Next
            wsDso.Name = GetDsoSheetName()
            On Error GoTo 0
        End If
        On Error Resume Next
        mdlPaymentPackageSupport.EnsurePaymentsFeatureInfrastructure
        mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
        On Error GoTo 0
        Exit Sub
    End If

    If Not hasStaffSheet Then missingList = missingList & "- " & DT("structure.sheet.staff", "Sheet") & " '" & mdlReferenceData.SHEET_STAFF & "'" & vbCrLf
    If Not hasDsoSheet Then missingList = missingList & "- " & DT("structure.sheet.sheet", "Sheet") & " '" & GetDsoSheetName() & "'" & vbCrLf
    If Not hasPaymentsSheet Then missingList = missingList & "- " & DT("structure.sheet.sheet", "Sheet") & " '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'" & vbCrLf
    If Not hasEnrollmentSheet Then missingList = missingList & "- " & DT("structure.sheet.sheet", "Sheet") & " '" & mdlReferenceData.SHEET_ENROLLMENT & "'" & vbCrLf

    If Len(missingList) > 0 Then
        If MsgBox(DT("structure.message.missing_sheets", "Required workbook sheets are missing:") & vbCrLf & vbCrLf & _
                  missingList & vbCrLf & _
                  DT("structure.message.run_recovery", "Run automatic structure recovery now?"), _
                  vbYesNo + vbCritical, DT("structure.caption.check", "Workbook structure check")) = vbYes Then
            Call HealWorkbookStructure
        End If
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = FindWorksheetByName(sheetName)
    SheetExists = Not (ws Is Nothing)
End Function

Private Function SheetExistsSafe(ByVal sheetName As String) As Boolean
    If Len(Trim$(sheetName)) = 0 Then
        SheetExistsSafe = False
    Else
        SheetExistsSafe = SheetExists(sheetName)
    End If
End Function

Private Function SheetExistsByIndex(ByVal sheetIndex As Long) As Boolean
    On Error Resume Next
    SheetExistsByIndex = Not ThisWorkbook.Worksheets(sheetIndex) Is Nothing
    On Error GoTo 0
End Function

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
Public Sub HealWorkbookStructure(Optional ByVal isSilent As Boolean = False)
    Application.ScreenUpdating = False
    Call RestoreSheetDSO
    Call RestoreSheetPayments
    Call RestoreSheetEnrollment
    On Error Resume Next
    mdlPaymentPackageSupport.EnsurePaymentsFeatureInfrastructure
    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    On Error GoTo 0
    Application.ScreenUpdating = True
    If Not isSilent Then
        MsgBox DT("structure.message.recovered", "Workbook structure restored."), vbInformation, DT("structure.caption.recovery", "Structure recovery")
    End If
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ
Private Sub RestoreSheetDSO()
    Dim ws As Worksheet, sheetName As String, i As Integer, colIndex As Integer
    sheetName = GetDsoSheetName()
    
    Set ws = FindDsoWorksheet()
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    ElseIf ws.Name <> sheetName Then
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    
    ' 1. пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    ws.Cells(1, 1).value = DT("dso.header.number", "No.")
    ws.Cells(1, 2).value = DT("dso.header.fio", "FIO")
    ws.Cells(1, 3).value = DT("dso.header.personal_number", "Personal number")
    ws.Cells(1, 4).value = DT("dso.header.basis", "Basis")
    
    ' 2. пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    colIndex = 5
    For i = 1 To 24
        ws.Cells(1, colIndex).value = DT("dso.header.start", "Start") & i
        ws.Cells(1, colIndex + 1).value = DT("dso.header.end", "End") & i
        colIndex = colIndex + 2
    Next i
    
    ' 3. пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    Call FormatHeaderRow(ws, colIndex - 1)
    
    ' 4. пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    ws.Columns(1).ColumnWidth = 6       ' пїЅ пїЅ/пїЅ
    ws.Columns(2).ColumnWidth = 35      ' пїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅ)
    ws.Columns(3).ColumnWidth = 15      ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    ws.Columns(4).ColumnWidth = 25      ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    
    ' пїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ
    For i = 5 To colIndex - 1
        ws.Columns(i).ColumnWidth = 11.5
    Next i
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
Private Sub RestoreSheetPayments()
    Dim ws As Worksheet, sheetName As String, headers As Variant, i As Integer
    
    On Error Resume Next
    sheetName = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS
    If sheetName = "" Then sheetName = DT("sheet.payments_no_periods", mdlHelper.Ru(1042, 1099, 1087, 1083, 1072, 1090, 1099, 95, 1041, 1077, 1079, 95, 1055, 1077, 1088, 1080, 1086, 1076, 1086, 1074))
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    headers = Array(DT("payments.header.row_number", "No."), DT("payments.header.payment_type", "Payment type"), DT("payments.header.fio", "FIO"), DT("payments.header.personal_number", "Personal number"), DT("payments.header.amount", "Amount"), DT("payments.header.basis", "Basis"), _
                    DT("payments.header.package_id", "Package ID"), DT("payments.header.mode", "Mode"), DT("payments.header.parameter", "Parameter"), DT("payments.header.shared_basis", "Shared basis"), DT("payments.header.group_export", "Grouped export"), _
                    DT("payments.header.note", "Note"), DT("payments.header.status", "Status"), DT("payments.header.source_enrollment_id", "Enrollment ID"))
    
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    Call FormatHeaderRow(ws, UBound(headers) + 1)
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    ws.Columns(1).ColumnWidth = 5
    ws.Columns(2).ColumnWidth = 26
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 16
    ws.Columns(5).ColumnWidth = 18
    ws.Columns(6).ColumnWidth = 34
    ws.Columns(7).ColumnWidth = 16
    ws.Columns(8).ColumnWidth = 12
    ws.Columns(9).ColumnWidth = 18
    ws.Columns(10).ColumnWidth = 34
    ws.Columns(11).ColumnWidth = 14
    ws.Columns(12).ColumnWidth = 24
    ws.Columns(13).ColumnWidth = 18
    ws.Columns(14).ColumnWidth = 14

    On Error Resume Next
    mdlPaymentPackageSupport.EnsurePaymentsSheetButtons ws
    On Error GoTo 0
End Sub

Private Function DT(ByVal key As String, ByVal fallback As String) As String
    DT = t(key, fallback)
End Function

Private Function GetDsoSheetName() As String
    GetDsoSheetName = mdlHelper.Ru(1044, 1057, 1054)
End Function

Private Function GetDsoSheetCodeName() As String
    GetDsoSheetCodeName = mdlHelper.Ru(1051, 1080, 1089, 1090) & "1"
End Function

Private Function FindDsoWorksheet() As Worksheet
    Set FindDsoWorksheet = FindWorksheetByName(GetDsoSheetName())
    If FindDsoWorksheet Is Nothing Then
        Set FindDsoWorksheet = FindWorksheetByCodeName(GetDsoSheetCodeName())
    End If
End Function

Private Function FindWorksheetByName(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim targetName As String

    targetName = NormalizeSheetName(sheetName)
    If Len(targetName) = 0 Then Exit Function

    For Each ws In ThisWorkbook.Worksheets
        If NormalizeSheetName(ws.Name) = targetName Then
            Set FindWorksheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function FindWorksheetByCodeName(ByVal codeName As String) As Worksheet
    Dim ws As Worksheet
    Dim targetCodeName As String

    targetCodeName = NormalizeSheetName(codeName)
    If Len(targetCodeName) = 0 Then Exit Function

    For Each ws In ThisWorkbook.Worksheets
        If NormalizeSheetName(ws.CodeName) = targetCodeName Then
            Set FindWorksheetByCodeName = ws
            Exit Function
        End If
    Next ws
End Function

Private Function NormalizeSheetName(ByVal value As String) As String
    Dim normalized As String

    normalized = Replace$(value, ChrW$(160), " ")
    normalized = Replace$(normalized, vbTab, " ")
    NormalizeSheetName = LCase$(Trim$(normalized))
End Function

Private Sub RestoreSheetEnrollment()
    On Error Resume Next
    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    On Error GoTo 0
End Sub

' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
Private Sub FormatHeaderRow(ws As Worksheet, lastCol As Integer)
    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    
    With headerRange
        ' пїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ)
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .Font.Bold = True
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ-пїЅпїЅпїЅпїЅпїЅ/пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ Office)
        .Interior.Color = RGB(217, 225, 242)
        
        ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ (пїЅпїЅпїЅпїЅпїЅ) пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ
    With headerRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    ws.Rows(1).RowHeight = 35
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    If Not ws.AutoFilterMode Then headerRange.AutoFilter
    
    ' пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ (Freeze Panes)
    Dim currentSheet As Worksheet
    Set currentSheet = ActiveSheet
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False ' пїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    
    currentSheet.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

