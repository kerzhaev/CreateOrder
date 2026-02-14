Attribute VB_Name = "mdlHelper"
' ==============================================================================
' Module: mdlHelper
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 14.02.2026
' Description: Universal utility functions and procedures for the project.
'              Contains logic for staff lookup, declension (Russian grammar),
'              and date handling.
' ==============================================================================

Option Explicit

' Global variables for column indexes (Cached for performance)
Public colFIO_Global As Long
Public colLichniyNomer_Global As Long
Public colZvanie_Global As Long
Public colDolzhnost_Global As Long
Public colVoinskayaChast_Global As Long

' /**
'  * Initializes column indexes for the "Staff" (Shtat) sheet.
'  * Must be called before accessing global column variables.
'  */
Public Sub InitStaffColumnIndexes()
    Dim wsStaff As Worksheet
    Set wsStaff = ThisWorkbook.Sheets("Штат") ' Sheet: Staff
    
    If Not FindColumnNumbers(wsStaff, colLichniyNomer_Global, colZvanie_Global, colFIO_Global, colDolzhnost_Global, colVoinskayaChast_Global) Then
        MsgBox "Корректные индексы столбцов не удалось определить. Работа программы невозможна.", vbCritical, "Ошибка инициализации"
        End
    End If
End Sub

' /**
'  * Finds a row number containing specific personal data (Optimized).
'  * Uses Application.Match for high performance instead of looping.
'  *
'  * @param ws Worksheet - Target worksheet
'  * @param lichniyNomer String - Value to search for
'  * @param colNum Long - Column index to search in
'  * @return Long - Row number or 0 if not found
'  */
Public Function FindStaffRow(ws As Worksheet, lichniyNomer As String, colNum As Long) As Long
    Dim res As Variant
    
    ' Optimization: Use Excel's native Match function
    res = Application.Match(lichniyNomer, ws.Columns(colNum), 0)
    
    If IsError(res) Then
        FindStaffRow = 0
    Else
        FindStaffRow = CLng(res)
    End If
End Function

' /**
'  * Universal function to determine column indexes on the "Staff" sheet.
'  * Logic:
'  *   - "Name" (FIO): First text-only column with spaces.
'  *   - "Position": First text-only column.
'  *   - "ID", "Rank", "Unit": By header keywords.
'  *
'  * @return Boolean - True if all critical columns are found.
'  */
Public Function FindColumnNumbers(ws As Worksheet, ByRef colLichniyNomer As Long, ByRef colZvanie As Long, ByRef colFIO As Long, ByRef colDolzhnost As Long, ByRef colVoinskayaChast As Long) As Boolean
    Dim lastCol As Long, i As Long, headerText As String
    Dim foundFIO As Boolean, foundDolzhnost As Boolean
    Dim msgErr As String

    ' Reset variables
    colLichniyNomer = 0: colZvanie = 0: colFIO = 0: colDolzhnost = 0: colVoinskayaChast = 0
    foundFIO = False: foundDolzhnost = False
    msgErr = ""

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' 1. Find Personal ID (by header)
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "личный номер") > 0 Then
            colLichniyNomer = i
            Exit For
        End If
    Next i
    If colLichniyNomer = 0 Then msgErr = msgErr & "Не найден столбец 'Личный номер'." & vbCrLf

    ' 2. Find Rank (by header)
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "воинское звание") > 0 Then
            colZvanie = i
            Exit For
        End If
    Next i
    If colZvanie = 0 Then msgErr = msgErr & "Не найден столбец 'Воинское звание'." & vbCrLf

    ' 3. Find Unit (by header)
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "часть") > 0 Or InStr(headerText, "раздел персонала") > 0 Then
            colVoinskayaChast = i
            Exit For
        End If
    Next i
    If colVoinskayaChast = 0 Then msgErr = msgErr & "Не найден столбец 'Часть' или 'Раздел персонала'." & vbCrLf

    ' 4. Find FIO (Name) - heuristic check
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If headerText = "лицо" Then
            If IsTextFIOColumn(ws, i) Then
                colFIO = i
                foundFIO = True
                Exit For
            End If
        End If
    Next i
    If Not foundFIO Then msgErr = msgErr & "Не найден корректный столбец 'Лицо' (ФИО)." & vbCrLf

    ' 5. Find Position - heuristic check
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "штатная должность") > 0 Then
            If IsTextColumn(ws, i) Then
                colDolzhnost = i
                foundDolzhnost = True
                Exit For
            End If
        End If
    Next i
    If Not foundDolzhnost Then msgErr = msgErr & "Не найден корректный столбец 'Штатная должность'." & vbCrLf

    ' Validate results
    If colLichniyNomer > 0 And colZvanie > 0 And colFIO > 0 And colDolzhnost > 0 And colVoinskayaChast > 0 Then
        FindColumnNumbers = True
    Else
        FindColumnNumbers = False
        MsgBox "Ошибка при определении столбцов на листе 'Штат':" & vbCrLf & msgErr, vbCritical, "Ошибка структуры"
    End If
End Function

' /**
'  * Checks if a column likely contains FIO (Full Name).
'  * Heuristic: Mostly text, contains spaces, not numeric.
'  */
Private Function IsTextFIOColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim lastRow As Long, i As Long, value As String
    Dim textCount As Long, totalCount As Long
    
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    ' Optimize: Check first 50 rows only to speed up init
    Dim checkLimit As Long: checkLimit = IIf(lastRow > 50, 50, lastRow)
    
    For i = 2 To checkLimit
        value = Trim(ws.Cells(i, colNum).value)
        If value <> "" Then
            totalCount = totalCount + 1
            If ContainsLetters(value) And InStr(value, " ") > 0 And Not IsNumeric(value) Then
                textCount = textCount + 1
            End If
        End If
    Next i
    
    If totalCount > 0 Then
        IsTextFIOColumn = (textCount / totalCount) > 0.7
    Else
        IsTextFIOColumn = False
    End If
End Function

' /**
'  * Extracts the Unit Number from a text string.
'  * Example: "Unit 12345 text" -> "12345"
'  */
Public Function ExtractVoinskayaChast(inputText As String) As String
    Dim text As String, i As Long, result As String, inNumber As Boolean
    text = Trim(inputText): result = "": inNumber = False
    
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            result = result & Mid(text, i, 1): inNumber = True
        Else
            If inNumber And Len(result) >= 4 Then
                ExtractVoinskayaChast = result
                Exit Function
            End If
            If inNumber Then result = "": inNumber = False
        End If
    Next i
    
    If Len(result) >= 4 Then ExtractVoinskayaChast = result Else ExtractVoinskayaChast = inputText
End Function

' /**
'  * Checks if the period is considered "Actual" (recent).
'  */
Public Function IsPeriodActual(dateEnd As Date) As Boolean
    IsPeriodActual = (dateEnd >= GetExportCutoffDate())
End Function

' /**
'  * Calculates the cutoff date for export (3 years + 1 month back).
'  */
Public Function GetExportCutoffDate() As Date
    Dim currentDate As Date, cutoffYear As Integer, cutoffMonth As Integer, cutoffDay As Integer
    currentDate = Date
    cutoffYear = Year(currentDate) - 3
    cutoffMonth = Month(currentDate) - 1
    cutoffDay = Day(currentDate)
    
    If cutoffMonth <= 0 Then
        cutoffMonth = cutoffMonth + 12
        cutoffYear = cutoffYear - 1
    End If
    GetExportCutoffDate = DateSerial(cutoffYear, cutoffMonth, cutoffDay)
End Function

' /**
'  * Checks if a column is numeric (>80% numbers).
'  */
Public Function IsNumericColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, numericCount As Long, totalCount As Long, cellValue As String, lastRow As Long, checkRows As Long
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    checkRows = IIf(lastRow - 1 > 10, 10, lastRow - 1)
    
    For i = 2 To 2 + checkRows - 1
        cellValue = Trim(ws.Cells(i, colNum).value)
        If cellValue <> "" Then
            totalCount = totalCount + 1
            If IsNumeric(cellValue) Then numericCount = numericCount + 1
        End If
    Next i
    
    If totalCount > 0 Then IsNumericColumn = (numericCount / totalCount) > 0.8 Else IsNumericColumn = False
End Function

' /**
'  * Checks if a column is textual (>70% letters).
'  */
Public Function IsTextColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, textCount As Long, totalCount As Long, cellValue As String, lastRow As Long, checkRows As Long
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    checkRows = IIf(lastRow - 1 > 20, 20, lastRow - 1)
    
    For i = 2 To 2 + checkRows - 1
        cellValue = Trim(ws.Cells(i, colNum).value)
        If cellValue <> "" Then
            totalCount = totalCount + 1
            If ContainsLetters(cellValue) And Not IsNumeric(cellValue) Then textCount = textCount + 1
        End If
    Next i
    
    If totalCount > 0 Then IsTextColumn = (textCount / totalCount) > 0.7 Else IsTextColumn = False
End Function

' /**
'  * Checks if a column contains "Voinskaya chast" (Unit).
'  */
Public Function IsVoinskayaChastColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, voinskayaChastCount As Long, totalCount As Long, cellValue As String, lastRow As Long, checkRows As Long
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    checkRows = IIf(lastRow - 1 > 10, 10, lastRow - 1)
    
    For i = 2 To 2 + checkRows - 1
        cellValue = LCase(Trim(ws.Cells(i, colNum).value))
        If cellValue <> "" Then
            totalCount = totalCount + 1
            If InStr(cellValue, "войсковая часть") > 0 And ContainsNumbers(cellValue) Then voinskayaChastCount = voinskayaChastCount + 1
        End If
    Next i
    
    If totalCount > 0 Then IsVoinskayaChastColumn = (voinskayaChastCount / totalCount) > 0.7 Else IsVoinskayaChastColumn = False
End Function

' /**
'  * Helper: String contains numbers?
'  */
Public Function ContainsNumbers(text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If char >= "0" And char <= "9" Then ContainsNumbers = True: Exit Function
    Next i
    ContainsNumbers = False
End Function

' /**
'  * Helper: String contains letters?
'  */
Public Function ContainsLetters(text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If (char >= "А" And char <= "я") Or (char >= "A" And char <= "z") Then
            ContainsLetters = True: Exit Function
        End If
    Next i
    ContainsLetters = False
End Function

' /**
'  * Collects all period pairs (Start/End/Days) for a person from "DSO" sheet.
'  * Returns a Collection of Collections.
'  */
Public Sub CollectAllPersonPeriods(ws As Worksheet, rowNum As Long, periods As Collection)
    Dim lastCol As Long, j As Long, dateStart As Date, dateEnd As Date
    On Error GoTo ErrorHandler
    
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    j = 5 ' Start from column E
    
    Do While j + 1 <= lastCol
        If ws.Cells(rowNum, j).value <> "" And ws.Cells(rowNum, j + 1).value <> "" Then
            dateStart = DateValue(ws.Cells(rowNum, j).value)
            dateEnd = DateValue(ws.Cells(rowNum, j + 1).value)
            
            If IsDate(dateStart) And IsDate(dateEnd) Then
                Dim DaysCount As Long: DaysCount = DateDiff("d", dateStart, dateEnd) + 1
                Dim newPeriod As Collection: Set newPeriod = New Collection
                newPeriod.Add dateStart
                newPeriod.Add dateEnd
                newPeriod.Add DaysCount
                periods.Add newPeriod
            End If
        End If
        j = j + 2
    Loop
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

' /**
'  * Checks if any period pair is invalid (End < Start).
'  */
Public Function HasInvalidPair(periods As Collection) As Boolean
    Dim p As Collection
    For Each p In periods
        If p(2) < p(1) Then HasInvalidPair = True: Exit Function
    Next p
    HasInvalidPair = False
End Function

' /**
'  * Sorts a collection of periods by Start Date.
'  */
Public Function SortPeriodsByDateStart(periods As Collection) As Collection
    Dim arr() As Variant, i As Long, j As Long, n As Long
    n = periods.count
    If n <= 1 Then Set SortPeriodsByDateStart = periods: Exit Function
    
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = periods(i): Next i
    
    ' Bubble sort (acceptable for small N periods)
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i)(1) > arr(j)(1) Then
                Dim tmp As Collection: Set tmp = arr(i)
                Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
    
    Dim resCol As Collection: Set resCol = New Collection
    For i = 1 To n: resCol.Add arr(i): Next i
    Set SortPeriodsByDateStart = resCol
End Function

' /**
'  * Checks "DSO" sheet for critical errors (Red cells).
'  */
Public Function hasCriticalErrors() As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ДСО")
    If ws Is Nothing Then hasCriticalErrors = True: Exit Function
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim i As Long, j As Long
    
    ' Loop is unavoidable here as we check formatting (color)
    For i = 2 To lastRow
        For j = 5 To lastCol Step 2
            If ws.Cells(i, j).Interior.Color = RGB(255, 100, 100) Or ws.Cells(i, j).Interior.Color = RGB(255, 200, 200) Then
                hasCriticalErrors = True: Exit Function
            End If
        Next j
    Next i
    hasCriticalErrors = False
End Function

' /**
'  * Checks a specific row for date errors (End < Start).
'  */
Public Function CheckRowForDateErrors(ws As Worksheet, rowNum As Long) As Boolean
    Dim lastCol As Long, j As Long, startValue As String, endValue As String
    Dim dateStart As Date, dateEnd As Date, hasErrors As Boolean
    
    On Error GoTo ErrorHandler
    hasErrors = False
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol > 50 Then lastCol = 50
    
    j = 5
    Do While j + 1 <= lastCol
        startValue = Trim(ws.Cells(rowNum, j).text)
        endValue = Trim(ws.Cells(rowNum, j + 1).text)
        
        If startValue <> "" And endValue <> "" Then
            On Error Resume Next
            dateStart = DateValue(startValue)
            dateEnd = DateValue(endValue)
            Err.Clear
            On Error GoTo ErrorHandler
            
            If IsDate(startValue) And IsDate(endValue) Then
                dateStart = DateValue(startValue)
                dateEnd = DateValue(endValue)
                If dateEnd < dateStart Then hasErrors = True: Exit Do
            Else
                hasErrors = True: Exit Do
            End If
        End If
        j = j + 2
    Loop
    CheckRowForDateErrors = hasErrors
    Exit Function
ErrorHandler:
    CheckRowForDateErrors = True
End Function

' ==========================================================
' RUSSIAN LANGUAGE GRAMMAR FUNCTIONS
' (Declensions for Names, Ranks, Positions)
' ==========================================================

Public Function SklonitZvanie(zvanie As String) As String
    Dim result As String
    Dim lowerZvanie As String
    lowerZvanie = LCase(Trim(zvanie))
    Select Case lowerZvanie
        Case "рядовой": result = "Рядовому"
        Case "ефрейтор": result = "Ефрейтору"
        Case "младший сержант": result = "Младшему сержанту"
        Case "сержант": result = "Сержанту"
        Case "старший сержант": result = "Старшему сержанту"
        Case "старшина": result = "Старшине"
        Case "прапорщик": result = "Прапорщику"
        Case "старший прапорщик": result = "Старшему прапорщику"
        Case "младший лейтенант": result = "Младшему лейтенанту"
        Case "лейтенант": result = "Лейтенанту"
        Case "старший лейтенант": result = "Старшему лейтенанту"
        Case "капитан": result = "Капитану"
        Case "майор": result = "Майору"
        Case "подполковник": result = "Подполковнику"
        Case "полковник": result = "Полковнику"
        Case "генерал-майор": result = "Генерал-майору"
        Case "генерал-лейтенант": result = "Генерал-лейтенанту"
        Case "генерал-полковник": result = "Генерал-полковнику"
        Case "генерал армии": result = "Генералу армии"
        Case Else
            result = UCase(Left(zvanie, 1)) & LCase(Mid(zvanie, 2))
    End Select
    SklonitZvanie = result
End Function

Public Function SklonitDolzhnost(dolzhnost As String, VoinskayaChast As String) As String
    Dim keepWords As Variant, cutWords As Variant
    ' Dictionary of words to keep or cut
    keepWords = Array("роты", "взвода", "отделения", "расчета", "группы", "команды", "экипажа")
    cutWords = Array("отдельного", "гвардейской", "общевойсковой", "мотострелковой", "танковой", "воздушно-десантной", "артиллерийской", "инженерной", "связи", "десантно-штурмовой", "батальона", "полка", "бригады", "дивизии", "корпуса", "армии", "округа")
    
    Dim dolzhnostLower As String, result As String, lastKeepPos As Long, lastKeepWord As String
    Dim i As Long, pos As Long
    dolzhnostLower = LCase(dolzhnost)
    lastKeepPos = -1: lastKeepWord = ""
    
    For i = LBound(keepWords) To UBound(keepWords)
        pos = InStrRev(dolzhnostLower, keepWords(i))
        If pos > lastKeepPos Then lastKeepPos = pos: lastKeepWord = keepWords(i)
    Next i
    
    If lastKeepPos > 0 Then
        Dim endKeepPos As Long, cutPosition As Long
        endKeepPos = lastKeepPos + Len(lastKeepWord) - 1
        cutPosition = 0
        For i = LBound(cutWords) To UBound(cutWords)
            pos = InStr(endKeepPos + 1, dolzhnostLower, cutWords(i))
            If pos > 0 Then If cutPosition = 0 Or pos < cutPosition Then cutPosition = pos
        Next i
        If cutPosition > 0 Then
            Dim startCutPosition As Long
            startCutPosition = cutPosition
            For i = cutPosition - 1 To endKeepPos + 1 Step -1
                Dim char As String
                char = Mid(dolzhnostLower, i, 1)
                If IsNumeric(char) Or char = " " Then startCutPosition = i Else Exit For
            Next i
            result = LCase(Trim(Left(dolzhnost, startCutPosition - 1))) & " войсковой части " & VoinskayaChast
        Else
            result = LCase(Trim(Left(dolzhnost, endKeepPos))) & " войсковой части " & VoinskayaChast
        End If
    Else
        cutPosition = 0
        For i = LBound(cutWords) To UBound(cutWords)
            pos = InStr(dolzhnostLower, cutWords(i))
            If pos > 0 Then If cutPosition = 0 Or pos < cutPosition Then cutPosition = pos
        Next i
        If cutPosition > 0 Then
            startCutPosition = cutPosition
            For i = cutPosition - 1 To 1 Step -1
                Dim char2 As String
                char2 = Mid(dolzhnostLower, i, 1)
                If IsNumeric(char2) Or char2 = " " Then startCutPosition = i Else Exit For
            Next i
            result = LCase(Trim(Left(dolzhnost, startCutPosition - 1))) & " войсковой части " & VoinskayaChast
        Else
            result = LCase(dolzhnost) & " войсковой части " & VoinskayaChast
        End If
    End If
    result = SklonitVoennayaDolzhnost(result)
    SklonitDolzhnost = result
End Function

Public Function SklonitVoennayaDolzhnost(dolzhnost As String) As String
    Dim result As String
    result = dolzhnost
    ' Hardcoded grammar rules for military positions
    result = Replace(result, "механик-радиотелефонист", "механику-радиотелефонисту")
    result = Replace(result, "разведчик-оператор", "разведчику-оператору")
    result = Replace(result, "командир ", "командиру ")
    result = Replace(result, "заместитель командира ", "заместителю командира ")
    result = Replace(result, "начальник ", "начальнику ")
    result = Replace(result, "заместитель начальника ", "заместителю начальника ")
    If Left(result, 8) = "старший " Then result = "старшему " & Mid(result, 9)
    If Left(result, 8) = "младший " Then result = "младшему " & Mid(result, 9)
    
    result = Replace(result, "механик", "механику")
    result = Replace(result, "радиотелефонист", "радиотелефонисту")
    result = Replace(result, "разведчик", "разведчику")
    result = Replace(result, "оператор", "оператору")
    result = Replace(result, "водитель", "водителю")
    result = Replace(result, "наводчик", "наводчику")
    result = Replace(result, "инструктор", "инструктору")
    result = Replace(result, "техник", "технику")
    SklonitVoennayaDolzhnost = result
End Function

Public Function SklonitFIO(fio As String) As String
    Dim parts() As String
    Dim familiya As String, imya As String, otchestvo As String, result As String, isWoman As Boolean
    parts = Split(Trim(fio), " ")
    If UBound(parts) >= 2 Then
        familiya = parts(0)
        imya = parts(1)
        otchestvo = parts(2)
        isWoman = (Right(LCase(otchestvo), 2) = "на")
        familiya = SklonitFamiliya(familiya, isWoman)
        imya = SklonitImya(imya, isWoman)
        otchestvo = SklonitOtchestvo(otchestvo, isWoman)
        result = familiya & " " & imya & " " & otchestvo
    Else
        result = fio
    End If
    SklonitFIO = result
End Function

Public Function SklonitFamiliya(familiya As String, isWoman As Boolean) As String
    Dim result As String
    result = familiya
    If isWoman Then
        If Right(familiya, 2) = "на" Then
            result = Left(familiya, Len(familiya) - 1) & "е"
        ElseIf Right(familiya, 1) = "а" Then
            result = Left(familiya, Len(familiya) - 1) & "е"
        ElseIf Right(familiya, 1) = "я" Then
            result = Left(familiya, Len(familiya) - 1) & "е"
        End If
    Else
        If Right(familiya, 2) = "ов" Or Right(familiya, 2) = "ёв" Or Right(familiya, 2) = "ин" Then
            result = familiya & "у"
        ElseIf Right(familiya, 2) = "ий" Then
            result = Left(familiya, Len(familiya) - 2) & "ому"
        ElseIf Right(familiya, 1) = "а" Then
            result = Left(familiya, Len(familiya) - 1) & "е"
        ElseIf Right(familiya, 1) = "я" Then
            result = Left(familiya, Len(familiya) - 1) & "е"
        End If
    End If
    SklonitFamiliya = result
End Function


Public Function SklonitImya(imya As String, isWoman As Boolean) As String
    Dim result As String
    result = imya
    If isWoman Then
        If Right(imya, 1) = "а" Then
            result = Left(imya, Len(imya) - 1) & "е"
        ElseIf Right(imya, 1) = "я" Then
            result = Left(imya, Len(imya) - 1) & "е"
        ElseIf Right(imya, 1) = "ь" Then
            result = Left(imya, Len(imya) - 1) & "и"
        ElseIf Right(imya, 1) = "и" Then
            result = imya
        End If
    Else
        If Right(imya, 1) = "р" Or Right(imya, 1) = "л" Or Right(imya, 1) = "н" Or Right(imya, 1) = "м" Then
            result = imya & "у"
        ElseIf Right(imya, 1) = "й" Then
            result = Left(imya, Len(imya) - 1) & "ю"
        ElseIf Right(imya, 1) = "а" Then
            result = Left(imya, Len(imya) - 1) & "е"
        ElseIf Right(imya, 1) = "я" Then
            result = Left(imya, Len(imya) - 1) & "е"
        End If
    End If
    SklonitImya = result
End Function


Public Function SklonitOtchestvo(otchestvo As String, isWoman As Boolean) As String
    Dim result As String
    result = otchestvo
    If isWoman Then
        If Right(otchestvo, 2) = "на" Then result = Left(otchestvo, Len(otchestvo) - 1) & "е"
    Else
        If Right(otchestvo, 2) = "ич" Then result = otchestvo & "у"
    End If
    SklonitOtchestvo = result
End Function

Public Function GetDolzhnostImenitelny(dolzhnost As String, VoinskayaChast As String) As String
    Dim keepWords As Variant
    keepWords = Array("роты", "взвода", "отделения", "расчета", "группы", "команды", "экипажа")
    Dim cutWords As Variant
    cutWords = Array("отдельного", "гвардейской", "общевойсковой", "мотострелковой", "танковой", "воздушно-десантной", "артиллерийской", "инженерной", "связи", "десантно-штурмовой", "батальона", "полка", "бригады", "дивизии", "корпуса", "армии", "округа")
    Dim dolzhnostLower As String
    Dim result As String
    Dim lastKeepPos As Long
    Dim lastKeepWord As String
    Dim i As Long
    Dim pos As Long
    dolzhnostLower = LCase(dolzhnost)
    lastKeepPos = -1
    lastKeepWord = ""
    For i = LBound(keepWords) To UBound(keepWords)
        pos = InStrRev(dolzhnostLower, keepWords(i))
        If pos > lastKeepPos Then
            lastKeepPos = pos
            lastKeepWord = keepWords(i)
        End If
    Next i
    If lastKeepPos > 0 Then
        Dim endKeepPos As Long
        endKeepPos = lastKeepPos + Len(lastKeepWord) - 1
        Dim cutPosition As Long
        cutPosition = 0
        For i = LBound(cutWords) To UBound(cutWords)
            pos = InStr(endKeepPos + 1, dolzhnostLower, cutWords(i))
            If pos > 0 Then
                If cutPosition = 0 Or pos < cutPosition Then
                    cutPosition = pos
                End If
            End If
        Next i
        If cutPosition > 0 Then
            Dim startCutPosition As Long
            startCutPosition = cutPosition
            For i = cutPosition - 1 To endKeepPos + 1 Step -1
                Dim char As String
                char = Mid(dolzhnostLower, i, 1)
                If IsNumeric(char) Or char = " " Then
                    startCutPosition = i
                Else
                    Exit For
                End If
            Next i
            result = LCase(Trim(Left(dolzhnost, startCutPosition - 1))) & " войсковой части " & VoinskayaChast
        Else
            result = LCase(Trim(Left(dolzhnost, endKeepPos))) & " войсковой части " & VoinskayaChast
        End If
    Else
        cutPosition = 0
        For i = LBound(cutWords) To UBound(cutWords)
            pos = InStr(dolzhnostLower, cutWords(i))
            If pos > 0 Then
                If cutPosition = 0 Or pos < cutPosition Then
                    cutPosition = pos
                End If
            End If
        Next i
        If cutPosition > 0 Then
            startCutPosition = cutPosition
            For i = cutPosition - 1 To 1 Step -1
                Dim char2 As String
                char2 = Mid(dolzhnostLower, i, 1)
                If IsNumeric(char2) Or char2 = " " Then
                    startCutPosition = i
                Else
                    Exit For
                End If
            Next i
            result = LCase(Trim(Left(dolzhnost, startCutPosition - 1))) & " войсковой части " & VoinskayaChast
        Else
            result = LCase(dolzhnost) & " войсковой части " & VoinskayaChast
        End If
    End If
    GetDolzhnostImenitelny = result
End Function

Public Function GetZvanieImenitelny(zvanie As String) As String
    Dim result As String
    Dim lowerZvanie As String
    lowerZvanie = LCase(Trim(zvanie))
    Select Case lowerZvanie
        Case "рядовой": result = "рядовой"
        Case "ефрейтор": result = "ефрейтор"
        Case "младший сержант": result = "младший сержант"
        Case "сержант": result = "сержант"
        Case "старший сержант": result = "старший сержант"
        Case "старшина": result = "старшина"
        Case "прапорщик": result = "прапорщик"
        Case "старший прапорщик": result = "старший прапорщик"
        Case "младший лейтенант": result = "младший лейтенант"
        Case "лейтенант": result = "лейтенант"
        Case "старший лейтенант": result = "старший лейтенант"
        Case "капитан": result = "капитан"
        Case "майор": result = "майор"
        Case "подполковник": result = "подполковник"
        Case "полковник": result = "полковник"
        Case "генерал-майор": result = "генерал-майор"
        Case "генерал-лейтенант": result = "генерал-лейтенант"
        Case "генерал-полковник": result = "генерал-полковник"
        Case "генерал армии": result = "генерал армии"
        Case Else: result = LCase(zvanie)
    End Select
    GetZvanieImenitelny = result
End Function

Public Function GetZvanieSkrasheno(zvanie As String) As String
    Dim result As String
    Dim lowerZvanie As String
    lowerZvanie = LCase(Trim(zvanie))
    Select Case lowerZvanie
        Case "рядовой": result = "рядовому"
        Case "ефрейтор": result = "ефрейтору"
        Case "младший сержант": result = "мл. сержанту"
        Case "сержант": result = "сержанту"
        Case "старший сержант": result = "ст. сержанту"
        Case "старшина": result = "старшине"
        Case "прапорщик": result = "прапорщику"
        Case "старший прапорщик": result = "ст. прапорщику"
        Case "младший лейтенант": result = "мл. лейтенанту"
        Case "лейтенант": result = "лейтенанту"
        Case "старший лейтенант": result = "ст. лейтенанту"
        Case "капитан": result = "капитану"
        Case "майор": result = "майору"
        Case "подполковник": result = "подполковнику"
        Case "полковник": result = "полковнику"
        Case "генерал-майор": result = "генерал-майору"
        Case "генерал-лейтенант": result = "генерал-лейтенанту"
        Case "генерал-полковник": result = "генерал-полковнику"
        Case "генерал армии": result = "генералу армии"
        Case Else: result = LCase(zvanie) & "у"
    End Select
    GetZvanieSkrasheno = result
End Function

Public Function GetMonthNameRussian(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetMonthNameRussian = "января"
        Case 2: GetMonthNameRussian = "февраля"
        Case 3: GetMonthNameRussian = "марта"
        Case 4: GetMonthNameRussian = "апреля"
        Case 5: GetMonthNameRussian = "мая"
        Case 6: GetMonthNameRussian = "июня"
        Case 7: GetMonthNameRussian = "июля"
        Case 8: GetMonthNameRussian = "августа"
        Case 9: GetMonthNameRussian = "сентября"
        Case 10: GetMonthNameRussian = "октября"
        Case 11: GetMonthNameRussian = "ноября"
        Case 12: GetMonthNameRussian = "декабря"
        Case Else: GetMonthNameRussian = "неизвестного месяца"
    End Select
End Function

Public Function GetFIOWithInitials(fio As String) As String
    Dim parts() As String
    Dim familiya As String, imya As String, otchestvo As String, result As String, isWoman As Boolean
    parts = Split(Trim(fio), " ")
    If UBound(parts) >= 2 Then
        familiya = parts(0)
        imya = parts(1)
        otchestvo = parts(2)
        isWoman = (Right(LCase(otchestvo), 2) = "на")
        Dim firstInitial As String, secondInitial As String
        firstInitial = UCase(Left(imya, 1))
        secondInitial = UCase(Left(otchestvo, 1))
        Dim declinedFamiliya As String
        declinedFamiliya = SklonitFamiliya(familiya, isWoman)
        result = firstInitial & "." & secondInitial & ". " & declinedFamiliya
    Else
        result = fio
    End If
    GetFIOWithInitials = result
End Function

Public Function GetFIOWithInitialsImenitelny(fio As String) As String
    Dim parts() As String
    Dim familiya As String, imya As String, otchestvo As String, result As String
    parts = Split(Trim(fio), " ")
    If UBound(parts) >= 2 Then
        familiya = parts(0)
        imya = parts(1)
        otchestvo = parts(2)
        Dim firstInitial As String, secondInitial As String
        firstInitial = UCase(Left(imya, 1))
        secondInitial = UCase(Left(otchestvo, 1))
        result = firstInitial & "." & secondInitial & ". " & familiya
    Else
        result = fio
    End If
    GetFIOWithInitialsImenitelny = result
End Function

Public Function GetZvanieImenitelnyForSignature(zvanie As String) As String
    Dim result As String
    Dim lowerZvanie As String
    lowerZvanie = LCase(Trim(zvanie))
    Select Case lowerZvanie
        Case "рядовой": result = "Рядовой"
        Case "ефрейтор": result = "Ефрейтор"
        Case "младший сержант": result = "Младший сержант"
        Case "сержант": result = "Сержант"
        Case "старший сержант": result = "Старший сержант"
        Case "старшина": result = "Старшина"
        Case "прапорщик": result = "Прапорщик"
        Case "старший прапорщик": result = "Старший прапорщик"
        Case "младший лейтенант": result = "Младший лейтенант"
        Case "лейтенант": result = "Лейтенант"
        Case "старший лейтенант": result = "Старший лейтенант"
        Case "капитан": result = "Капитан"
        Case "майор": result = "Майор"
        Case "подполковник": result = "Подполковник"
        Case "полковник": result = "Полковник"
        Case "генерал-майор": result = "Генерал-майор"
        Case "генерал-лейтенант": result = "Генерал-лейтенант"
        Case "генерал-полковник": result = "Генерал-полковник"
        Case "генерал армии": result = "Генерал армии"
        Case Else: result = UCase(Left(zvanie, 1)) & LCase(Mid(zvanie, 2))
    End Select
    GetZvanieImenitelnyForSignature = result
End Function

' ==========================================================
' GENERAL UTILITIES
' ==========================================================

' /**
'  * Finds a column index by header name (in the first row).
'  * @return Integer - Column index or -1 if not found.
'  */
Public Function FindColumn(ws As Worksheet, headerName As String) As Integer
    Dim i As Integer, lastCol As Integer
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For i = 1 To lastCol
        If Trim(ws.Cells(1, i).value) = Trim(headerName) Then
            FindColumn = i
            Exit Function
        End If
    Next i
    FindColumn = -1 ' Not found
End Function

' /**
'  * Retrieves all staff data (Dictionary) searching by ID or Name.
'  * Optimized to use Match instead of looping.
'  *
'  * @param queryValue String - Value to search
'  * @param byLichniyNomer Boolean - True=Search by ID, False=Search by Name
'  * @return Object (Dictionary)
'  */
Public Function GetStaffData(queryValue As String, Optional byLichniyNomer As Boolean = True) As Object

    Call EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim foundOk As Boolean
    Dim resultDict As Object
    Dim searchCol As Long
    Dim rowIndex As Long

    Set wsStaff = ThisWorkbook.Sheets("Штат")

    foundOk = FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast)
    If Not foundOk Then
        Set GetStaffData = CreateObject("Scripting.Dictionary")
        Exit Function
    End If

    Set resultDict = CreateObject("Scripting.Dictionary")

    ' Determine which column to search
    If byLichniyNomer Then
        searchCol = colLichniyNomer
    Else
        searchCol = colFIO
    End If

    ' Optimize: Use Match instead of Loop
    rowIndex = FindStaffRow(wsStaff, queryValue, searchCol)

    If rowIndex > 0 Then
        resultDict("Лицо") = wsStaff.Cells(rowIndex, colFIO).value
        resultDict("Личный номер") = wsStaff.Cells(rowIndex, colLichniyNomer).value
        resultDict("Воинское звание") = wsStaff.Cells(rowIndex, colZvanie).value
        resultDict("Часть") = wsStaff.Cells(rowIndex, colVoinskayaChast).value
        resultDict("Штатная должность") = wsStaff.Cells(rowIndex, colDolzhnost).value
        Set GetStaffData = resultDict
        Exit Function
    End If

    ' Not found
    MsgBox "Данные сотрудника не найдены по запросу: '" & queryValue & "'.", vbExclamation, "Ошибка поиска"
    Set GetStaffData = CreateObject("Scripting.Dictionary")
End Function

' /**
'  * Finds the "Table Number" column (Column "Name" with numeric values).
'  */
Public Function FindTableNumberColumn(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim lastCol As Long, i As Long
    Dim headerText As String
    Dim testValue As Variant
    Dim numericCount As Long, totalCount As Long
    
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    FindTableNumberColumn = 0
    
    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If headerText = "лицо" Then
            ' Check if column contains numeric values (first 20 rows)
            numericCount = 0
            totalCount = 0
            Dim lastRow As Long, j As Long
            lastRow = ws.Cells(ws.Rows.count, i).End(xlUp).Row
            If lastRow > 1 Then
                For j = 2 To Application.WorksheetFunction.Min(lastRow, 20)
                    testValue = ws.Cells(j, i).value
                    If Not IsEmpty(testValue) Then
                        totalCount = totalCount + 1
                        If IsNumeric(testValue) Then
                            numericCount = numericCount + 1
                        End If
                    End If
                Next j
                ' If >50% numeric, assume it's Table Number
                If totalCount > 0 And numericCount > totalCount / 2 Then
                    FindTableNumberColumn = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    Exit Function
ErrorHandler:
    FindTableNumberColumn = 0
End Function

' /**
'  * Searches staff by Table Number.
'  */
Public Function GetStaffDataByTableNumber(tableNumber As String) As Object
    On Error GoTo ErrorHandler
    
    Call EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet
    Dim colTableNumber As Long
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim resultDict As Object
    Dim rowIndex As Long
    Dim foundOk As Boolean
    
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    foundOk = FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast)
    If Not foundOk Then
        Set GetStaffDataByTableNumber = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    colTableNumber = FindTableNumberColumn(wsStaff)
    If colTableNumber = 0 Then
        Set GetStaffDataByTableNumber = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    ' Optimize: Use Match
    rowIndex = FindStaffRow(wsStaff, tableNumber, colTableNumber)
    
    If rowIndex > 0 Then
        resultDict("Лицо") = wsStaff.Cells(rowIndex, colFIO).value
        resultDict("Личный номер") = wsStaff.Cells(rowIndex, colLichniyNomer).value
        resultDict("Воинское звание") = wsStaff.Cells(rowIndex, colZvanie).value
        resultDict("Часть") = wsStaff.Cells(rowIndex, colVoinskayaChast).value
        resultDict("Штатная должность") = wsStaff.Cells(rowIndex, colDolzhnost).value
        Set GetStaffDataByTableNumber = resultDict
        Exit Function
    End If
    
    Set GetStaffDataByTableNumber = CreateObject("Scripting.Dictionary")
    Exit Function
    
ErrorHandler:
    Set GetStaffDataByTableNumber = CreateObject("Scripting.Dictionary")
End Function

' /**
'  * Universal Search: First by ID, then by Table Number.
'  */
Public Function FindEmployeeByAnyNumber(number As String) As Object
    On Error GoTo ErrorHandler
    
    Dim staffData As Object
    Dim numberTrimmed As String
    
    numberTrimmed = Trim(number)
    If numberTrimmed = "" Then
        Set FindEmployeeByAnyNumber = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    
    ' 1. Try by Personal ID
    Set staffData = GetStaffData(numberTrimmed, True)
    If staffData.count > 0 Then
        Set FindEmployeeByAnyNumber = staffData
        Exit Function
    End If
    
    ' 2. Try by Table Number
    Set staffData = GetStaffDataByTableNumber(numberTrimmed)
    If staffData.count > 0 Then
        Set FindEmployeeByAnyNumber = staffData
        Exit Function
    End If
    
    ' Not found
    Set FindEmployeeByAnyNumber = CreateObject("Scripting.Dictionary")
    Exit Function
    
ErrorHandler:
    Set FindEmployeeByAnyNumber = CreateObject("Scripting.Dictionary")
End Function

' /**
'  * Ensures global column indexes are set.
'  */
Public Sub EnsureStaffColumnsInitialized()
    If colLichniyNomer_Global = 0 Or colFIO_Global = 0 Then
        InitStaffColumnIndexes
    End If
End Sub

' ==========================================================
' EXCEL 2010/2016 COMPATIBILITY
' ==========================================================

' /**
'  * Saves Word Document safely (handles diff versions).
'  */
Public Sub SaveWordDocumentSafe(wdDoc As Object, filePath As String)
    On Error Resume Next
    
    ' Try SaveAs2 (Word 2010+)
    wdDoc.SaveAs2 filePath
    
    ' Fallback to SaveAs
    If Err.number <> 0 Then
        Err.Clear
        On Error Resume Next
        Dim fileFormat As Long
        ' Determine format code
        If Right(LCase(filePath), 5) = ".docx" Then
            fileFormat = 16 ' wdFormatXMLDocument
        ElseIf Right(LCase(filePath), 4) = ".doc" Then
            fileFormat = 0 ' wdFormatDocument
        Else
            fileFormat = 16
        End If
        
        wdDoc.SaveAs filePath, fileFormat
    End If
    
    On Error GoTo 0
End Sub

' /**
'  * Checks minimal Excel version (2010+ required).
'  */
Public Function CheckExcelVersion() As Boolean
    Dim version As String
    version = Application.version
    
    Dim majorVersion As Integer
    Dim dotPos As Integer
    dotPos = InStr(version, ".")
    
    If dotPos > 0 Then
        majorVersion = CInt(Left(version, dotPos - 1))
    Else
        majorVersion = CInt(version)
    End If
    
    If majorVersion < 14 Then
        MsgBox "Требуется Microsoft Excel 2010 или выше. " & _
               "Текущая версия: " & version, vbCritical
        CheckExcelVersion = False
    Else
        CheckExcelVersion = True
    End If
End Function

' /**
'  * Safely creates or gets Word Application instance.
'  */
Public Function CreateWordAppSafely() As Object
    Dim wdApp As Object
    
    On Error Resume Next
    ' Try to get existing instance
    Set wdApp = GetObject(, "Word.Application")
    
    If wdApp Is Nothing Then
        ' Create new instance
        Set wdApp = CreateObject("Word.Application")
        
        If wdApp Is Nothing Then
            MsgBox "Не удалось создать экземпляр Microsoft Word. " & _
                   "Убедитесь, что Word установлен и доступен.", vbCritical, "Ошибка Word"
            Set CreateWordAppSafely = Nothing
            Exit Function
        End If
    End If
    
    On Error GoTo 0
    Set CreateWordAppSafely = wdApp
End Function

' /**
'  * checks if Word is available on the system.
'  */
Public Function IsWordAvailable() As Boolean
    On Error Resume Next
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")
    If Err.number = 0 And Not wdApp Is Nothing Then
        wdApp.Quit False
        IsWordAvailable = True
    Else
        IsWordAvailable = False
    End If
    On Error GoTo 0
End Function

