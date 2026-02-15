Attribute VB_Name = "mdlHelper"
' ==============================================================================
' Module: mdlHelper
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Date: 14.02.2026 (Final Fix)
' Description: Universal utility functions and procedures.
'              Fixed: Restored missing GetDolzhnostImenitelny function.
' ==============================================================================

Option Explicit

Public colFIO_Global As Long
Public colLichniyNomer_Global As Long
Public colZvanie_Global As Long
Public colDolzhnost_Global As Long
Public colVoinskayaChast_Global As Long

Public Sub InitStaffColumnIndexes()
    Dim wsStaff As Worksheet
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    If Not FindColumnNumbers(wsStaff, colLichniyNomer_Global, colZvanie_Global, colFIO_Global, colDolzhnost_Global, colVoinskayaChast_Global) Then
        MsgBox "Корректные индексы столбцов не удалось определить. Работа программы невозможна.", vbCritical
        End
    End If
End Sub

Public Function FindStaffRow(ws As Worksheet, lichniyNomer As String, colNum As Long) As Long
    Dim res As Variant
    res = Application.Match(lichniyNomer, ws.Columns(colNum), 0)
    If isError(res) Then
        FindStaffRow = 0
    Else
        FindStaffRow = CLng(res)
    End If
End Function

Public Function FindColumnNumbers(ws As Worksheet, ByRef colLichniyNomer As Long, ByRef colZvanie As Long, ByRef colFIO As Long, ByRef colDolzhnost As Long, ByRef colVoinskayaChast As Long) As Boolean
    Dim lastCol As Long, i As Long, headerText As String
    Dim foundFIO As Boolean, foundDolzhnost As Boolean
    Dim msgErr As String

    colLichniyNomer = 0: colZvanie = 0: colFIO = 0: colDolzhnost = 0: colVoinskayaChast = 0
    foundFIO = False: foundDolzhnost = False
    msgErr = ""

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "личный номер") > 0 Then
            colLichniyNomer = i
            Exit For
        End If
    Next i
    If colLichniyNomer = 0 Then msgErr = msgErr & "Не найден столбец 'Личный номер'." & vbCrLf

    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "воинское звание") > 0 Then
            colZvanie = i
            Exit For
        End If
    Next i
    If colZvanie = 0 Then msgErr = msgErr & "Не найден столбец 'Воинское звание'." & vbCrLf

    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        If InStr(headerText, "часть") > 0 Or InStr(headerText, "раздел персонала") > 0 Then
            colVoinskayaChast = i
            Exit For
        End If
    Next i
    If colVoinskayaChast = 0 Then msgErr = msgErr & "Не найден столбец 'Часть' или 'Раздел персонала'." & vbCrLf

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

    If colLichniyNomer > 0 And colZvanie > 0 And colFIO > 0 And colDolzhnost > 0 And colVoinskayaChast > 0 Then
        FindColumnNumbers = True
    Else
        FindColumnNumbers = False
        MsgBox "Ошибка при определении столбцов на листе 'Штат':" & vbCrLf & msgErr, vbCritical, "Ошибка структуры"
    End If
End Function

Private Function IsTextFIOColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim lastRow As Long, i As Long, value As String
    Dim textCount As Long, totalCount As Long
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
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
    If totalCount > 0 Then IsTextFIOColumn = (textCount / totalCount) > 0.7 Else IsTextFIOColumn = False
End Function

Public Function ExtractVoinskayaChast(inputText As String) As String
    Dim Text As String, i As Long, result As String, inNumber As Boolean
    Text = Trim(inputText): result = "": inNumber = False
    For i = 1 To Len(Text)
        If IsNumeric(Mid(Text, i, 1)) Then
            result = result & Mid(Text, i, 1): inNumber = True
        Else
            If inNumber And Len(result) >= 4 Then ExtractVoinskayaChast = result: Exit Function
            If inNumber Then result = "": inNumber = False
        End If
    Next i
    If Len(result) >= 4 Then ExtractVoinskayaChast = result Else ExtractVoinskayaChast = inputText
End Function

Public Function IsPeriodActual(dateEnd As Date) As Boolean
    IsPeriodActual = (dateEnd >= GetExportCutoffDate())
End Function

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

Public Function ContainsNumbers(Text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        If char >= "0" And char <= "9" Then ContainsNumbers = True: Exit Function
    Next i
    ContainsNumbers = False
End Function

Public Function ContainsLetters(Text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        If (char >= "А" And char <= "я") Or (char >= "A" And char <= "z") Then
            ContainsLetters = True: Exit Function
        End If
    Next i
    ContainsLetters = False
End Function



' /**
'  * Collects all period pairs (Start/End/Days) for a person from "DSO" sheet.
'  * Uses ParseDateSafe to handle text dates.
'  */
Public Sub CollectAllPersonPeriods(ws As Worksheet, rowNum As Long, periods As Collection)
    Dim lastCol As Long, j As Long, dateStart As Date, dateEnd As Date
    On Error GoTo ErrorHandler
    
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    j = 5 ' Start from column E
    
    Do While j + 1 <= lastCol
        ' Attempt to parse both values using the robust parser
        dateStart = ParseDateSafe(ws.Cells(rowNum, j).value)
        dateEnd = ParseDateSafe(ws.Cells(rowNum, j + 1).value)
        
        ' Check if both are valid (> year 2000)
        If dateStart > 0 And dateEnd > 0 Then
            Dim daysCount As Long
            daysCount = DateDiff("d", dateStart, dateEnd) + 1
            
            Dim newPeriod As Collection: Set newPeriod = New Collection
            newPeriod.Add dateStart
            newPeriod.Add dateEnd
            newPeriod.Add daysCount
            periods.Add newPeriod
        End If
        j = j + 2
    Loop
    Exit Sub
ErrorHandler:
    Resume Next
End Sub

Public Function HasInvalidPair(periods As Collection) As Boolean
    Dim p As Collection
    For Each p In periods
        If p(2) < p(1) Then HasInvalidPair = True: Exit Function
    Next p
    HasInvalidPair = False
End Function

Public Function SortPeriodsByDateStart(periods As Collection) As Collection
    Dim arr() As Variant, i As Long, j As Long, n As Long
    n = periods.count
    If n <= 1 Then Set SortPeriodsByDateStart = periods: Exit Function
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = periods(i): Next i
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

Public Function hasCriticalErrors() As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("ДСО")
    If ws Is Nothing Then hasCriticalErrors = True: Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim i As Long, j As Long
    For i = 2 To lastRow
        For j = 5 To lastCol Step 2
            If ws.Cells(i, j).Interior.Color = RGB(255, 100, 100) Or ws.Cells(i, j).Interior.Color = RGB(255, 200, 200) Then
                hasCriticalErrors = True: Exit Function
            End If
        Next j
    Next i
    hasCriticalErrors = False
End Function

Public Function CheckRowForDateErrors(ws As Worksheet, rowNum As Long) As Boolean
    Dim lastCol As Long, j As Long, startValue As String, endValue As String
    Dim dateStart As Date, dateEnd As Date, hasErrors As Boolean
    On Error GoTo ErrorHandler
    hasErrors = False
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol > 50 Then lastCol = 50
    j = 5
    Do While j + 1 <= lastCol
        startValue = Trim(ws.Cells(rowNum, j).Text)
        endValue = Trim(ws.Cells(rowNum, j + 1).Text)
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

' ================= GRAMMAR FUNCTIONS =================
Public Function SklonitZvanie(zvanie As String) As String
    Dim result As String, lowerZvanie As String
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
        Case "генерал-лейтенант": result = "Генерал-лейтенант"
        Case "генерал-полковник": result = "Генерал-полковнику"
        Case "генерал армии": result = "Генералу армии"
        Case Else: result = UCase(Left(zvanie, 1)) & LCase(Mid(zvanie, 2))
    End Select
    SklonitZvanie = result
End Function

Public Function SklonitDolzhnost(dolzhnost As String, VoinskayaChast As String) As String
    Dim keepWords As Variant, cutWords As Variant
    ' Слова, после которых мы ожидаем продолжение (обычно родительный падеж подразделения)
    keepWords = Array("роты", "взвода", "отделения", "расчета", "группы", "команды", "экипажа", "службы", "батареи")
    ' Слова, с которых начинается "хвост" (название вышестоящего подразделения, которое мы хотим заменить на в/ч)
    cutWords = Array("отдельного", "гвардейской", "общевойсковой", "мотострелковой", "танковой", "воздушно-десантной", "артиллерийской", "инженерной", "связи", "десантно-штурмовой", "батальона", "полка", "бригады", "дивизии", "корпуса", "армии", "округа")
    
    Dim dolzhnostLower As String, result As String
    Dim lastKeepPos As Long, lastKeepWord As String
    Dim i As Long, pos As Long
    
    dolzhnostLower = LCase(dolzhnost)
    lastKeepPos = -1
    lastKeepWord = ""
    
    ' 1. Ищем самое "глубокое" ключевое слово подразделения (отделения, группы и т.д.)
    For i = LBound(keepWords) To UBound(keepWords)
        pos = InStrRev(dolzhnostLower, keepWords(i))
        If pos > lastKeepPos Then
            lastKeepPos = pos
            lastKeepWord = keepWords(i)
        End If
    Next i
    
    Dim startCutPosition As Long
    startCutPosition = 0
    Dim foundCut As Boolean
    foundCut = False
    
    ' Определение зоны поиска "хвоста"
    Dim searchStartPos As Long
    If lastKeepPos > 0 Then
        searchStartPos = lastKeepPos + Len(lastKeepWord)
    Else
        searchStartPos = 1
    End If
    
    ' 2. Ищем слово, с которого нужно "резать" (отдельного, батальона и т.д.)
    Dim cutPosition As Long
    cutPosition = 0
    
    For i = LBound(cutWords) To UBound(cutWords)
        pos = InStr(searchStartPos, dolzhnostLower, cutWords(i))
        If pos > 0 Then
            ' Берем самое первое найденное слово-разделитель после keepWord
            If cutPosition = 0 Or pos < cutPosition Then
                cutPosition = pos
            End If
        End If
    Next i
    
    ' 3. Если нашли место разреза, ищем начало номера перед ним (например, "111 отдельного")
    If cutPosition > 0 Then
        ' Идем назад от слова-разделителя, пока видим цифры или пробелы
        startCutPosition = cutPosition
        Dim j As Long
        For j = cutPosition - 1 To searchStartPos Step -1
            Dim char As String
            char = Mid(dolzhnostLower, j, 1)
            ' Если это цифра или пробел - сдвигаем точку разреза влево
            If IsNumeric(char) Or char = " " Then
                startCutPosition = j
            Else
                ' Наткнулись на букву или скобку - стоп
                Exit For
            End If
        Next j
        foundCut = True
    End If
    
    ' 4. Формируем результат
    If foundCut And startCutPosition > 2 Then
        ' Безопасная обрезка: берем левую часть
        result = LCase(Trim(Left(dolzhnost, startCutPosition - 1)))
    Else
        ' Если не нашли, где резать, или обрезка слишком агрессивная - берем всё
        result = LCase(Trim(dolzhnost))
    End If
    
    ' Добавляем номер части
    result = result & " войсковой части " & VoinskayaChast
    
    ' 5. Склоняем саму должность (Начальник -> Начальнику)
    result = SklonitVoennayaDolzhnost(result)
    
    SklonitDolzhnost = result
End Function

Public Function SklonitVoennayaDolzhnost(dolzhnost As String) As String
    Dim result As String
    result = dolzhnost ' Входная строка уже в нижнем регистре от предыдущей функции
    
    ' Сложные названия
    result = Replace(result, "механик-радиотелефонист", "механику-радиотелефонисту")
    result = Replace(result, "разведчик-оператор", "разведчику-оператору")
    result = Replace(result, "наводчик-оператор", "наводчику-оператору")
    result = Replace(result, "механик-водитель", "механику-водителю")
    result = Replace(result, "заместитель командира", "заместителю командира")
    result = Replace(result, "заместитель начальника", "заместителю начальника")
    result = Replace(result, "помощник начальника", "помощнику начальника")
    
    ' Префиксы
    If Left(result, 8) = "старший " Then result = "старшему " & Mid(result, 9)
    If Left(result, 8) = "младший " Then result = "младшему " & Mid(result, 9)
    If Left(result, 8) = "главный " Then result = "главному " & Mid(result, 9)
    
    ' Одиночные должности
    result = Replace(result, "командир ", "командиру ")
    result = Replace(result, "начальник ", "начальнику ")
    result = Replace(result, "механик", "механику")
    result = Replace(result, "радиотелефонист", "радиотелефонисту")
    result = Replace(result, "разведчик", "разведчику")
    result = Replace(result, "оператор", "оператору")
    result = Replace(result, "водитель", "водителю")
    result = Replace(result, "наводчик", "наводчику")
    result = Replace(result, "инструктор", "инструктору")
    result = Replace(result, "техник", "технику")
    result = Replace(result, "пулеметчик", "пулеметчику")
    result = Replace(result, "гранатометчик", "гранатометчику")
    result = Replace(result, "стрелок", "стрелку")
    result = Replace(result, "сапер", "саперу")
    
    SklonitVoennayaDolzhnost = result
End Function

Public Function SklonitFIO(sName As String) As String
    ' Обертка для новой функции FIO (запрашиваем Дательный падеж "Д" - Кому?)
    ' Аргумент переименован в sName, чтобы не конфликтовать с именем функции FIO
    SklonitFIO = fio(sName, "Д")
End Function

' ==========================================================
' ФУНКЦИИ СКЛОНЕНИЯ (ИСПРАВЛЕННЫЕ)
' ==========================================================

' ==========================================================
' ФУНКЦИИ СКЛОНЕНИЯ (ИСПРАВЛЕННЫЕ)
' ==========================================================

'Public Function SklonitFamiliya(familiya As String, IsWoman As Boolean) As String
'    Dim result As String
'    result = familiya
'    If IsWoman Then
'        ' Женские фамилии
'        If Right(familiya, 2) = "на" Then
'            result = Left(familiya, Len(familiya) - 1) & "е"
'        ElseIf Right(familiya, 1) = "а" Then
'            result = Left(familiya, Len(familiya) - 1) & "е"
'        ElseIf Right(familiya, 1) = "я" Then
'            result = Left(familiya, Len(familiya) - 1) & "е"
'        End If
'    Else
'        ' Мужские фамилии
'        ' ДОБАВЛЕНО: "ев", "ын"
'        If Right(familiya, 2) = "ов" Or Right(familiya, 2) = "ёв" Or Right(familiya, 2) = "ин" Or Right(familiya, 2) = "ев" Or Right(familiya, 2) = "ын" Then
'            result = familiya & "у"
'        ElseIf Right(familiya, 2) = "ий" Then
'            result = Left(familiya, Len(familiya) - 2) & "ому"
'        ElseIf Right(familiya, 2) = "ая" Then ' Например, Белая -> Белому (редко, но бывает у мужчин)
'             result = Left(familiya, Len(familiya) - 2) & "ому"
'        ElseIf Right(familiya, 1) = "а" Then
'            result = Left(familiya, Len(familiya) - 1) & "е"
'        ElseIf Right(familiya, 1) = "я" Then
'            result = Left(familiya, Len(familiya) - 1) & "е"
'        End If
'    End If
'    SklonitFamiliya = result
'End Function

'Public Function SklonitImya(imya As String, IsWoman As Boolean) As String
'    Dim result As String
'    result = imya
'    If IsWoman Then
'        If Right(imya, 1) = "а" Then
'            result = Left(imya, Len(imya) - 1) & "е"
'        ElseIf Right(imya, 1) = "я" Then
'            result = Left(imya, Len(imya) - 1) & "е"
'        ElseIf Right(imya, 1) = "ь" Then
'            result = Left(imya, Len(imya) - 1) & "и" ' Любовь -> Любови
'        ElseIf Right(imya, 1) = "и" Then
'            result = imya
'        End If
'    Else
'        ' Мужские имена
'        If Right(imya, 1) = "р" Or Right(imya, 1) = "л" Or Right(imya, 1) = "н" Or Right(imya, 1) = "м" Or Right(imya, 1) = "б" Or Right(imya, 1) = "г" Or Right(imya, 1) = "к" Or Right(imya, 1) = "с" Then
'            result = imya & "у"
'        ElseIf Right(imya, 1) = "й" Then
'            result = Left(imya, Len(imya) - 1) & "ю"
'        ElseIf Right(imya, 1) = "а" Then
'            result = Left(imya, Len(imya) - 1) & "е"
'        ElseIf Right(imya, 1) = "я" Then
'            result = Left(imya, Len(imya) - 1) & "е"
'        ' ДОБАВЛЕНО: Обработка мягкого знака для мужских имен (Игорь -> Игорю)
'        ElseIf Right(imya, 1) = "ь" Then
'            result = Left(imya, Len(imya) - 1) & "ю"
'        End If
'    End If
'    SklonitImya = result
'End Function


'Public Function SklonitOtchestvo(otchestvo As String, IsWoman As Boolean) As String
'    Dim result As String
'    result = otchestvo
'    If IsWoman Then
'        If Right(otchestvo, 2) = "на" Then result = Left(otchestvo, Len(otchestvo) - 1) & "е"
'    Else
'        If Right(otchestvo, 2) = "ич" Then result = otchestvo & "у"
'    End If
'    SklonitOtchestvo = result
'End Function

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

Public Function GetFIOWithInitials(sName As String) As String
    ' Возвращает "И.О. Фамилия" в Дательном падеже (Кому?)
    Dim sResult As String
    Dim parts() As String
    
    ' Используем sName вместо fio
    sResult = fio(sName, "Д", True)
    
    parts = Split(sResult, " ")
    If UBound(parts) = 1 Then
        GetFIOWithInitials = parts(1) & " " & parts(0)
    Else
        GetFIOWithInitials = sResult
    End If
End Function

Public Function GetFIOWithInitialsImenitelny(sName As String) As String
    ' Возвращает "И.О. Фамилия" в Именительном падеже (Кто?)
    Dim sResult As String
    Dim parts() As String
    
    ' Используем sName вместо fio
    sResult = fio(sName, "И", True)
    
    parts = Split(sResult, " ")
    If UBound(parts) = 1 Then
        GetFIOWithInitialsImenitelny = parts(1) & " " & parts(0)
    Else
        GetFIOWithInitialsImenitelny = sResult
    End If
End Function

Public Function GetZvanieImenitelnyForSignature(zvanie As String) As String
    Dim result As String, lowerZvanie As String
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

' ==========================================================
' EXPORT REFACTOR HELPERS (Step 1)
' ==========================================================

' /**
'  * Gets the last used row number in a specific column.
'  * @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'  * @param ws [Worksheet] Target sheet.
'  * @param colIndex [Variant] Column index (Long) or column letter (String), e.g. "C".
'  * @return [Long] Last row with data, or 0 if column empty.
'  */
Public Function GetLastRow(ws As Worksheet, colIndex As Variant) As Long
    Dim colNum As Long
    On Error GoTo ErrorHandler

    If VarType(colIndex) = vbString Then
        colNum = ws.Range(CStr(colIndex) & "1").Column
    Else
        colNum = CLng(colIndex)
    End If
    GetLastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    Exit Function
ErrorHandler:
    GetLastRow = 0
End Function

' /**
'  * Gets existing Word Application or creates new instance. Sets wasNotRunning if new instance created.
'  * @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'  * @param wasNotRunning [Boolean] ByRef. Set to True if a new instance was created.
'  * @return [Object] Word.Application. Nothing if creation failed. Visible = False by default.
'  */
Public Function GetOrCreateWordApp(ByRef wasNotRunning As Boolean) As Object
    Dim wdApp As Object
    wasNotRunning = False

    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    On Error GoTo 0

    If wdApp Is Nothing Then
        On Error Resume Next
        Set wdApp = CreateObject("Word.Application")
        On Error GoTo 0
        If wdApp Is Nothing Then
            MsgBox "Не удалось создать экземпляр Microsoft Word. Убедитесь, что Word установлен.", vbCritical, "Ошибка Word"
            Set GetOrCreateWordApp = Nothing
            Exit Function
        End If
        wasNotRunning = True
    End If
    wdApp.Visible = False
    Set GetOrCreateWordApp = wdApp
End Function

' /**
'  * Terminates any remaining WINWORD.EXE processes (zombie cleanup) via WMI.
'  * @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'  */
Public Sub KillZombieWordProcesses()
    Dim wmi As Object, procs As Object, proc As Object
    On Error Resume Next
    Set wmi = GetObject("winmgmts:")
    If wmi Is Nothing Then Exit Sub
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='WINWORD.EXE'")
    If procs Is Nothing Then Exit Sub
    For Each proc In procs
        proc.Terminate
    Next proc
    On Error GoTo 0
End Sub

' /**
'  * Closes and releases Word Application. If wasNotRunning, quits app and runs zombie cleanup.
'  * @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'  * @param wdApp [Object] ByRef. Word.Application reference. Set to Nothing on exit.
'  * @param wasNotRunning [Boolean] True if this instance was created by GetOrCreateWordApp.
'  */
Public Sub CloseWordAppSafe(ByRef wdApp As Object, ByVal wasNotRunning As Boolean)
    On Error Resume Next
    If Not wdApp Is Nothing Then
        If wasNotRunning Then
            wdApp.Quit False
            KillZombieWordProcesses
        End If
        Set wdApp = Nothing
    End If
    On Error GoTo 0
End Sub

' /**
'  * Builds formatted period block for Word: "- from dd.mm.yy to dd.mm.yy (X days)...", (NOT ACTUAL) if past cutoff.
'  * Validates periods (End >= Start), sorts by start date, computes total and rest days.
'  * @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'  * @param periods [Collection] Collection of period Collections from CollectAllPersonPeriods: (1)=Start, (2)=End, (3)=Days.
'  * @param cutoffDate [Date] Periods ending before this are marked as not actual.
'  * @param outTotalDays [Long] ByRef. Sum of days of all periods.
'  * @param outRestDays [Long] ByRef. outTotalDays \ 3 * 2 (rest days formula).
'  * @return [String] Formatted text block; empty string and outTotalDays=0, outRestDays=0 if invalid pair. Raises on invalid.
'  */
Public Function FormatPeriodsForWord(periods As Collection, cutoffDate As Date, ByRef outTotalDays As Long, ByRef outRestDays As Long) As String
    Dim sorted As Collection, p As Collection, i As Long, line As String
    Dim dateStart As Date, dateEnd As Date, daysCount As Long

    outTotalDays = 0
    outRestDays = 0
    FormatPeriodsForWord = ""

    On Error GoTo ErrorHandler

    If periods Is Nothing Or periods.count = 0 Then Exit Function

    If HasInvalidPair(periods) Then
        Err.Raise 5, "FormatPeriodsForWord", "Invalid period: end date is less than start date."
    End If

    Set sorted = SortPeriodsByDateStart(periods)

    For i = 1 To sorted.count
        Set p = sorted(i)
        dateStart = p(1)
        dateEnd = p(2)
        daysCount = p(3)
        outTotalDays = outTotalDays + daysCount

        line = "- с " & Format(dateStart, "dd.mm.yy") & " по " & Format(dateEnd, "dd.mm.yy") & " (" & daysCount & " сут.)"
        If dateEnd < cutoffDate Then
            line = line & " (НЕ АКТУАЛЕН — старше 3 лет + 1 месяц!)"
        End If
        FormatPeriodsForWord = FormatPeriodsForWord & line & vbCrLf
    Next i

    outRestDays = (outTotalDays \ 3) * 2
    Exit Function
ErrorHandler:
    outTotalDays = 0
    outRestDays = 0
    FormatPeriodsForWord = ""
    Err.Raise Err.number, Err.Source, Err.Description
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

' ==============================================================================
' HOTFIX: ROBUST DATE PARSER
' ==============================================================================
Public Function ParseDateSafe(val As Variant) As Date
    On Error Resume Next
    ParseDateSafe = 0 ' Default failure
    
    If IsEmpty(val) Then Exit Function
    
    ' 1. Try standard conversion (works for real dates and proper system text)
    If IsDate(val) Then
        Dim d As Date
        d = CDate(val)
        ' Logic check: ignore default 1899 dates (value 0)
        If Year(d) > 2000 Then
            ParseDateSafe = d
            Exit Function
        End If
    End If
    
    ' 2. Manual parsing for stubborn text formats (e.g. "01.02.25" on US system)
    Dim sVal As String
    sVal = Trim(CStr(val))
    
    Dim parts() As String
    If InStr(sVal, ".") > 0 Then
        parts = Split(sVal, ".")
    ElseIf InStr(sVal, "/") > 0 Then
        parts = Split(sVal, "/")
    End If
    
    ' Expecting 3 parts: Day, Month, Year
    If (Not Not parts) <> 0 Then
        If UBound(parts) = 2 Then
            Dim dInt As Integer, mInt As Integer, yInt As Integer
            dInt = CInt(parts(0))
            mInt = CInt(parts(1))
            yInt = CInt(parts(2))
            
            ' Fix 2-digit year
            If yInt < 100 Then
                If yInt < 50 Then yInt = 2000 + yInt Else yInt = 1900 + yInt
            End If
            
            ' Logic check: Year > 2000
            If yInt > 2000 Then
                ParseDateSafe = DateSerial(yInt, mInt, dInt)
            End If
        End If
    End If
End Function

' ==========================================================
' НОВЫЙ БЛОК СКЛОНЕНИЯ (UDFs_FIO Engine)
' ==========================================================

Private Function IsMan(ByVal sName As String) As Boolean
    Dim arMenNames As Variant
    Dim i As Long
    
    arMenNames = Array("Абай", "Абрам", "Абраам", "Аваз", "Авазбек", "Авдей", "Адилет", "Адольф", "Азамат", "Акбар", "Аксентий", "Агафон", "Айбек", "Айрат", "Алдар", "Алишер", "Алан", "Александр", "Алексей", "Али", "Алмат", "Альберт", "Альвиан", "Альфред", "Анатолий", "Андрей", "Антон", "Антонин", "Аристарх", "Аркадий", "Армен", "Арнольд", "Арон", "Арсен", "Арсений", "Артем", "Артём", "Артемий", _
        "Артур", "Аскольд", "Афанасий", "Ашот", "Батыр", "Бауыржан", "Богдан", "Борис", "Вадим", "Валентин", "Валерий", "Валерьян", "Варлам", "Василий", "Вахтанг", "Венедикт", "Вениамин", "Виктор", "Виталий", "Влад", "Владилен", "Владимир", "Владислав", "Владлен", "Вольф", "Всеволод", "Вячеслав", "Гавриил", "Гаврил", "Гайдар", _
        "Геласий", "Геннадий", "Генрих", "Георгий", "Герасим", "Герман", "Глеб", "Гордей", "Григорий", "Гурген", "Давид", "Дамир", "Даниил", "Данил", "Данияр", "Дастан", "Демьян", "Денис", "Диас", "Динишбек", "Дмитрий", "Дорофей", "Евгений", "Евграф", "Евдоким", "Евсей", "Егор", _
        "Еремей", "Ернар", "Ермолай", "Ефим", "Жонибек", "Заур", "Зиновий", "Иакинф", "Иван", "Игнатий", "Игнат", "Игорь", "Иларион", "Илларион", "Ильдар", "Ильшат", "Илья", "Иннокентий", "Иосиф", "Ипполит", "Ирек", "Ириней", "Исидор", "Исаак", "Исхак", "Иулиан", "Казимир", "Кайрат", "Камиль", "Карл", "Касьян", "Керим", "Ким", "Кирилл", "Клавдий", "Кондрат", "Константин", _
        "Кристиан", "Кузьма", "Куприян", "Лаврентий", "Лев", "Ленар", "Леонард", "Леонид", "Леонтий", "Лука", "Лукий", "Лукьян", "Людвиг", "Магомед", "Магомет", "Майк", "Макар", "Максат", "Макс", "Максим", "Марат", "Марк", "Мартын", "Матвей", "Махач", "Махмуд", "Мелентий", "Мирлан", "Мирослав", _
        "Митрофан", "Михаил", "Модест", "Моисей", "Мстислав", "Мурад", "Мухамед", "Мухаммед", "Муса", "Мэлор", "Наум", "Никита", "Никифор", "Николай", "Нурбек", "Нуржан", "Нурлан", "Олег", "Онисим", "Осип", "Отар", "Павел", "Пантелеймон", "Парфений", "Пётр", "Петр", "Платон", "Порфирий", "Прокопий", "Протасий", "Прохор", "Радомир", "Разумник", "Рамазан", "Рамзан", "Рафаэль", _
        "Рафик", "Ринат", "Роман", "Роберт", "Ростислав", "Рубен", "Рудольф", "Руслан", "Рустам", "Рустем", "Сабир", "Савва", "Савелий", "Святослав", "Семён", "Семен", "Серафим", "Сергей", "Серик", "Созон", "Соломон", "Спиридон", "Станислав", "Степан", "Султан", "Тагир", "Тарас", "Темир", "Темирхан", "Тигран", "Тимофей", "Тимур", "Тихон", "Трифон", _
        "Трофим", "Фадей", "Фаддей", "Федор", "Фёдор", "Федосей", "Федот", "Феликс", "Филат", "Филипп", "Фома", "Фрол", "Харитон", "Хафиз", "Христофор", "Чеслав", "Шамиль", "Шамхал", "Эдуард", "Эльдар", "Эльман", "Эмиль", "Эммануил", "Эраст", "Юлиан", "Юлиус", "Юлий", "Юрий", "Юстин", "Яков", "Якун", "Ян", "Ярослав")

    For i = LBound(arMenNames) To UBound(arMenNames)
        If sName = arMenNames(i) Then
            IsMan = True
            Exit Function
        End If
    Next i
End Function

Private Function IsWoman(ByVal sName As String) As Boolean
    Dim arWomenNames As Variant
    Dim i As Long

    arWomenNames = Array("Августа", "Авдотья", "Агафья", "Агриппина", "Адиля", "Аида", "Аиша", "Айару", "Айгерим", "Айгуль", "Айнур", "Айнура", "Аксинья", "Акулина", "Алевтина", "Александра", "Александрина", "Алексина", "Алёна", "Алеся", "Алина", "Алиса", "Алла", "Алсу", "Алтынай", "Альбина", "Альфия", "Амина", "Амра", "Анастасия", "Ангелина", _
        "Анель", "Анжела", "Анжелика", "Анна", "Антонина", "Арина", "Армине", "Аружан", "Асель", "Асем", "Асмик", "Асоль", "Ася", "Аурика", "Ая", "Аяла", "Айя", "Белла", "Бэлла", "Бося", "Валентина", "Валерия", "Варвара", "Василиса", "Вера", "Вероника", "Виктория", "Виолетта", "Владилена", "Владислава", "Галина", "Глафира", "Гузель", "Гулнар", "Гульнара", _
        "Гульшат", "Гюзель", "Давлят", "Дана", "Дарья", "Дария", "Джамиля", "Диана", "Диляра", "Дина", "Динара", "Ева", "Евгения", "Евдокия", "Евпраксия", "Евфросиния", "Екатерина", "Елена", "Елизавета", "Жанат", "Жанар", "Жанара", "Жанна", "Жанетта", "Жулдыз", "Зауре", "Земфира", "Зимфира", "Зинаида", "Злата", _
        "Зоя", "Иванна", "Инга", "Инесса", "Инна", "Ираида", "Ирина", "Ирма", "Ия", "Капитолина", "Карина", "Каринэ", "Каролина", "Катерина", "Катрин", "Кира", "Клавдия", "Клара", "Кристина", "Ксения", "Лада", "Лариса", "Лейла", "Лейли", "Лейсан", "Лениза", "Леся", "Лиана", "Лига", "Лидия", _
        "Лилия", "Лия", "Лэйсэн", "Любовь", "Людмила", "Ляйсан", "Мадина", "Майя", "Маргарита", "Маржан", "Мариана", "Марианна", "Марина", "Мария", "Марфа", "Матрёна", "Матрена", "Мацак", "Милена", "Милана", "Мира", "Мирослава", "Муза", "Муит", "Надежда", "Назира", "Наида", "Наина", "Наринэ", "Наталья", "Наталия", "Нелли", "Нигина", "Николета", _
        "Нина", "Нинель", "Нонна", "Оксана", "Октябрина", "Олеся", "Ольга", "Пелагея", "Полина", "Прасковья", "Раиса", "Регина", "Ригина", "Римма", "Рита", "Роза", "Розалия", "Ромина", "Русина", "Руслана", "Руфина", "Сабина", "Салтанат", "Светлана", "Серафима", "Снежана", "София", "Софья", "Стелла", "Стефания", _
        "Таисия", "Тайя", "Тамара", "Татевик", "Татьяна", "Томирис", "Ульяна", "Фаина", "Феврония", "Фёкла", "Феодора", "Ханзада", "Целестина", "Шамиля", "Элеонора", "Элина", "Элла", "Эльвира", "Эльза", "Эмилия", "Эмма", "Эсфирь", "Юлия", "Яна", "Ярослава")
    
    For i = LBound(arWomenNames) To UBound(arWomenNames)
        If sName = arWomenNames(i) Then
            IsWoman = True
            Exit Function
        End If
    Next i
End Function

Public Function GetSex(ByVal cell As String) As Integer
    Dim arWords
    Dim iGender As Integer ' Исправлено String -> Integer для корректной работы
    Dim i As Integer

    'разбираем на слова
    arWords = Split(Application.WorksheetFunction.Trim(cell), " ")
    
    iGender = 0
    'если имя есть в справочниках - определяем пол сразу
    For i = LBound(arWords) To UBound(arWords)
        If IsMan(arWords(i)) Then iGender = -1
        If IsWoman(arWords(i)) Then iGender = 1
    Next i
        
    'если имени нет в справочниках - пытаемся определить по отчеству
    If iGender = 0 Then
        For i = LBound(arWords) To UBound(arWords)
            If Right(arWords(i), 3) = "вна" Or Right(arWords(i), 3) = "чна" Then iGender = 1
            If Right(arWords(i), 3) = "вич" Or Right(arWords(i), 3) = "ьич" Then iGender = -1
        Next i
    End If
    
    GetSex = iGender
End Function

Public Function fio(NameAsText As String, Optional NameCase As String = "И", Optional ShortForm As Boolean = False) As String
    'выстраивает ФИО в правильном порядке, склоняет по падежам и, при желании, выводит в сокращенной форме

    Dim iGender As Integer
    Dim sName As String, sName2 As String, sMidName As String, sMidName2 As String, sSurName As String, sSurName2 As String
    Dim arWords
    
    '----------------------- ОПРЕДЕЛЯЕМ ГДЕ ИМЯ, ГДЕ ФАМИЛИЯ, А ГДЕ ОТЧЕСТВО -----------------------------------------
    iGender = 0
    iGender = GetSex(NameAsText)        'определяем пол
    arWords = Split(Application.WorksheetFunction.Trim(NameAsText), " ")        'разбираем ФИО на слова
        
    'если в ячейке полное ФИО, т.е. есть и отчество
    If UBound(arWords) = 2 Then
        If iGender = -1 Then
            If Right(arWords(1), 3) = "вич" Or Right(arWords(1), 3) = "тич" Then
                sSurName = arWords(2)
                sName = arWords(0)
                sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вич" Or Right(arWords(2), 3) = "тич" Then
                sSurName = arWords(0)
                sName = arWords(1)
                sMidName = arWords(2)
            End If
        End If

        If iGender = 1 Then
            If Right(arWords(1), 3) = "вна" Or Right(arWords(1), 3) = "чна" Then
                sSurName = arWords(2)
                sName = arWords(0)
                sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вна" Or Right(arWords(2), 3) = "чна" Then
                sSurName = arWords(0)
                sName = arWords(1)
                sMidName = arWords(2)
            End If
        End If
    End If
        
    'если есть только фамилия и имя без отчества - ищем имена по справочникам
    If UBound(arWords) = 1 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then
            sName = arWords(0)
            sSurName = arWords(1)
        End If
        If IsMan(arWords(1)) Or IsWoman(arWords(1)) Then
            sName = arWords(1)
            sSurName = arWords(0)
        End If
    End If
    
    'если в ячейке только одно слово - пытаемся по полу определить - это имя или фамилия
    If UBound(arWords) = 0 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then
            'если пол определился - значит это имя
            sName = arWords(0)
        Else
            'если не определился - значит это фамилия
            sSurName = arWords(0)
            'пытаемся определить пол по окончанию фамилии, если возможно
            If sSurName Like "*ов" Or sSurName Like "*ев" Or sSurName Like "*ин" Or sSurName Like "*ий" Or sSurName Like "*ой" Then iGender = -1
            If sSurName Like "*ва" Or sSurName Like "*на" Or sSurName Like "*ая" Then iGender = -1
            'если пол так и не определился, то выходим
            If iGender = 0 Then
                fio = ""
                Exit Function
            End If
        End If
    End If

    
    
    '--------------------------- ИМЕНИТЕЛЬНЫЙ ПАДЕЖ (КТО) ---------------------------------------------------------
    sName2 = sName
    sSurName2 = sSurName
    sMidName2 = sMidName
        
    '--------------------------- ДАТЕЛЬНЫЙ ПАДЕЖ (КОМУ) ---------------------------------------------------------
    
    If UCase(NameCase) = "Д" Or UCase(NameCase) = "D" Then
        'формируем дательный падеж для имени
        If sName <> "" Then
            sName2 = sName
            If iGender = -1 Then
                If sName Like "*[ая]" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*[бвгджзклмнпрстфхцчшщ]" Then sName2 = sName & "у"
                If sName Like "*[йь]" Then sName2 = Left(sName, Len(sName) - 1) & "ю"
            End If
            If iGender = 1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*ия" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ея" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ья" Then sName2 = Left(sName, Len(sName) - 1) & "е"
                If sName Like "*ь" Then sName2 = Left(sName, Len(sName) - 1) & "и"
            End If
        End If
        
        'формируем дательный падеж для отчества
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "е"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "у"
        End If
        
        'формируем дательный падеж для фамилии
        If sSurName <> "" Then
            sSurName2 = sSurName
            If iGender = -1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "е"
                If sSurName Like "*й" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ому"
                If sSurName Like "*ай" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ю"
                If sSurName Like "*ь" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ю"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName & "у"
                If sSurName Like "*ых" Or sSurName Like "*их" Or sSurName Like "*иа" Or sSurName Like "*ия" Or sSurName Like "*уя" Or sSurName Like "*ая" Then sSurName2 = sSurName
                If sSurName Like "*ок" Or sSurName Like "*их" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ку"
            End If
            If iGender = 1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ой"
                If sSurName Like "*ая" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ой"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName
            End If
            
        End If
    End If

        
    '--------------------------- РОДИТЕЛЬНЫЙ ПАДЕЖ (КОГО) ---------------------------------------------------------
    
    If UCase(NameCase) = "Р" Or UCase(NameCase) = "R" Then
        'формируем родительный падеж для имени
        If sName <> "" Then
            sName2 = sName
            If iGender = -1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "ы"
                If sName Like "*[бвгджзклмнпрстфхцчшщ]" Then sName2 = sName & "а"
                If sName Like "*[йь]" Then sName2 = Left(sName, Len(sName) - 1) & "я"
            End If
            If iGender = 1 Then
                If sName Like "*а" Then sName2 = Left(sName, Len(sName) - 1) & "ы"
                If sName Like "*ия" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ея" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ья" Then sName2 = Left(sName, Len(sName) - 1) & "и"
                If sName Like "*ь" Then sName2 = Left(sName, Len(sName) - 1) & "и"
            End If
        End If
        
        'формируем родительный падеж для отчества
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "ы"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "а"
        End If
        
        'формируем родительный падеж для фамилии
        If sSurName <> "" Then
            sSurName2 = sSurName
            If iGender = -1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ы"
                If sSurName Like "*й" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ого"
                If sSurName Like "*ай" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "я"
                If sSurName Like "*ь" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "я"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName & "а"
                If sSurName Like "*ок" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ка"
                If sSurName Like "*ых" Or sSurName Like "*их" Or sSurName Like "*иа" Or sSurName Like "*ия" Or sSurName Like "*уя" Or sSurName Like "*ая" Then sSurName2 = sSurName
            End If
            If iGender = 1 Then
                If sSurName Like "*а" Then sSurName2 = Left(sSurName, Len(sSurName) - 1) & "ой"
                If sSurName Like "*ая" Then sSurName2 = Left(sSurName, Len(sSurName) - 2) & "ой"
                If sSurName Like "*[бвгджзклмнпрстфхцчшщ]" Then sSurName2 = sSurName
            End If
        End If
    End If
    
    '------------------- ВЫВОДИМ РЕЗУЛЬТАТЫ -------------------------------------------------------------------------
    If ShortForm Then
        fio = sSurName2 & " " & Left(sName2, 1) & "." & Left(sMidName2, 1) & "."
    Else
        fio = sSurName2 & " " & sName2 & " " & sMidName2
    End If
            
    If sMidName = "" Then fio = Left(fio, Len(fio) - 1)    'если нет отчества - убираем лишний последний пробел или точку
    fio = Trim(fio)
        
End Function

' ==========================================================
' РЕГУЛЯРНЫЕ ВЫРАЖЕНИЯ (RegExp Engine)
' ==========================================================

' @description Извлекает текст по шаблону регулярного выражения
' @param Text [String] Исходный текст
' @param Pattern [String] Шаблон (например, "\d+" для цифр)
' @param item [Integer] Номер совпадения (1 - первое, -1 - вернуть все через разделитель)
' @return [String] Найденное значение или пустая строка
Public Function RegExpExtract(ByVal Text As String, ByVal Pattern As String, Optional ByVal item As Integer = 1, Optional ByVal MatchCase As Boolean = False, Optional ByVal delim As String = ", ") As String
    Dim objRegExp As Object
    Dim objMatches As Object
    Dim strResult As String
    Dim i As Integer
    
    On Error Resume Next
    
    ' Используем Late Binding для совместимости (не требует Tools->References)
    Set objRegExp = CreateObject("VBScript.RegExp")
    
    If objRegExp Is Nothing Then
        RegExpExtract = ""
        Exit Function
    End If
    
    With objRegExp
        .Global = True
        .IgnoreCase = Not MatchCase
        .Pattern = Pattern
    End With
    
    If objRegExp.Test(Text) Then
        Set objMatches = objRegExp.Execute(Text)
        
        If item = -1 Then
            ' Режим "Все совпадения": собираем через разделитель
            For i = 0 To objMatches.count - 1
                If strResult = "" Then
                    strResult = objMatches.item(i).value
                Else
                    strResult = strResult & delim & objMatches.item(i).value
                End If
            Next i
            RegExpExtract = strResult
        ElseIf item > 0 Then
            ' Режим "Конкретное совпадение"
            If objMatches.count >= item Then
                RegExpExtract = objMatches.item(item - 1).value
            End If
        End If
    End If
    
    Set objRegExp = Nothing
End Function

' @description Проверяет, соответствует ли текст шаблону (например, формат даты)
Public Function RegExpMatch(ByVal Text As String, ByVal Pattern As String) As Boolean
    Dim objRegExp As Object
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = False
        .IgnoreCase = True
        .Pattern = Pattern
    End With
    RegExpMatch = objRegExp.Test(Text)
    Set objRegExp = Nothing
End Function
