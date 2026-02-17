Attribute VB_Name = "mdlHelper"
' ==============================================================================
' Module: mdlHelper
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Version: 1.5.9 (FIXED: ParseDateSafe Syntax Error & Restored Search Functions)
' Description: Universal utility functions, Smart Position Parser & FIO Engine.
' ==============================================================================

Option Explicit

Public colFIO_Global As Long
Public colLichniyNomer_Global As Long
Public colZvanie_Global As Long
Public colDolzhnost_Global As Long
Public colVoinskayaChast_Global As Long

' ==============================================================================
' 1. INITIALIZATION & COLUMN FINDING
' ==============================================================================

Public Sub InitStaffColumnIndexes()
    Dim wsStaff As Worksheet
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    If Not FindColumnNumbers(wsStaff, colLichniyNomer_Global, colZvanie_Global, colFIO_Global, colDolzhnost_Global, colVoinskayaChast_Global) Then
        MsgBox "Корректные индексы столбцов не удалось определить. Работа программы невозможна.", vbCritical
        End
    End If
End Sub

' ВОССТАНОВЛЕННАЯ ФУНКЦИЯ (была потеряна)
Public Sub EnsureStaffColumnsInitialized()
    If colLichniyNomer_Global = 0 Or colFIO_Global = 0 Then
        InitStaffColumnIndexes
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

    colLichniyNomer = 0: colZvanie = 0: colFIO = 0: colDolzhnost = 0: colVoinskayaChast = 0
    foundFIO = False: foundDolzhnost = False

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For i = 1 To lastCol
        headerText = LCase(Trim(ws.Cells(1, i).value))
        
        ' 1. Личный номер
        If InStr(headerText, "личный номер") > 0 Then colLichniyNomer = i
        
        ' 2. Воинское звание
        If InStr(headerText, "воинское звание") > 0 Then colZvanie = i
        
        ' 3. Часть
        If InStr(headerText, "часть") > 0 Or InStr(headerText, "раздел персонала") > 0 Then colVoinskayaChast = i
        
        ' 4. ФИО (Лицо)
        If headerText = "лицо" Then
            If IsTextFIOColumn(ws, i) Then colFIO = i: foundFIO = True
        End If
        
        ' 5. Должность (Ищем ДЛИННУЮ с цифрами приоритетно)
        If InStr(headerText, "штатная должность") > 0 Then
            If IsLongPositionColumn(ws, i) Then
                colDolzhnost = i: foundDolzhnost = True
            ElseIf colDolzhnost = 0 And IsTextColumn(ws, i) Then
                colDolzhnost = i: foundDolzhnost = True
            End If
        End If
    Next i

    If colLichniyNomer > 0 And colZvanie > 0 And colFIO > 0 And colDolzhnost > 0 And colVoinskayaChast > 0 Then
        FindColumnNumbers = True
    Else
        FindColumnNumbers = False
        MsgBox "Ошибка при определении столбцов. Проверьте заголовки.", vbCritical
    End If
End Function

Private Function IsTextFIOColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim lastRow As Long, i As Long, value As String
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    Dim checkLimit As Long: checkLimit = IIf(lastRow > 50, 50, lastRow)
    Dim textCount As Long
    For i = 2 To checkLimit
        value = Trim(ws.Cells(i, colNum).value)
        If value <> "" And ContainsLetters(value) And InStr(value, " ") > 0 And Not IsNumeric(value) Then
            textCount = textCount + 1
        End If
    Next i
    IsTextFIOColumn = (textCount > 0)
End Function

Private Function IsTextColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, value As String
    For i = 2 To 20
        value = Trim(ws.Cells(i, colNum).value)
        If value <> "" And ContainsLetters(value) Then IsTextColumn = True: Exit Function
    Next i
    IsTextColumn = False
End Function

Private Function IsLongPositionColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, numCount As Long, totalCount As Long, val As String
    Dim lastRow As Long: lastRow = IIf(ws.Cells(ws.Rows.count, colNum).End(xlUp).Row > 50, 50, ws.Cells(ws.Rows.count, colNum).End(xlUp).Row)
    For i = 2 To lastRow
        val = Trim(ws.Cells(i, colNum).value)
        If val <> "" Then
            totalCount = totalCount + 1
            If ContainsLetters(val) And ContainsNumbers(val) Then numCount = numCount + 1
        End If
    Next i
    If totalCount > 0 Then IsLongPositionColumn = (numCount / totalCount) > 0.3 Else IsLongPositionColumn = False
End Function

' ==============================================================================
' 2. GENERAL UTILITIES
' ==============================================================================

Public Function ContainsLetters(Text As String) As Boolean
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "[a-zA-Zа-яА-Я]"
    ContainsLetters = objRegExp.Test(Text)
End Function

Public Function ContainsNumbers(Text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        If char >= "0" And char <= "9" Then ContainsNumbers = True: Exit Function
    Next i
    ContainsNumbers = False
End Function

Public Function GetLastRow(ws As Worksheet, colIndex As Variant) As Long
    Dim colNum As Long
    If VarType(colIndex) = vbString Then colNum = ws.Range(CStr(colIndex) & "1").Column Else colNum = CLng(colIndex)
    GetLastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
End Function

Public Function ExtractVoinskayaChast(inputText As String) As String
    ExtractVoinskayaChast = RegExpExtract(inputText, "\d{4,5}")
    If ExtractVoinskayaChast = "" Then ExtractVoinskayaChast = inputText
End Function

Public Function GetExportCutoffDate() As Date
    GetExportCutoffDate = DateAdd("m", -1, DateAdd("yyyy", -3, Date))
End Function

' ==============================================================================
' 3. SETTINGS & CONFIGURATION
' ==============================================================================

Public Function GetSettingCutBattalion() As Boolean
    On Error Resume Next
    Dim ws As Worksheet, val As Variant
    Set ws = ThisWorkbook.Sheets("Настройки")
    If ws Is Nothing Then GetSettingCutBattalion = True: Exit Function
    val = ws.Range("B2").value
    If IsEmpty(val) Then GetSettingCutBattalion = True: Exit Function
    If UCase(CStr(val)) = "НЕТ" Or val = False Or val = 0 Then GetSettingCutBattalion = False Else GetSettingCutBattalion = True
End Function

Public Sub SetupSettingsSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Настройки")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "Настройки"
        ws.Cells(1, 1).value = "Параметр": ws.Cells(1, 2).value = "Значение"
        ws.Cells(2, 1).value = "Обрезать название батальона?": ws.Cells(2, 2).value = "ДА"
        ws.Columns("A:B").AutoFit
        MsgBox "Лист 'Настройки' создан.", vbInformation
    Else
        MsgBox "Лист 'Настройки' уже существует.", vbInformation
    End If
End Sub

' ==============================================================================
' 4. DATE PARSING & SEARCH (FIXED)
' ==============================================================================

Public Function ParseDateSafe(val As Variant) As Date
    On Error Resume Next
    ParseDateSafe = 0
    If IsEmpty(val) Or Trim(CStr(val)) = "" Then Exit Function
    
    Dim d As Date, sVal As String
    sVal = Trim(CStr(val))
    
    ' 1. Стандартная конвертация
    If IsDate(sVal) Then
        d = CDate(sVal)
        If Year(d) < 2000 And Year(d) > 1900 Then d = DateSerial(Year(d) + 100, Month(d), Day(d))
        If Year(d) > 2000 And Year(d) < 2100 Then ParseDateSafe = d: Exit Function
    End If
    
    ' 2. Ручной разбор (ИСПРАВЛЕНА СИНТАКСИЧЕСКАЯ ОШИБКА ЗДЕСЬ)
    Dim parts() As String
    If InStr(sVal, ".") > 0 Then
        parts = Split(sVal, ".")
    ElseIf InStr(sVal, "/") > 0 Then
        parts = Split(sVal, "/")
    End If
    
    If (Not Not parts) <> 0 Then
        If UBound(parts) = 2 Then
            Dim y As Integer: y = CInt(parts(2))
            If y < 100 Then y = 2000 + y
            If y > 2000 And y < 2100 Then ParseDateSafe = DateSerial(y, CInt(parts(1)), CInt(parts(0)))
        End If
    End If
End Function

Public Function GetStaffData(queryValue As String, Optional byLichniyNomer As Boolean = True) As Object
    Call EnsureStaffColumnsInitialized
    Dim ws As Worksheet, d As Object, r As Long, searchCol As Long
    Set ws = ThisWorkbook.Sheets("Штат")
    Set d = CreateObject("Scripting.Dictionary")
    
    If byLichniyNomer Then searchCol = colLichniyNomer_Global Else searchCol = colFIO_Global
    r = FindStaffRow(ws, queryValue, searchCol)
    
    If r > 0 Then
        d("Лицо") = ws.Cells(r, colFIO_Global).value
        d("Личный номер") = ws.Cells(r, colLichniyNomer_Global).value
        d("Воинское звание") = ws.Cells(r, colZvanie_Global).value
        d("Часть") = ws.Cells(r, colVoinskayaChast_Global).value
        d("Штатная должность") = ws.Cells(r, colDolzhnost_Global).value
    End If
    Set GetStaffData = d
End Function

Public Function FindEmployeeByAnyNumber(number As String) As Object
    Dim res As Object
    Set res = GetStaffData(number, True)
    
    ' Если не нашли по личному, пробуем по табельному
    If res.count = 0 Then
        Dim wsStaff As Worksheet
        Dim colTable As Long
        Set wsStaff = ThisWorkbook.Sheets("Штат")
        colTable = FindTableNumberColumn(wsStaff)
        
        If colTable > 0 Then
            Dim r As Long
            r = FindStaffRow(wsStaff, number, colTable)
            If r > 0 Then
                Set res = CreateObject("Scripting.Dictionary")
                res("Лицо") = wsStaff.Cells(r, colFIO_Global).value
                res("Личный номер") = wsStaff.Cells(r, colLichniyNomer_Global).value
                res("Воинское звание") = wsStaff.Cells(r, colZvanie_Global).value
                res("Часть") = wsStaff.Cells(r, colVoinskayaChast_Global).value
                res("Штатная должность") = wsStaff.Cells(r, colDolzhnost_Global).value
            End If
        End If
    End If
    Set FindEmployeeByAnyNumber = res
End Function

Public Function FindTableNumberColumn(ws As Worksheet) As Long
    Dim i As Long, val As Variant
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If LCase(Trim(ws.Cells(1, i).value)) = "лицо" Then
            ' Проверяем, числовые ли там данные
            val = ws.Cells(2, i).value
            If IsNumeric(val) And Not IsEmpty(val) Then
                FindTableNumberColumn = i
                Exit Function
            End If
        End If
    Next i
    FindTableNumberColumn = 0
End Function

Public Function GetStaffDataByTableNumber(tableNumber As String) As Object
    ' Алиас для совместимости
    Set GetStaffDataByTableNumber = FindEmployeeByAnyNumber(tableNumber)
End Function

Public Sub SaveWordDocumentSafe(wdDoc As Object, filePath As String)
    On Error Resume Next
    wdDoc.SaveAs2 filePath, 16 ' wdFormatXMLDocument
    If Err.number <> 0 Then wdDoc.SaveAs filePath
    On Error GoTo 0
End Sub

Public Function IsWordAvailable() As Boolean
    On Error Resume Next
    Dim app As Object: Set app = CreateObject("Word.Application")
    IsWordAvailable = Not app Is Nothing
    If Not app Is Nothing Then app.Quit
End Function

Public Sub CollectAllPersonPeriods(ws As Worksheet, rowNum As Long, periods As Collection)
    Dim lastCol As Long, j As Long, d1 As Date, d2 As Date
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    For j = 5 To lastCol Step 2
        d1 = ParseDateSafe(ws.Cells(rowNum, j).value)
        d2 = ParseDateSafe(ws.Cells(rowNum, j + 1).value)
        If d1 > 0 And d2 > 0 Then
            Dim p As Collection: Set p = New Collection
            p.Add d1: p.Add d2: p.Add (DateDiff("d", d1, d2) + 1)
            periods.Add p
        End If
    Next j
End Sub

Public Function FormatPeriodsForWord(periods As Collection, cutoff As Date, ByRef t As Long, ByRef r As Long) As String
    Dim p As Collection, s As String, i As Long
    t = 0: r = 0
    If periods.count = 0 Then Exit Function
    
    Dim sorted As Collection
    Set sorted = SortPeriodsByDateStart(periods)
    
    For i = 1 To sorted.count
        Set p = sorted(i)
        t = t + p(3)
        s = s & "- с " & Format(p(1), "dd.mm.yy") & " по " & Format(p(2), "dd.mm.yy") & " (" & p(3) & " сут.)"
        If p(2) < cutoff Then s = s & " (НЕ АКТУАЛЕН)"
        s = s & vbCrLf
    Next i
    r = (t \ 3) * 2
    FormatPeriodsForWord = s
End Function

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

' ==============================================================================
' 5. SMART POSITION PARSER (FIXED & RESTORED)
' ==============================================================================

Public Function SklonitDolzhnost(dolzhnost As String, VoinskayaChast As String) As String
    Dim clean As String, role As String, body As String, res As String
    clean = LCase(Trim(dolzhnost))
    
    ' 1. Режем хвост
    clean = CutUnitTail(clean)
    
    ' 2. Разделяем
    Call SplitRoleAndBody(clean, role, body)
    
    ' 3. Склоняем роль
    role = SklonitVoennayaDolzhnost(role)
    
    res = role
    If body <> "" Then res = res & " " & body
    SklonitDolzhnost = res & " войсковой части " & VoinskayaChast
End Function

Private Function CutUnitTail(Text As String) As String
    Dim patterns As Variant, i As Long, cutBattalion As Boolean
    cutBattalion = GetSettingCutBattalion()
    Dim t As String: t = Text
    
    ' 1. ОБЯЗАТЕЛЬНАЯ РЕЗКА: Юрлица (цифры + "отдельного", "гвардейского")
    ' ВАЖНО: Убрал просто "разведывательного" без цифр, чтобы не резало "командир разведывательного взвода"
    patterns = Array( _
        "(\d+\s+)?(отдельного|гвардейского|краснознаменного)\s+.*", _
        "\d+\s+(армии|дивизии|бригады|полка|батальона).*" _
    )
    For i = LBound(patterns) To UBound(patterns)
        t = RegExpReplace(t, patterns(i), "")
    Next i
    
    ' 2. Крупные соединения (без цифр) - всегда режем
    patterns = Array( _
        "(управления|штаба)\s+(полка|бригады|дивизии).*" _
    )
    For i = LBound(patterns) To UBound(patterns)
        t = RegExpReplace(t, patterns(i), "")
    Next i
    
    ' 3. Условная резка (Батальоны внутри полка)
    If cutBattalion Then
        patterns = Array("(\d+\s+)?(батальона|дивизиона|эскадрильи).*")
        For i = LBound(patterns) To UBound(patterns)
            t = RegExpReplace(t, patterns(i), "")
        Next i
    End If
    
    CutUnitTail = Trim(t)
End Function

Private Sub SplitRoleAndBody(Text As String, ByRef roleOut As String, ByRef bodyOut As String)
    Dim words() As String, i As Long, splitIdx As Long
    words = Split(Text, " ")
    splitIdx = UBound(words)
    
    For i = 0 To UBound(words)
        ' Если встретили слово-маркер подразделения ("взвода", "роты") - всё, что ДО него, это Роль.
        If IsUnitKeyword(words(i)) Then
            splitIdx = i - 1
            Exit For
        End If
    Next i
    
    Dim r As String, b As String
    For i = 0 To UBound(words)
        If i <= splitIdx Then r = r & words(i) & " " Else b = b & words(i) & " "
    Next i
    roleOut = Trim(r): bodyOut = Trim(b)
End Sub

Private Function IsUnitKeyword(word As String) As Boolean
    Select Case LCase(word)
        Case "взвода", "роты", "отделения", "расчета", "расчёта", "группы", "команды", "экипажа", "батареи", "службы", "пункта", "узла", "центра", "поста", "станции", "части", "секретной", "склада", "столовой", "гауптвахты"
            IsUnitKeyword = True
        Case Else
            IsUnitKeyword = False
    End Select
End Function

Public Function SklonitVoennayaDolzhnost(dolzhnost As String) As String
    Dim res As String: res = dolzhnost
    ' Префиксы
    If Left(res, 8) = "старший " Then res = "старшему " & Mid(res, 9)
    If Left(res, 8) = "младший " Then res = "младшему " & Mid(res, 9)
    If Left(res, 8) = "главный " Then res = "главному " & Mid(res, 9)
    If Left(res, 8) = "ведущий " Then res = "ведущему " & Mid(res, 9)
    ' Замены
    res = Replace(res, "командир", "командиру")
    res = Replace(res, "начальник", "начальнику")
    res = Replace(res, "заместитель", "заместителю")
    res = Replace(res, "помощник", "помощнику")
    res = Replace(res, "механик", "механику")
    res = Replace(res, "водитель", "водителю")
    res = Replace(res, "радиотелефонист", "радиотелефонисту")
    res = Replace(res, "разведчик", "разведчику")
    res = Replace(res, "оператор", "оператору")
    res = Replace(res, "наводчик", "наводчику")
    res = Replace(res, "инструктор", "инструктору")
    res = Replace(res, "техник", "технику")
    res = Replace(res, "электрик", "электрику")
    res = Replace(res, "пулеметчик", "пулеметчику")
    res = Replace(res, "гранатометчик", "гранатометчику")
    res = Replace(res, "стрелок", "стрелку")
    res = Replace(res, "сапер", "саперу")
    res = Replace(res, "снайпер", "снайперу")
    res = Replace(res, "курсант", "курсанту")
    res = Replace(res, "делопроизводитель", "делопроизводителю")
    res = Replace(res, "психолог", "психологу")
    res = Replace(res, "старшина", "старшине")
    res = Replace(res, "бухгалтер", "бухгалтеру")
    res = Replace(res, "врач", "врачу")
    res = Replace(res, "фельдшер", "фельдшеру")
    SklonitVoennayaDolzhnost = res
End Function

Public Function GetDolzhnostImenitelny(dolzhnost As String, VoinskayaChast As String) As String
    GetDolzhnostImenitelny = CutUnitTail(LCase(Trim(dolzhnost))) & " войсковой части " & VoinskayaChast
End Function

Public Function SklonitZvanie(zvanie As String) As String
    Dim z As String: z = LCase(Trim(zvanie))
    Select Case z
        Case "рядовой": SklonitZvanie = "Рядовому"
        Case "ефрейтор": SklonitZvanie = "Ефрейтору"
        Case "младший сержант": SklonitZvanie = "Младшему сержанту"
        Case "сержант": SklonitZvanie = "Сержанту"
        Case "старший сержант": SklonitZvanie = "Старшему сержанту"
        Case "старшина": SklonitZvanie = "Старшине"
        Case "прапорщик": SklonitZvanie = "Прапорщику"
        Case "старший прапорщик": SklonitZvanie = "Старшему прапорщику"
        Case "младший лейтенант": SklonitZvanie = "Младшему лейтенанту"
        Case "лейтенант": SklonitZvanie = "Лейтенанту"
        Case "старший лейтенант": SklonitZvanie = "Старшему лейтенанту"
        Case "капитан": SklonitZvanie = "Капитану"
        Case "майор": SklonitZvanie = "Майору"
        Case "подполковник": SklonitZvanie = "Подполковнику"
        Case "полковник": SklonitZvanie = "Полковнику"
        Case "генерал-майор": SklonitZvanie = "Генерал-майору"
        Case Else: SklonitZvanie = UCase(Left(z, 1)) & Mid(z, 2)
    End Select
End Function

Public Function GetZvanieImenitelny(zvanie As String) As String
    GetZvanieImenitelny = LCase(Trim(zvanie))
End Function

Public Function GetZvanieSkrasheno(zvanie As String) As String
    Dim z As String: z = LCase(Trim(zvanie))
    Select Case z
        Case "младший сержант": GetZvanieSkrasheno = "мл. сержанту"
        Case "старший сержант": GetZvanieSkrasheno = "ст. сержанту"
        Case "старший прапорщик": GetZvanieSkrasheno = "ст. прапорщику"
        Case "младший лейтенант": GetZvanieSkrasheno = "мл. лейтенанту"
        Case "старший лейтенант": GetZvanieSkrasheno = "ст. лейтенанту"
        Case Else: GetZvanieSkrasheno = SklonitZvanie(z)
    End Select
End Function

Public Function GetZvanieImenitelnyForSignature(zvanie As String) As String
    GetZvanieImenitelnyForSignature = UCase(Left(zvanie, 1)) & LCase(Mid(zvanie, 2))
End Function

Public Function GetFIOWithInitials(sName As String) As String
    Dim s As String: s = fio(sName, "Д", True)
    Dim p(): p = Split(s, " ")
    If UBound(p) = 1 Then GetFIOWithInitials = p(1) & " " & p(0) Else GetFIOWithInitials = s
End Function

Public Function GetFIOWithInitialsImenitelny(sName As String) As String
    Dim s As String: s = fio(sName, "И", True)
    Dim p(): p = Split(s, " ")
    If UBound(p) = 1 Then GetFIOWithInitialsImenitelny = p(1) & " " & p(0) Else GetFIOWithInitialsImenitelny = s
End Function

Public Function SklonitFIO(sName As String) As String
    SklonitFIO = fio(sName, "Д")
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

' ==========================================================
' FIO ENGINE (FULL)
' ==========================================================

Private Function IsMan(ByVal sName As String) As Boolean
    Dim arMenNames As Variant, i As Long
    arMenNames = Array("Абай", "Абрам", "Абраам", "Аваз", "Авазбек", "Авдей", "Адилет", "Адольф", "Азамат", "Акбар", "Аксентий", "Агафон", "Айбек", "Айрат", "Алдар", "Алишер", "Алан", "Александр", "Алексей", "Али", "Алмат", "Альберт", "Альвиан", "Альфред", "Анатолий", "Андрей", "Антон", "Антонин", "Аристарх", "Аркадий", "Армен", "Арнольд", "Арон", "Арсен", "Арсений", "Артем", "Артём", "Артемий", _
        "Артур", "Аскольд", "Афанасий", "Ашот", "Батыр", "Бауыржан", "Богдан", "Борис", "Вадим", "Валентин", "Валерий", "Валерьян", "Варлам", "Василий", "Вахтанг", "Венедикт", "Вениамин", "Виктор", "Виталий", "Влад", "Владилен", "Владимир", "Владислав", "Владлен", "Вольф", "Всеволод", "Вячеслав", "Гавриил", "Гаврил", "Гайдар", _
        "Геласий", "Геннадий", "Генрих", "Георгий", "Герасим", "Герман", "Глеб", "Гордей", "Григорий", "Гурген", "Давид", "Дамир", "Даниил", "Данил", "Данияр", "Дастан", "Демьян", "Денис", "Диас", "Динишбек", "Дмитрий", _
        "Дорофей", "Евгений", "Евграф", "Евдоким", "Евсей", "Егор", "Еремей", "Ернар", "Ермолай", "Ефим", "Жонибек", "Заур", "Зиновий", "Иакинф", "Иван", "Игнатий", "Игнат", "Игорь", "Иларион", "Илларион", "Ильдар", "Ильшат", "Илья", "Иннокентий", "Иосиф", "Ипполит", "Ирек", "Ириней", "Исидор", "Исаак", "Исхак", "Иулиан", "Казимир", "Кайрат", "Камиль", "Карл", "Касьян", "Керим", "Ким", "Кирилл", "Клавдий", "Кондрат", "Константин", _
        "Кристиан", "Кузьма", "Куприян", "Лаврентий", "Лев", "Ленар", "Леонард", "Леонид", "Леонтий", "Лука", "Лукий", "Лукьян", "Людвиг", "Магомед", "Магомет", "Майк", "Макар", "Максат", "Макс", "Максим", "Марат", "Марк", "Мартын", "Матвей", "Махач", "Махмуд", "Мелентий", "Мирлан", "Мирослав", _
        "Митрофан", "Михаил", "Модест", "Моисей", "Мстислав", "Мурад", "Мухамед", "Мухаммед", "Муса", "Мэлор", "Наум", "Никита", "Никифор", "Николай", "Нурбек", "Нуржан", "Нурлан", "Олег", "Онисим", "Осип", "Отар", "Павел", "Пантелеймон", "Парфений", "Пётр", "Петр", "Платон", "Порфирий", "Прокопий", "Протасий", "Прохор", "Радомир", "Разумник", "Рамазан", "Рамзан", "Рафаэль", _
        "Рафик", "Ринат", "Роман", "Роберт", "Ростислав", "Рубен", "Рудольф", "Руслан", "Рустам", "Рустем", "Сабир", "Савва", "Савелий", "Святослав", "Семён", "Семен", "Серафим", "Сергей", "Серик", "Созон", "Соломон", "Спиридон", "Станислав", "Степан", "Султан", "Тагир", "Тарас", "Темир", "Темирхан", "Тигран", "Тимофей", "Тимур", "Тихон", "Трифон", _
        "Трофим", "Фадей", "Фаддей", "Федор", "Фёдор", "Федосей", "Федот", "Феликс", "Филат", "Филипп", "Фома", "Фрол", "Харитон", "Хафиз", "Христофор", "Чеслав", "Шамиль", "Шамхал", "Эдуард", "Эльдар", "Эльман", "Эмиль", "Эммануил", "Эраст", "Юлиан", "Юлиус", "Юлий", "Юрий", "Юстин", "Яков", "Якун", "Ян", "Ярослав")
    For i = LBound(arMenNames) To UBound(arMenNames)
        If sName = arMenNames(i) Then IsMan = True: Exit Function
    Next i
End Function

Private Function IsWoman(ByVal sName As String) As Boolean
    Dim arWomenNames As Variant, i As Long
    arWomenNames = Array("Августа", "Авдотья", "Агафья", "Агриппина", "Адиля", "Аида", "Аиша", "Айару", "Айгерим", "Айгуль", "Айнур", "Айнура", "Аксинья", "Акулина", "Алевтина", "Александра", "Александрина", "Алексина", "Алёна", "Алеся", "Алина", "Алиса", "Алла", "Алсу", "Алтынай", "Альбина", "Альфия", "Амина", "Амра", "Анастасия", "Ангелина", _
        "Анель", "Анжела", "Анжелика", "Анна", "Антонина", "Арина", "Армине", "Аружан", "Асель", "Асем", "Асмик", "Асоль", "Ася", "Аурика", "Ая", "Аяла", "Айя", "Белла", "Бэлла", "Бося", "Валентина", "Валерия", "Варвара", "Василиса", "Вера", "Вероника", "Виктория", "Виолетта", "Владилена", "Владислава", "Галина", "Глафира", "Гузель", "Гулнар", "Гульнара", _
        "Гульшат", "Гюзель", "Давлят", "Дана", "Дарья", "Дария", "Джамиля", "Диана", "Диляра", "Дина", "Динара", "Ева", "Евгения", "Евдокия", "Евпраксия", "Евфросиния", "Екатерина", "Елена", "Елизавета", "Жанат", "Жанар", "Жанара", "Жанна", "Жанетта", "Жулдыз", "Зауре", "Земфира", "Зимфира", "Зинаида", "Злата", _
        "Зоя", "Иванна", "Инга", "Инесса", "Инна", "Ираида", "Ирина", "Ирма", "Ия", "Капитолина", "Карина", "Каринэ", "Каролина", "Катерина", "Катрин", "Кира", "Клавдия", "Клара", "Кристина", "Ксения", "Лада", "Лариса", "Лейла", "Лейли", "Лейсан", "Лениза", "Леся", "Лиана", "Лига", "Лидия", _
        "Лилия", "Лия", "Лэйсэн", "Любовь", "Людмила", "Ляйсан", "Мадина", "Майя", "Маргарита", "Маржан", "Мариана", "Марианна", "Марина", "Мария", "Марфа", "Матрёна", "Матрена", "Мацак", "Милена", "Милана", "Мира", "Мирослава", "Муза", "Муит", "Надежда", "Назира", "Наида", "Наина", "Наринэ", "Наталья", "Наталия", "Нелли", "Нигина", "Николета", _
        "Нина", "Нинель", "Нонна", "Оксана", "Октябрина", "Олеся", "Ольга", "Пелагея", "Полина", "Прасковья", "Раиса", "Регина", "Ригина", "Римма", "Рита", "Роза", "Розалия", "Ромина", "Русина", "Руслана", "Руфина", "Сабина", "Салтанат", "Светлана", "Серафима", "Снежана", "София", "Софья", "Стелла", "Стефания", _
        "Таисия", "Тайя", "Тамара", "Татевик", "Татьяна", "Томирис", "Ульяна", "Фаина", "Феврония", "Фёкла", "Феодора", "Ханзада", "Целестина", "Шамиля", "Элеонора", "Элина", "Элла", "Эльвира", "Эльза", "Эмилия", "Эмма", "Эсфирь", "Юлия", "Яна", "Ярослава")
    For i = LBound(arWomenNames) To UBound(arWomenNames)
        If sName = arWomenNames(i) Then IsWoman = True: Exit Function
    Next i
End Function

Public Function GetSex(ByVal cell As String) As Integer
    Dim arWords, iGender As Integer, i As Integer
    arWords = Split(Application.WorksheetFunction.Trim(cell), " ")
    iGender = 0
    For i = LBound(arWords) To UBound(arWords)
        If IsMan(arWords(i)) Then iGender = -1
        If IsWoman(arWords(i)) Then iGender = 1
    Next i
    If iGender = 0 Then
        For i = LBound(arWords) To UBound(arWords)
            If Right(arWords(i), 3) = "вна" Or Right(arWords(i), 3) = "чна" Then iGender = 1
            If Right(arWords(i), 3) = "вич" Or Right(arWords(i), 3) = "ьич" Then iGender = -1
        Next i
    End If
    GetSex = iGender
End Function

Public Function fio(NameAsText As String, Optional NameCase As String = "И", Optional ShortForm As Boolean = False) As String
    Dim iGender As Integer
    Dim sName As String, sName2 As String, sMidName As String, sMidName2 As String, sSurName As String, sSurName2 As String
    Dim arWords
    
    iGender = 0
    iGender = GetSex(NameAsText)
    arWords = Split(Application.WorksheetFunction.Trim(NameAsText), " ")
        
    If UBound(arWords) = 2 Then
        If iGender = -1 Then
            If Right(arWords(1), 3) = "вич" Or Right(arWords(1), 3) = "тич" Then
                sSurName = arWords(2): sName = arWords(0): sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вич" Or Right(arWords(2), 3) = "тич" Then
                sSurName = arWords(0): sName = arWords(1): sMidName = arWords(2)
            End If
        End If
        If iGender = 1 Then
            If Right(arWords(1), 3) = "вна" Or Right(arWords(1), 3) = "чна" Then
                sSurName = arWords(2): sName = arWords(0): sMidName = arWords(1)
            End If
            If Right(arWords(2), 3) = "вна" Or Right(arWords(2), 3) = "чна" Then
                sSurName = arWords(0): sName = arWords(1): sMidName = arWords(2)
            End If
        End If
    End If
        
    If UBound(arWords) = 1 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then sName = arWords(0): sSurName = arWords(1)
        If IsMan(arWords(1)) Or IsWoman(arWords(1)) Then sName = arWords(1): sSurName = arWords(0)
    End If
    
    If UBound(arWords) = 0 Then
        If IsMan(arWords(0)) Or IsWoman(arWords(0)) Then
            sName = arWords(0)
        Else
            sSurName = arWords(0)
            If sSurName Like "*ов" Or sSurName Like "*ев" Or sSurName Like "*ин" Or sSurName Like "*ий" Or sSurName Like "*ой" Then iGender = -1
            If sSurName Like "*ва" Or sSurName Like "*на" Or sSurName Like "*ая" Then iGender = -1
            If iGender = 0 Then fio = "": Exit Function
        End If
    End If

    sName2 = sName: sSurName2 = sSurName: sMidName2 = sMidName
        
    If UCase(NameCase) = "Д" Or UCase(NameCase) = "D" Then
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
        
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "е"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "у"
        End If
        
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

    If UCase(NameCase) = "Р" Or UCase(NameCase) = "R" Then
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
        
        If sMidName <> "" Then
            sMidName2 = sMidName
            If Right(sMidName, 1) = "а" Then sMidName2 = Left(sMidName, Len(sMidName) - 1) & "ы"
            If Right(sMidName, 1) = "ч" Then sMidName2 = sMidName & "а"
        End If
        
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
    
    If ShortForm Then
        fio = sSurName2 & " " & Left(sName2, 1) & "." & Left(sMidName2, 1) & "."
    Else
        fio = sSurName2 & " " & sName2 & " " & sMidName2
    End If
    If sMidName = "" Then fio = Left(fio, Len(fio) - 1)
    fio = Trim(fio)
End Function

' ==========================================================
' REGEXP ENGINE
' ==========================================================

Public Function RegExpExtract(ByVal Text As String, ByVal Pattern As String, Optional ByVal item As Integer = 1, Optional ByVal MatchCase As Boolean = False, Optional ByVal delim As String = ", ") As String
    Dim objRegExp As Object, objMatches As Object, strResult As String, i As Integer
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    If objRegExp Is Nothing Then RegExpExtract = "": Exit Function
    With objRegExp
        .Global = True
        .IgnoreCase = Not MatchCase
        .Pattern = Pattern
    End With
    If objRegExp.Test(Text) Then
        Set objMatches = objRegExp.Execute(Text)
        If item = -1 Then
            For i = 0 To objMatches.count - 1
                If strResult = "" Then strResult = objMatches.item(i).value Else strResult = strResult & delim & objMatches.item(i).value
            Next i
            RegExpExtract = strResult
        ElseIf item > 0 Then
            If objMatches.count >= item Then RegExpExtract = objMatches.item(item - 1).value
        End If
    End If
    Set objRegExp = Nothing
End Function

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

Public Function RegExpReplace(ByVal Text As String, ByVal Pattern As String, ByVal ReplaceWith As String) As String
    Dim objRegExp As Object
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    If objRegExp Is Nothing Then RegExpReplace = Text: Exit Function
    With objRegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = Pattern
    End With
    RegExpReplace = objRegExp.Replace(Text, ReplaceWith)
    Set objRegExp = Nothing
End Function

