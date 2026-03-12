Attribute VB_Name = "mdlZP12Validation"
Option Explicit

Private Const HISTORY_SHEET_NAME As String = "История_проверок_ZP12"
Private Const HISTORY_HEADER_ROW As Long = 1
Private Const TEMPLATE_HEADER_ROW As Long = 1
Private Const TEMPLATE_SUBHEADER_ROW As Long = 2
Private Const TEMPLATE_DATA_START_ROW As Long = 4

Private Const COL_TABLE_NUMBER As Long = 2
Private Const COL_PERSONAL_NUMBER As Long = 3
Private Const COL_RANK As Long = 4
Private Const COL_SURNAME As Long = 5
Private Const COL_NAME As Long = 6
Private Const COL_PATRONYMIC As Long = 7
Private Const COL_BIRTHDATE As Long = 8
Private Const COL_PERIOD_FROM As Long = 9
Private Const COL_PERIOD_TO As Long = 10
Private Const COL_FOUNDATION As Long = 11

Private Const STATUS_NEW As String = "NEW"
Private Const STATUS_OPEN As String = "OPEN"
Private Const STATUS_RESOLVED As String = "RESOLVED"
Private Const DUPLICATE_ROW_MARKER As Long = -1

Public Sub ValidateZP12Template(Optional ByVal selectedFilePath As String = "", Optional ByVal suppressSummary As Boolean = False)
    Dim selectedFile As Variant
    Dim templateWb As Workbook
    Dim templateWs As Worksheet
    Dim historyWs As Worksheet
    Dim staffContext As Object
    Dim previousErrors As Object
    Dim currentErrors As Collection
    Dim resolvedErrors As Collection
    Dim currentKeys As Object
    Dim counts As Object
    Dim checkedRows As Long
    Dim previousRunId As Long
    Dim currentRunId As Long
    Dim startedAt As Date

    On Error GoTo ErrorHandler

    If Len(Trim$(selectedFilePath)) > 0 Then
        selectedFile = selectedFilePath
    Else
        selectedFile = PickTemplateFile()
        If VarType(selectedFile) = vbBoolean Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Открытие шаблона ZP12..."

    Set templateWb = GetOrOpenWorkbook(CStr(selectedFile))
    Set templateWs = FindTemplateSheet(templateWb)
    If templateWs Is Nothing Then
        MsgBox "Не удалось найти лист шаблона ZP12 с ожидаемой структурой.", vbExclamation, "Проверка ZP12"
        GoTo SafeExit
    End If

    Set historyWs = GetOrCreateHistorySheet(templateWb)
    Set previousErrors = LoadPreviousRunErrors(historyWs, previousRunId)
    currentRunId = previousRunId + 1
    startedAt = Now

    Application.StatusBar = "Подготовка данных Штат..."
    Set staffContext = BuildStaffContext()

    Application.StatusBar = "Очистка предыдущей разметки..."
    ClearTemplateValidation templateWs

    Application.StatusBar = "Проверка строк шаблона..."
    Set currentErrors = ValidateTemplateRows(templateWs, staffContext, checkedRows)
    Set currentKeys = BuildCurrentKeyMap(currentErrors)
    ApplyStatuses currentErrors, previousErrors
    Set resolvedErrors = BuildResolvedErrors(previousErrors, currentKeys)

    Application.StatusBar = "Подсветка ошибок в шаблоне..."
    ApplyErrorsToSheet templateWs, currentErrors

    Application.StatusBar = "Запись истории проверки..."
    AppendHistoryRecords historyWs, currentRunId, startedAt, templateWb.Name, templateWs.Name, currentErrors, resolvedErrors

    Set counts = BuildStatusCounts(currentErrors, resolvedErrors)

    templateWb.Activate
    templateWs.Activate
    Application.StatusBar = False

    If Not suppressSummary Then
        MsgBox BuildSummaryMessage(checkedRows, counts, templateWb.Name, historyWs.Name), vbInformation, "Проверка ZP12"
    End If

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    MsgBox "Ошибка проверки ZP12: " & Err.description, vbCritical, "Проверка ZP12"
End Sub


Private Function PickTemplateFile() As Variant
    PickTemplateFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Выберите файл шаблона ZP12")
End Function

Private Function GetOrOpenWorkbook(ByVal filePath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook = Application.Workbooks.Open(fileName:=filePath, UpdateLinks:=0, ReadOnly:=False)
End Function

Private Function FindTemplateSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If IsTemplateSheet(ws) Then
            Set FindTemplateSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function IsTemplateSheet(ByVal ws As Worksheet) As Boolean
    IsTemplateSheet = _
        InStr(1, NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_TABLE_NUMBER)), "табельный номер в изделии", vbTextCompare) > 0 And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_PERSONAL_NUMBER)) = "личный номер" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_RANK)) = "воинское звание" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_SURNAME)) = "фамилия" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_NAME)) = "имя" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_PATRONYMIC)) = "отчество" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_BIRTHDATE)) = "дата рождения" And _
        NormalizeText(GetCellText(ws, TEMPLATE_HEADER_ROW, COL_PERIOD_FROM)) = "период" And _
        NormalizeText(GetCellText(ws, TEMPLATE_SUBHEADER_ROW, COL_PERIOD_FROM)) = "с" And _
        NormalizeText(GetCellText(ws, TEMPLATE_SUBHEADER_ROW, COL_PERIOD_TO)) = "по"
End Function

Private Function BuildStaffContext() As Object
    Dim wsStaff As Worksheet
    Dim colPersonal As Long
    Dim colRank As Long
    Dim colFIO As Long
    Dim colPosition As Long
    Dim colUnit As Long
    Dim colTable As Long
    Dim colBirth As Long
    Dim colContractKind As Long
    Dim colContractType As Long
    Dim colEventType As Long
    Dim colEffectiveStart As Long
    Dim context As Object

    Set wsStaff = ThisWorkbook.Worksheets("Штат")
    If Not mdlHelper.FindColumnNumbers(wsStaff, colPersonal, colRank, colFIO, colPosition, colUnit) Then
        Err.Raise vbObjectError + 701, "mdlZP12Validation", "Не удалось определить базовые столбцы листа Штат."
    End If

    colTable = mdlHelper.FindTableNumberColumn(wsStaff)
    If colTable = 0 Then
        Err.Raise vbObjectError + 703, "mdlZP12Validation", "На листе Штат не найден числовой столбец Лицо."
    End If

    colBirth = FindRequiredHeaderColumn(wsStaff, "Дата рождения")
    colContractKind = FindRequiredHeaderColumn(wsStaff, "Вид контракта")
    colContractType = FindRequiredHeaderColumn(wsStaff, "Тип контракта")
    colEventType = FindRequiredHeaderColumn(wsStaff, "Вид мероприятия")
    colEffectiveStart = FindRequiredHeaderColumn(wsStaff, "Начало срока действия")

    Set context = CreateObject("Scripting.Dictionary")
    context.Add "Sheet", wsStaff
    context.Add "ColTable", colTable
    context.Add "ColPersonal", colPersonal
    context.Add "ColRank", colRank
    context.Add "ColFIO", colFIO
    context.Add "ColBirth", colBirth
    context.Add "ColContractKind", colContractKind
    context.Add "ColContractType", colContractType
    context.Add "ColEventType", colEventType
    context.Add "ColEffectiveStart", colEffectiveStart
    context.Add "TableIndex", BuildStaffIndex(wsStaff, colTable, True)
    context.Add "PersonalIndex", BuildStaffIndex(wsStaff, colPersonal, False)

    Set BuildStaffContext = context
End Function

Private Function FindRequiredHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim i As Long

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For i = 1 To lastCol
        If NormalizeText(GetCellText(ws, 1, i)) = NormalizeText(headerName) Then
            FindRequiredHeaderColumn = i
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 702, "mdlZP12Validation", "На листе Штат не найден столбец '" & headerName & "'."
End Function

Private Function BuildStaffIndex(ByVal ws As Worksheet, ByVal colNum As Long, ByVal numericMode As Boolean) As Object
    Dim result As Object
    Dim lastRow As Long
    Dim i As Long
    Dim key As String

    Set result = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row

    For i = 2 To lastRow
        If numericMode Then
            key = NormalizeTableNumber(GetCellText(ws, i, colNum))
        Else
            key = NormalizePersonalNumber(GetCellText(ws, i, colNum))
        End If

        If key <> "" Then
            If Not result.exists(key) Then
                result.Add key, i
            Else
                result(key) = DUPLICATE_ROW_MARKER
            End If
        End If
    Next i

    Set BuildStaffIndex = result
End Function

Private Function ValidateTemplateRows(ByVal ws As Worksheet, ByVal staffContext As Object, ByRef checkedRows As Long) As Collection
    Dim errors As Collection
    Dim periodGroups As Object
    Dim lastRow As Long
    Dim rowNum As Long

    Set errors = New Collection
    Set periodGroups = CreateObject("Scripting.Dictionary")
    checkedRows = 0
    lastRow = GetTemplateLastRow(ws)

    For rowNum = TEMPLATE_DATA_START_ROW To lastRow
        If RowHasTemplateData(ws, rowNum) Then
            checkedRows = checkedRows + 1
            ValidateTemplateRow ws, rowNum, staffContext, errors, periodGroups
        End If
    Next rowNum

    ApplyPeriodOverlapChecks periodGroups, errors
    Set ValidateTemplateRows = errors
End Function

Private Sub ValidateTemplateRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal staffContext As Object, ByVal errors As Collection, ByVal periodGroups As Object)
    Dim tableNumber As String
    Dim personalNumber As String
    Dim rankName As String
    Dim surname As String
    Dim givenName As String
    Dim patronymic As String
    Dim templateFIO As String
    Dim birthRaw As String
    Dim birthDate As Date
    Dim periodFromRaw As String
    Dim periodToRaw As String
    Dim periodFromDate As Date
    Dim periodToDate As Date
    Dim rowIdentity As String
    Dim resolvedIdentity As Object
    Dim resolvedRow As Long
    Dim resolvedKey As String
    Dim staffWs As Worksheet
    Dim staffFIO As String
    Dim staffRank As String
    Dim staffBirthDate As Date
    Dim effectiveStartDate As Date
    Dim eventType As String

    tableNumber = GetCellText(ws, rowNum, COL_TABLE_NUMBER)
    personalNumber = GetCellText(ws, rowNum, COL_PERSONAL_NUMBER)
    rankName = GetCellText(ws, rowNum, COL_RANK)
    surname = GetCellText(ws, rowNum, COL_SURNAME)
    givenName = GetCellText(ws, rowNum, COL_NAME)
    patronymic = GetCellText(ws, rowNum, COL_PATRONYMIC)
    birthRaw = GetCellText(ws, rowNum, COL_BIRTHDATE)
    periodFromRaw = GetCellText(ws, rowNum, COL_PERIOD_FROM)
    periodToRaw = GetCellText(ws, rowNum, COL_PERIOD_TO)
    templateFIO = ComposeTemplateFIO(surname, givenName, patronymic)
    rowIdentity = BuildRawIdentityKey(tableNumber, personalNumber, templateFIO)

    ValidateRequiredField rowNum, "B", tableNumber, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "C", personalNumber, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "D", rankName, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "E", surname, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "F", givenName, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "G", patronymic, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "H", birthRaw, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "I", periodFromRaw, rowIdentity, errors, tableNumber, personalNumber, templateFIO
    ValidateRequiredField rowNum, "J", periodToRaw, rowIdentity, errors, tableNumber, personalNumber, templateFIO

    birthDate = mdlHelper.ParseDateSafe(birthRaw)
    periodFromDate = mdlHelper.ParseDateSafe(periodFromRaw)
    periodToDate = mdlHelper.ParseDateSafe(periodToRaw)

    If birthRaw <> "" And birthDate = 0 Then
        AddError errors, rowNum, "H", "BIRTHDATE_INVALID", "Некорректная дата рождения.", rowIdentity, tableNumber, personalNumber, templateFIO, 0, 0
    End If

    If periodFromRaw <> "" And periodFromDate = 0 Then
        AddError errors, rowNum, "I", "PERIOD_INVALID", "Некорректная дата начала периода.", rowIdentity, tableNumber, personalNumber, templateFIO, 0, 0
    End If

    If periodToRaw <> "" And periodToDate = 0 Then
        AddError errors, rowNum, "J", "PERIOD_INVALID", "Некорректная дата окончания периода.", rowIdentity, tableNumber, personalNumber, templateFIO, 0, 0
    End If

    If periodFromDate > 0 And periodToDate > 0 And periodToDate < periodFromDate Then
        AddError errors, rowNum, "I:J", "PERIOD_REVERSED", "Дата окончания периода раньше даты начала.", rowIdentity, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
    End If

    Set resolvedIdentity = ResolveTemplateEmployee(tableNumber, personalNumber, staffContext)
    If resolvedIdentity("HasError") Then
        AddIdentityErrors errors, rowNum, resolvedIdentity, rowIdentity, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
        Exit Sub
    End If

    resolvedRow = CLng(resolvedIdentity("ResolvedRow"))
    resolvedKey = CStr(resolvedIdentity("ResolvedKey"))
    Set staffWs = staffContext("Sheet")

    staffRank = GetCellText(staffWs, resolvedRow, CLng(staffContext("ColRank")))
    staffFIO = GetCellText(staffWs, resolvedRow, CLng(staffContext("ColFIO")))
    staffBirthDate = mdlHelper.ParseDateSafe(GetCellText(staffWs, resolvedRow, CLng(staffContext("ColBirth"))))

    If NormalizeText(rankName) <> "" And NormalizeText(rankName) <> NormalizeText(staffRank) Then
        AddError errors, rowNum, "D", "RANK_MISMATCH", "Воинское звание не совпадает со значением на листе Штат.", resolvedKey, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
    End If

    If NormalizeText(templateFIO) <> "" And NormalizeFullNameText(templateFIO) <> NormalizeFullNameText(staffFIO) Then
        AddError errors, rowNum, "E:G", "FIO_MISMATCH", "ФИО не совпадает со значением на листе Штат: " & staffFIO, resolvedKey, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
    End If

    If birthDate > 0 And staffBirthDate > 0 Then
        If NormalizeDateValue(birthDate) <> NormalizeDateValue(staffBirthDate) Then
            AddError errors, rowNum, "H", "BIRTHDATE_MISMATCH", "Дата рождения не совпадает со значением на листе Штат.", resolvedKey, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
        End If
    End If

    If ContainsMobilization( _
        GetCellText(staffWs, resolvedRow, CLng(staffContext("ColContractKind"))), _
        GetCellText(staffWs, resolvedRow, CLng(staffContext("ColContractType")))) Then
        AddError errors, rowNum, "B:C", "MOBILIZATION_CONTRACT", "Для сотрудника указан контракт по мобилизации.", resolvedKey, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
    End If

    eventType = NormalizeText(GetCellText(staffWs, resolvedRow, CLng(staffContext("ColEventType"))))
    effectiveStartDate = mdlHelper.ParseDateSafe(GetCellText(staffWs, resolvedRow, CLng(staffContext("ColEffectiveStart"))))
    If periodFromDate > 0 And eventType = NormalizeText("Зачисление в списки части") Then
        If effectiveStartDate > 0 And periodFromDate < effectiveStartDate Then
            AddError errors, rowNum, "I", "START_BEFORE_EFFECTIVE", "Период начинается раньше даты 'Начало срока действия' на листе Штат.", resolvedKey, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
        End If
    End If

    If periodFromDate > 0 And periodToDate > 0 Then
        RegisterPeriodCandidate periodGroups, resolvedKey, rowNum, tableNumber, personalNumber, templateFIO, periodFromDate, periodToDate
    End If
End Sub

Private Sub ValidateRequiredField(ByVal rowNum As Long, ByVal fieldCode As String, ByVal value As String, ByVal identityKey As String, ByVal errors As Collection, ByVal tableNumber As String, ByVal personalNumber As String, ByVal fio As String)
    If Trim$(value) = "" Then
        AddError errors, rowNum, fieldCode, "REQUIRED", "Не заполнено обязательное поле.", identityKey, tableNumber, personalNumber, fio, 0, 0
    End If
End Sub

Private Function ResolveTemplateEmployee(ByVal tableNumber As String, ByVal personalNumber As String, ByVal context As Object) As Object
    Dim result As Object
    Dim tableKey As String
    Dim personalKey As String
    Dim tableRow As Long
    Dim personalRow As Long

    Set result = CreateObject("Scripting.Dictionary")
    result("HasError") = False
    result("TableState") = "EMPTY"
    result("PersonalState") = "EMPTY"
    result("TableRow") = 0
    result("PersonalRow") = 0
    result("ResolvedRow") = 0
    result("ResolvedKey") = ""
    result("ErrorCode") = ""
    result("ErrorMessage") = ""

    tableKey = NormalizeTableNumber(tableNumber)
    personalKey = NormalizePersonalNumber(personalNumber)

    tableRow = ResolveIndexedRow(context("TableIndex"), tableKey)
    personalRow = ResolveIndexedRow(context("PersonalIndex"), personalKey)

    result("TableRow") = tableRow
    result("PersonalRow") = personalRow
    result("TableState") = ResolveStateLabel(tableKey, tableRow)
    result("PersonalState") = ResolveStateLabel(personalKey, personalRow)

    If tableRow > 0 And personalRow > 0 Then
        If tableRow <> personalRow Then
            result("HasError") = True
            result("ErrorCode") = "ID_CONFLICT"
            result("ErrorMessage") = "Табельный номер и личный номер указывают на разных сотрудников."
            Set ResolveTemplateEmployee = result
            Exit Function
        End If

        result("ResolvedRow") = tableRow
        result("ResolvedKey") = BuildResolvedIdentityKey(context("Sheet"), tableRow, CLng(context("ColPersonal")))
        Set ResolveTemplateEmployee = result
        Exit Function
    End If

    If result("TableState") = "DUPLICATE" Or result("PersonalState") = "DUPLICATE" Then
        result("HasError") = True
        result("ErrorCode") = "STAFF_DUPLICATE"
        result("ErrorMessage") = "В листе Штат обнаружены дублирующиеся идентификаторы сотрудника."
        Set ResolveTemplateEmployee = result
        Exit Function
    End If

    If tableKey <> "" And tableRow = 0 Then
        result("HasError") = True
        result("ErrorCode") = "TABLE_NOT_FOUND"
        result("ErrorMessage") = "Сотрудник не найден по табельному номеру."
    End If

    If personalKey <> "" And personalRow = 0 Then
        result("HasError") = True
        If result("ErrorCode") = "TABLE_NOT_FOUND" Then
            result("ErrorCode") = "IDENTITY_NOT_FOUND"
            result("ErrorMessage") = "Сотрудник не найден по табельному и/или личному номеру."
        Else
            result("ErrorCode") = "PERSONAL_NOT_FOUND"
            result("ErrorMessage") = "Сотрудник не найден по личному номеру."
        End If
    End If

    Set ResolveTemplateEmployee = result
End Function

Private Sub AddIdentityErrors(ByVal errors As Collection, ByVal rowNum As Long, ByVal resolvedIdentity As Object, ByVal rowIdentity As String, ByVal tableNumber As String, ByVal personalNumber As String, ByVal fio As String, ByVal periodFromDate As Date, ByVal periodToDate As Date)
    Select Case CStr(resolvedIdentity("ErrorCode"))
        Case "ID_CONFLICT"
            AddError errors, rowNum, "B:C", "ID_CONFLICT", CStr(resolvedIdentity("ErrorMessage")), rowIdentity, tableNumber, personalNumber, fio, periodFromDate, periodToDate
        Case "STAFF_DUPLICATE"
            AddError errors, rowNum, "B:C", "STAFF_DUPLICATE", CStr(resolvedIdentity("ErrorMessage")), rowIdentity, tableNumber, personalNumber, fio, periodFromDate, periodToDate
        Case "TABLE_NOT_FOUND"
            AddError errors, rowNum, "B", "TABLE_NOT_FOUND", CStr(resolvedIdentity("ErrorMessage")), rowIdentity, tableNumber, personalNumber, fio, periodFromDate, periodToDate
        Case "PERSONAL_NOT_FOUND"
            AddError errors, rowNum, "C", "PERSONAL_NOT_FOUND", CStr(resolvedIdentity("ErrorMessage")), rowIdentity, tableNumber, personalNumber, fio, periodFromDate, periodToDate
        Case "IDENTITY_NOT_FOUND"
            AddError errors, rowNum, "B:C", "IDENTITY_NOT_FOUND", CStr(resolvedIdentity("ErrorMessage")), rowIdentity, tableNumber, personalNumber, fio, periodFromDate, periodToDate
    End Select
End Sub

Private Function ResolveIndexedRow(ByVal indexDict As Object, ByVal lookupKey As String) As Long
    If lookupKey = "" Then Exit Function
    If Not indexDict.exists(lookupKey) Then Exit Function
    ResolveIndexedRow = CLng(indexDict(lookupKey))
End Function

Private Function ResolveStateLabel(ByVal lookupKey As String, ByVal rowValue As Long) As String
    If lookupKey = "" Then
        ResolveStateLabel = "EMPTY"
    ElseIf rowValue = DUPLICATE_ROW_MARKER Then
        ResolveStateLabel = "DUPLICATE"
    ElseIf rowValue = 0 Then
        ResolveStateLabel = "NOT_FOUND"
    Else
        ResolveStateLabel = "FOUND"
    End If
End Function

Private Sub RegisterPeriodCandidate(ByVal periodGroups As Object, ByVal resolvedKey As String, ByVal rowNum As Long, ByVal tableNumber As String, ByVal personalNumber As String, ByVal fio As String, ByVal periodFromDate As Date, ByVal periodToDate As Date)
    Dim entry As Object
    Dim bucket As Collection

    Set entry = CreateObject("Scripting.Dictionary")
    entry("Row") = rowNum
    entry("TableNumber") = tableNumber
    entry("PersonalNumber") = personalNumber
    entry("FIO") = fio
    entry("PeriodFrom") = periodFromDate
    entry("PeriodTo") = periodToDate

    If Not periodGroups.exists(resolvedKey) Then
        Set bucket = New Collection
        periodGroups.Add resolvedKey, bucket
    End If

    Set bucket = periodGroups(resolvedKey)
    bucket.Add entry
End Sub

Private Sub ApplyPeriodOverlapChecks(ByVal periodGroups As Object, ByVal errors As Collection)
    Dim groupKey As Variant
    Dim sortedPeriods As Collection
    Dim i As Long
    Dim j As Long
    Dim entryA As Object
    Dim entryB As Object
    Dim ownKey As String
    Dim otherKey As String

    For Each groupKey In periodGroups.keys
        Set sortedPeriods = SortPeriodEntries(periodGroups(groupKey))
        For i = 1 To sortedPeriods.count - 1
            Set entryA = sortedPeriods(i)
            For j = i + 1 To sortedPeriods.count
                Set entryB = sortedPeriods(j)
                If CLng(NormalizeDateValue(entryB("PeriodFrom"))) > CLng(NormalizeDateValue(entryA("PeriodTo"))) Then Exit For
                If entryA("PeriodFrom") <= entryB("PeriodTo") And entryA("PeriodTo") >= entryB("PeriodFrom") Then
                    ownKey = BuildPeriodOverlapKey(CStr(groupKey), entryA("PeriodFrom"), entryA("PeriodTo"), entryB("PeriodFrom"), entryB("PeriodTo"))
                    otherKey = BuildPeriodOverlapKey(CStr(groupKey), entryB("PeriodFrom"), entryB("PeriodTo"), entryA("PeriodFrom"), entryA("PeriodTo"))

                    AddError errors, CLng(entryA("Row")), "I:J", "PERIOD_OVERLAP", _
                        "Период пересекается со строкой " & entryB("Row") & ": " & Format(entryB("PeriodFrom"), "dd.mm.yyyy") & " - " & Format(entryB("PeriodTo"), "dd.mm.yyyy"), _
                        ownKey, CStr(entryA("TableNumber")), CStr(entryA("PersonalNumber")), CStr(entryA("FIO")), entryA("PeriodFrom"), entryA("PeriodTo"), True

                    AddError errors, CLng(entryB("Row")), "I:J", "PERIOD_OVERLAP", _
                        "Период пересекается со строкой " & entryA("Row") & ": " & Format(entryA("PeriodFrom"), "dd.mm.yyyy") & " - " & Format(entryA("PeriodTo"), "dd.mm.yyyy"), _
                        otherKey, CStr(entryB("TableNumber")), CStr(entryB("PersonalNumber")), CStr(entryB("FIO")), entryB("PeriodFrom"), entryB("PeriodTo"), True
                End If
            Next j
        Next i
    Next groupKey
End Sub

Private Function SortPeriodEntries(ByVal entries As Collection) As Collection
    Dim arr() As Object
    Dim i As Long
    Dim j As Long
    Dim tmp As Object
    Dim result As Collection

    ReDim arr(1 To entries.count)
    For i = 1 To entries.count
        Set arr(i) = entries(i)
    Next i

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)("PeriodFrom") > arr(j)("PeriodFrom") Then
                Set tmp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = tmp
            End If
        Next j
    Next i

    Set result = New Collection
    For i = 1 To UBound(arr)
        result.Add arr(i)
    Next i

    Set SortPeriodEntries = result
End Function

Private Sub AddError(ByVal errors As Collection, ByVal rowNum As Long, ByVal fieldCode As String, ByVal errorCode As String, ByVal description As String, ByVal identityKey As String, ByVal tableNumber As String, ByVal personalNumber As String, ByVal fio As String, ByVal periodFromDate As Date, ByVal periodToDate As Date, Optional ByVal identityKeyIsFinal As Boolean = False)
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item("Row") = rowNum
    item("Field") = fieldCode
    item("ErrorCode") = errorCode
    item("Description") = description
    item("TableNumber") = tableNumber
    item("PersonalNumber") = personalNumber
    item("FIO") = fio
    item("PeriodFrom") = FormatDateForHistory(periodFromDate)
    item("PeriodTo") = FormatDateForHistory(periodToDate)
    item("Status") = ""
    item("PreviousRunID") = ""
    If identityKeyIsFinal Then
        item("ErrorKey") = identityKey
    Else
        item("ErrorKey") = BuildErrorKey(identityKey, errorCode, fieldCode, periodFromDate, periodToDate)
    End If
    errors.Add item
End Sub

Private Function BuildErrorKey(ByVal identityKey As String, ByVal errorCode As String, ByVal fieldCode As String, ByVal periodFromDate As Date, ByVal periodToDate As Date) As String
    BuildErrorKey = identityKey & "|" & errorCode & "|" & fieldCode & "|" & FormatDateForKey(periodFromDate) & "|" & FormatDateForKey(periodToDate)
End Function

Private Function BuildPeriodOverlapKey(ByVal resolvedKey As String, ByVal ownFrom As Date, ByVal ownTo As Date, ByVal otherFrom As Date, ByVal otherTo As Date) As String
    BuildPeriodOverlapKey = resolvedKey & "|PERIOD_OVERLAP|I:J|" & FormatDateForKey(ownFrom) & "|" & FormatDateForKey(ownTo) & "|" & FormatDateForKey(otherFrom) & "|" & FormatDateForKey(otherTo)
End Function

Private Function BuildRawIdentityKey(ByVal tableNumber As String, ByVal personalNumber As String, ByVal fio As String) As String
    BuildRawIdentityKey = "RAW|" & NormalizeTableNumber(tableNumber) & "|" & NormalizePersonalNumber(personalNumber) & "|" & NormalizeFullNameText(fio)
End Function

Private Function BuildResolvedIdentityKey(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal personalCol As Long) As String
    BuildResolvedIdentityKey = "STAFF|" & NormalizePersonalNumber(GetCellText(ws, rowNum, personalCol)) & "|" & rowNum
End Function

Private Function BuildCurrentKeyMap(ByVal currentErrors As Collection) As Object
    Dim result As Object
    Dim i As Long
    Dim item As Object

    Set result = CreateObject("Scripting.Dictionary")
    For i = 1 To currentErrors.count
        Set item = currentErrors(i)
        result(item("ErrorKey")) = True
    Next i

    Set BuildCurrentKeyMap = result
End Function

Private Sub ApplyStatuses(ByVal currentErrors As Collection, ByVal previousErrors As Object)
    Dim i As Long
    Dim item As Object
    Dim previousItem As Object

    For i = 1 To currentErrors.count
        Set item = currentErrors(i)
        If previousErrors.exists(item("ErrorKey")) Then
            Set previousItem = previousErrors(item("ErrorKey"))
            item("Status") = STATUS_OPEN
            item("PreviousRunID") = previousItem("RunID")
        Else
            item("Status") = STATUS_NEW
            item("PreviousRunID") = ""
        End If
    Next i
End Sub

Private Function BuildResolvedErrors(ByVal previousErrors As Object, ByVal currentKeys As Object) As Collection
    Dim result As Collection
    Dim key As Variant
    Dim item As Object

    Set result = New Collection
    For Each key In previousErrors.keys
        If Not currentKeys.exists(CStr(key)) Then
            Set item = CloneHistoryItem(previousErrors(key))
            item("Status") = STATUS_RESOLVED
            item("PreviousRunID") = previousErrors(key)("RunID")
            result.Add item
        End If
    Next key

    Set BuildResolvedErrors = result
End Function

Private Function CloneHistoryItem(ByVal sourceItem As Object) As Object
    Dim result As Object
    Dim key As Variant

    Set result = CreateObject("Scripting.Dictionary")
    For Each key In sourceItem.keys
        result(key) = sourceItem(key)
    Next key

    Set CloneHistoryItem = result
End Function

Private Function BuildStatusCounts(ByVal currentErrors As Collection, ByVal resolvedErrors As Collection) As Object
    Dim result As Object
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result("NEW") = 0
    result("OPEN") = 0
    result("RESOLVED") = resolvedErrors.count
    result("CURRENT") = currentErrors.count

    For i = 1 To currentErrors.count
        If currentErrors(i)("Status") = STATUS_NEW Then
            result("NEW") = CLng(result("NEW")) + 1
        ElseIf currentErrors(i)("Status") = STATUS_OPEN Then
            result("OPEN") = CLng(result("OPEN")) + 1
        End If
    Next i

    Set BuildStatusCounts = result
End Function

Private Function BuildSummaryMessage(ByVal checkedRows As Long, ByVal counts As Object, ByVal workbookName As String, ByVal historySheetName As String) As String
    BuildSummaryMessage = _
        "Проверка файла завершена." & vbCrLf & vbCrLf & _
        "Файл: " & workbookName & vbCrLf & _
        "Проверено строк: " & checkedRows & vbCrLf & _
        "Текущих ошибок: " & counts("CURRENT") & vbCrLf & _
        "Новых ошибок: " & counts("NEW") & vbCrLf & _
        "Ошибок без изменений: " & counts("OPEN") & vbCrLf & _
        "Исправлено ошибок: " & counts("RESOLVED") & vbCrLf & vbCrLf & _
        "История записана на лист '" & historySheetName & "'."
End Function

Private Sub ApplyErrorsToSheet(ByVal ws As Worksheet, ByVal errors As Collection)
    Dim i As Long
    Dim item As Object
    Dim fieldColumns As Collection
    Dim colIndex As Variant
    Dim commentMap As Object
    Dim addressKey As Variant

    Set commentMap = CreateObject("Scripting.Dictionary")

    For i = 1 To errors.count
        Set item = errors(i)
        Set fieldColumns = GetFieldColumns(CStr(item("Field")))
        For Each colIndex In fieldColumns
            ws.Cells(CLng(item("Row")), CLng(colIndex)).Interior.Color = RGB(255, 100, 100)
            AppendCommentText commentMap, ws.Cells(CLng(item("Row")), CLng(colIndex)).Address(False, False), CStr(item("Description"))
        Next colIndex
    Next i

    For Each addressKey In commentMap.keys
        SetCellComment ws.Range(CStr(addressKey)), CStr(commentMap(addressKey))
    Next addressKey
End Sub

Private Function GetFieldColumns(ByVal fieldCode As String) As Collection
    Dim result As Collection
    Dim parts() As String
    Dim i As Long

    Set result = New Collection
    parts = Split(fieldCode, ":")

    If UBound(parts) = 0 Then
        result.Add ColumnLetterToNumber(parts(0))
    Else
        For i = ColumnLetterToNumber(parts(0)) To ColumnLetterToNumber(parts(1))
            result.Add i
        Next i
    End If

    Set GetFieldColumns = result
End Function

Private Function ColumnLetterToNumber(ByVal columnCode As String) As Long
    ColumnLetterToNumber = Range(columnCode & "1").Column
End Function

Private Sub AppendCommentText(ByVal commentMap As Object, ByVal addressKey As String, ByVal lineText As String)
    If commentMap.exists(addressKey) Then
        If InStr(1, commentMap(addressKey), lineText, vbTextCompare) = 0 Then
            commentMap(addressKey) = commentMap(addressKey) & vbCrLf & "- " & lineText
        End If
    Else
        commentMap.Add addressKey, "- " & lineText
    End If
End Sub

Private Sub SetCellComment(ByVal targetCell As Range, ByVal commentText As String)
    RemoveCellComment targetCell
    targetCell.AddComment commentText
End Sub

Private Sub RemoveCellComment(ByVal targetCell As Range)
    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
    targetCell.ClearComments
    On Error GoTo 0
End Sub

Private Sub ClearTemplateValidation(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range

    lastRow = GetTemplateLastRow(ws)
    If lastRow < TEMPLATE_DATA_START_ROW Then Exit Sub

    Set rng = ws.Range(ws.Cells(TEMPLATE_DATA_START_ROW, COL_TABLE_NUMBER), ws.Cells(lastRow, COL_FOUNDATION))

    For Each cell In rng.Cells
        If cell.Interior.Color = RGB(255, 100, 100) Then
            cell.Interior.ColorIndex = xlNone
        End If
        RemoveCellComment cell
    Next cell
End Sub

Private Function GetTemplateLastRow(ByVal ws As Worksheet) As Long
    Dim colNum As Long
    Dim maxRow As Long
    Dim currentRow As Long

    maxRow = TEMPLATE_DATA_START_ROW - 1
    For colNum = COL_TABLE_NUMBER To COL_FOUNDATION
        currentRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
        If currentRow > maxRow Then maxRow = currentRow
    Next colNum

    GetTemplateLastRow = maxRow
End Function

Private Function RowHasTemplateData(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    RowHasTemplateData = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(rowNum, COL_TABLE_NUMBER), ws.Cells(rowNum, COL_FOUNDATION))) > 0
End Function

Private Function GetCellText(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As String
    If isError(ws.Cells(rowNum, colNum).value) Then Exit Function
    GetCellText = Trim$(CStr(ws.Cells(rowNum, colNum).value))
End Function

Private Function NormalizeText(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Trim$(Replace(Replace(CStr(value), vbCr, " "), vbLf, " "))
    cleaned = Replace(cleaned, Chr$(160), " ")
    cleaned = Replace(cleaned, "ё", "е")
    cleaned = Replace(cleaned, "Ё", "Е")

    Do While InStr(cleaned, "  ") > 0
        cleaned = Replace(cleaned, "  ", " ")
    Loop

    NormalizeText = LCase$(cleaned)
End Function

Private Function NormalizePersonalNumber(ByVal value As String) As String
    NormalizePersonalNumber = NormalizeText(Replace(CStr(value), " ", ""))
End Function

Private Function NormalizeTableNumber(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Replace(NormalizeText(value), " ", "")
    If cleaned = "" Then Exit Function

    If IsNumeric(cleaned) Then
        NormalizeTableNumber = CStr(CLng(cleaned))
    Else
        NormalizeTableNumber = cleaned
    End If
End Function

Private Function ComposeTemplateFIO(ByVal surname As String, ByVal givenName As String, ByVal patronymic As String) As String
    ComposeTemplateFIO = Trim$(surname & " " & givenName & " " & patronymic)
End Function

Private Function NormalizeFullNameText(ByVal value As String) As String
    NormalizeFullNameText = NormalizeText(value)
End Function

Private Function NormalizeDateValue(ByVal value As Date) As Long
    If value <= 0 Then Exit Function
    NormalizeDateValue = CLng(DateSerial(Year(value), Month(value), Day(value)))
End Function

Private Function FormatDateForKey(ByVal value As Date) As String
    If value <= 0 Then
        FormatDateForKey = "-"
    Else
        FormatDateForKey = Format$(value, "yyyymmdd")
    End If
End Function

Private Function FormatDateForHistory(ByVal value As Date) As String
    If value <= 0 Then
        FormatDateForHistory = ""
    Else
        FormatDateForHistory = Format$(value, "dd.mm.yyyy")
    End If
End Function

Private Function ContainsMobilization(ByVal contractKind As String, ByVal contractType As String) As Boolean
    ContainsMobilization = _
        InStr(1, NormalizeText(contractKind), "мобилиз", vbTextCompare) > 0 Or _
        InStr(1, NormalizeText(contractType), "мобилиз", vbTextCompare) > 0
End Function

Private Function GetOrCreateHistorySheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, HISTORY_SHEET_NAME, vbTextCompare) = 0 Then
            Set GetOrCreateHistorySheet = ws
            Exit Function
        End If
    Next ws

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = HISTORY_SHEET_NAME
    InitializeHistorySheet ws
    Set GetOrCreateHistorySheet = ws
End Function

Private Sub InitializeHistorySheet(ByVal ws As Worksheet)
    Dim headers As Variant

    headers = Array("RunID", "Проверено", "Имя файла", "Лист", "Строка шаблона", "Табельный", "Личный номер", "ФИО", "Период с", "Период по", "Код ошибки", "Поле", "Описание", "Статус", "Ключ ошибки", "Предыдущий RunID")
    ws.Range("A1:P1").value = headers

    On Error Resume Next
    ws.Rows(HISTORY_HEADER_ROW).Font.Bold = True
    ws.Rows(HISTORY_HEADER_ROW).Interior.Color = RGB(220, 230, 241)
    ws.Columns("A:P").HorizontalAlignment = xlLeft
    ws.Columns("A:P").EntireColumn.AutoFit
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:P1").AutoFilter
    EnsureHistorySheetLayout ws
    On Error GoTo 0
End Sub

Private Function LoadPreviousRunErrors(ByVal ws As Worksheet, ByRef lastRunId As Long) As Object
    Dim result As Object
    Dim lastRow As Long
    Dim rowNum As Long
    Dim runId As Long
    Dim item As Object
    Dim statusText As String

    Set result = CreateObject("Scripting.Dictionary")
    lastRunId = GetLastRunId(ws)
    If lastRunId = 0 Then
        Set LoadPreviousRunErrors = result
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For rowNum = 2 To lastRow
        runId = 0
        If IsNumeric(ws.Cells(rowNum, 1).value) Then runId = CLng(ws.Cells(rowNum, 1).value)
        If runId = lastRunId Then
            statusText = UCase$(Trim$(CStr(ws.Cells(rowNum, 14).value)))
            If statusText <> STATUS_RESOLVED Then
                Set item = CreateObject("Scripting.Dictionary")
                item("RunID") = runId
                item("CheckedAt") = CStr(ws.Cells(rowNum, 2).value)
                item("WorkbookName") = CStr(ws.Cells(rowNum, 3).value)
                item("SheetName") = CStr(ws.Cells(rowNum, 4).value)
                item("Row") = CLng(val(ws.Cells(rowNum, 5).value))
                item("TableNumber") = CStr(ws.Cells(rowNum, 6).value)
                item("PersonalNumber") = CStr(ws.Cells(rowNum, 7).value)
                item("FIO") = CStr(ws.Cells(rowNum, 8).value)
                item("PeriodFrom") = CStr(ws.Cells(rowNum, 9).value)
                item("PeriodTo") = CStr(ws.Cells(rowNum, 10).value)
                item("ErrorCode") = CStr(ws.Cells(rowNum, 11).value)
                item("Field") = CStr(ws.Cells(rowNum, 12).value)
                item("Description") = CStr(ws.Cells(rowNum, 13).value)
                item("Status") = statusText
                item("ErrorKey") = CStr(ws.Cells(rowNum, 15).value)
                item("PreviousRunID") = CStr(ws.Cells(rowNum, 16).value)
                result.Add CStr(item("ErrorKey")), item
            End If
        End If
    Next rowNum

    Set LoadPreviousRunErrors = result
End Function

Private Function GetLastRunId(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    Dim rowNum As Long

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For rowNum = lastRow To 2 Step -1
        If IsNumeric(ws.Cells(rowNum, 1).value) Then
            GetLastRunId = CLng(ws.Cells(rowNum, 1).value)
            Exit Function
        End If
    Next rowNum
End Function

Private Sub AppendHistoryRecords(ByVal ws As Worksheet, ByVal runId As Long, ByVal checkedAt As Date, ByVal workbookName As String, ByVal sheetName As String, ByVal currentErrors As Collection, ByVal resolvedErrors As Collection)
    Dim nextRow As Long
    Dim i As Long

    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    For i = 1 To currentErrors.count
        WriteHistoryRow ws, nextRow, runId, checkedAt, workbookName, sheetName, currentErrors(i)
        nextRow = nextRow + 1
    Next i

    For i = 1 To resolvedErrors.count
        WriteHistoryRow ws, nextRow, runId, checkedAt, workbookName, sheetName, resolvedErrors(i)
        nextRow = nextRow + 1
    Next i

    On Error Resume Next
    ws.Columns("A:P").AutoFit
    If Not ws.AutoFilterMode Then ws.Range("A1:P1").AutoFilter
    EnsureHistorySheetLayout ws
    On Error GoTo 0
End Sub

Private Sub WriteHistoryRow(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal runId As Long, ByVal checkedAt As Date, ByVal workbookName As String, ByVal sheetName As String, ByVal item As Object)
    ws.Cells(targetRow, 1).value = runId
    ws.Cells(targetRow, 2).value = Format$(checkedAt, "dd.mm.yyyy hh:mm:ss")
    ws.Cells(targetRow, 3).value = workbookName
    ws.Cells(targetRow, 4).value = sheetName
    ws.Cells(targetRow, 5).value = item("Row")
    ws.Cells(targetRow, 6).value = item("TableNumber")
    ws.Cells(targetRow, 7).value = item("PersonalNumber")
    ws.Cells(targetRow, 8).value = item("FIO")
    ws.Cells(targetRow, 9).value = item("PeriodFrom")
    ws.Cells(targetRow, 10).value = item("PeriodTo")
    ws.Cells(targetRow, 11).value = item("ErrorCode")
    ws.Cells(targetRow, 12).value = item("Field")
    ws.Cells(targetRow, 13).value = item("Description")
    ws.Cells(targetRow, 14).value = item("Status")
    ws.Cells(targetRow, 15).value = item("ErrorKey")
    ws.Cells(targetRow, 16).value = item("PreviousRunID")
End Sub

Private Sub EnsureHistorySheetLayout(ByVal ws As Worksheet)
    Dim previousSheet As Worksheet

    On Error Resume Next
    Set previousSheet = ws.Parent.ActiveSheet
    ws.Activate
    ws.Range("A2").Select
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.SplitColumn = 0
        ActiveWindow.SplitRow = 1
        ActiveWindow.FreezePanes = True
    End If
    If Not previousSheet Is Nothing Then previousSheet.Activate
    On Error GoTo 0
End Sub
