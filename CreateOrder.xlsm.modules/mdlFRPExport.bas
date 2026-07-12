Attribute VB_Name = "mdlFRPExport"
' ===================================================================
' Module mdlFRPExport (Universal)
' Version: 3.2.0
' Date: 29.06.2026
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Description: Р¤РѕСЂРјРёСЂСѓРµС‚ Excel-РІС‹РіСЂСѓР·РєРё РґР»СЏ РђР»СѓС€С‚С‹/Р¤Р Рџ Рё
' РґРѕРїРѕР»РЅСЏРµС‚ РёС… РєР°РґСЂРѕРІРѕР№ РґРёР°РіРЅРѕСЃС‚РёРєРѕР№ РїРѕ Р»РёСЃС‚Сѓ "РЁС‚Р°С‚".
' ===================================================================
Option Explicit

Private Const COLOR_EXPIRED As Long = 13158655
Private Const COLOR_DATE_WARNING As Long = 10284031
Private Const COLOR_ORG_WARNING As Long = 10092543
Private Const COLOR_MISSING_STAFF As Long = 14277081
Private Const HEADER_COLOR As Long = 13158600
Private Const STAFF_COL_EVENT_TYPE As Long = 22
Private Const STAFF_COL_EVENT_REASON As Long = 23
Private Const STAFF_COL_SERVICE_START As Long = 24
Private Const STAFF_COL_PERSONNEL_SECTION_1 As Long = 36
Private Const STAFF_COL_PERSONNEL_SECTION_2 As Long = 37
Private Const LOCALIZATION_SHEET_NAME As String = "Localization"
Private Const LOCALIZATION_LANGUAGE_CODE As String = "ru"

Public Sub ExportPeriodsToExcel_WithChoice()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim choice As VbMsgBoxResult
    choice = MsgBox(LocalizeExportText("export.periods.choice.prompt", "Choose report type:") & vbCrLf & vbCrLf & _
                    LocalizeExportText("export.periods.choice.dso", "Yes - DSO report") & vbCrLf & _
                    LocalizeExportText("export.periods.choice.risk", "No - FRP Risk report") & vbCrLf & _
                    LocalizeExportText("export.periods.choice.cancel", "Cancel - exit"), _
                    vbYesNoCancel + vbQuestion, LocalizeExportText("export.periods.choice.title", "Report type"))

    If choice = vbYes Then
        Call CreateExcelReportPeriodsByLichniyNomer
    ElseIf choice = vbNo Then
        Call CreateRiskExcelReport
    End If
End Sub

Sub CreateExcelReportPeriodsByLichniyNomer()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim wbNew As Workbook, wsNew As Worksheet
    Dim wsMain As Worksheet, wsStaff As Worksheet
    Dim i As Long, j As Long, outputRow As Long, lastRowMain As Long
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim colTableNumber As Long

    Dim uniquePersons As Collection, personData As Collection, periodList As Collection
    Dim periodArr() As Variant, cutoffDate As Date, filePath As String
    Dim staffColumnMap As Object
    Dim primaryOrganization As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = LocalizeExportText("export.dso.statusbar", "Building DSO report...")

    Set wsMain = mdlHelper.GetDsoWorksheet()
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsMain Is Nothing Or wsStaff Is Nothing Then
        MsgBox LocalizeExportText("export.staff.columns_missing", "Failed to detect required sheets."), vbCritical, LocalizeExportText("common.error", "Error")
        GoTo CleanUp
    End If

    If Not mdlHelper.FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast) Then
        MsgBox LocalizeExportText("export.staff.columns_missing", "Failed to detect required columns on Staff sheet."), vbCritical, LocalizeExportText("common.error", "Error")
        GoTo CleanUp
    End If

    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)
    Set staffColumnMap = BuildStaffReportColumnMap(wsStaff)

    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "C").End(xlUp).Row

    Set uniquePersons = New Collection
    For i = 2 To lastRowMain
        Dim currentLichniyNomer As String
        currentLichniyNomer = Trim(CStr(wsMain.Cells(i, 3).Value))
        If currentLichniyNomer <> "" Then
            On Error Resume Next
            Set personData = uniquePersons(currentLichniyNomer)
            If Err.Number <> 0 Then
                Set personData = New Collection
                personData.Add currentLichniyNomer, "lichniyNomer"
                personData.Add Trim(CStr(wsMain.Cells(i, 2).Value)), "fio"
                personData.Add vbNullString, "tableNumber"

                Set periodList = New Collection
                personData.Add periodList, "periods"
                uniquePersons.Add personData, currentLichniyNomer
            End If
            Err.Clear
            On Error GoTo ErrorHandler

            mdlHelper.CollectAllPersonPeriods wsMain, i, personData("periods")
        End If
    Next i

    primaryOrganization = DetectPrimaryOrganization(uniquePersons, wsStaff, colLichniyNomer, colTableNumber, staffColumnMap)

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = BuildSafeWorksheetName(LocalizeExportText("export.dso.sheet_name", "DSO Report"))

    WriteDsoHeaders wsNew
    outputRow = 2

    Dim infoRow As Long
    For infoRow = 1 To uniquePersons.Count
        Dim staffInfo As Object

        Set personData = uniquePersons(infoRow)
        Set periodList = personData("periods")
        Set staffInfo = GetStaffExportInfo(wsStaff, CStr(personData("lichniyNomer")), colLichniyNomer, colTableNumber, staffColumnMap)

        If periodList.Count > 0 Then
            ReDim periodArr(1 To periodList.Count, 1 To 3)
            For j = 1 To periodList.Count
                periodArr(j, 1) = periodList(j)(1)
                periodArr(j, 2) = periodList(j)(2)
                periodArr(j, 3) = periodList(j)(3)
            Next j

            Call SortArray(periodArr)

            Dim totalDays As Long, totalRestDays As Long, restDaysArr() As Long
            totalDays = 0
            For j = 1 To UBound(periodArr)
                totalDays = totalDays + periodArr(j, 3)
            Next j
            totalRestDays = Int(totalDays / 3) * 2

            ReDim restDaysArr(1 To UBound(periodArr))
            Dim restBase As Long, restExtra As Long
            If periodList.Count > 0 Then
                restBase = totalRestDays \ periodList.Count
                restExtra = totalRestDays Mod periodList.Count
                For j = 1 To periodList.Count
                    restDaysArr(j) = restBase
                    If restExtra > 0 Then
                        restDaysArr(j) = restDaysArr(j) + 1
                        restExtra = restExtra - 1
                    End If
                    If restDaysArr(j) = 0 And totalRestDays > 0 Then restDaysArr(j) = 1
                Next j
            End If

            For j = 1 To UBound(periodArr)
                wsNew.Cells(outputRow, 1).Value = outputRow - 1
                wsNew.Cells(outputRow, 2).Value = personData("fio")
                wsNew.Cells(outputRow, 3).Value = personData("lichniyNomer")
                wsNew.Cells(outputRow, 4).Value = staffInfo("tableNumber")
                wsNew.Cells(outputRow, 5).Value = periodArr(j, 1)
                wsNew.Cells(outputRow, 6).Value = periodArr(j, 2)
                wsNew.Cells(outputRow, 7).Value = periodArr(j, 3)
                wsNew.Cells(outputRow, 8).Value = restDaysArr(j)
                wsNew.Cells(outputRow, 9).Value = IIf(periodArr(j, 2) >= cutoffDate, LocalizeExportText("common.yes", "Yes"), LocalizeExportText("common.no", "No"))
                wsNew.Cells(outputRow, 10).Value = staffInfo("eventType")
                wsNew.Cells(outputRow, 11).Value = staffInfo("eventReason")
                wsNew.Cells(outputRow, 12).Value = staffInfo("serviceStartText")
                wsNew.Cells(outputRow, 13).Value = staffInfo("personnelSection")
                wsNew.Cells(outputRow, 14).Value = BuildValidationMessage(periodArr(j, 1), staffInfo, primaryOrganization)

                ApplyValidationFormatting wsNew, outputRow, periodArr(j, 1), periodArr(j, 2), cutoffDate, staffInfo, primaryOrganization, False
                outputRow = outputRow + 1
            Next j
        End If
    Next infoRow

    FormatExportSheet wsNew, "A1:N1", "A:N"

    filePath = ThisWorkbook.Path & "\" & LocalizeExportText("export.dso.file_prefix", "Export_DSO_") & Format(Date, "dd.mm.yyyy") & ".xlsx"
    Application.DisplayAlerts = False
    If Dir(filePath) <> "" Then Kill filePath
    wbNew.SaveAs filePath
    wbNew.Close False
    Set wbNew = Nothing

    MsgBox Replace$(LocalizeExportText("export.dso.success", "DSO report created: {path}"), "{path}", filePath), vbInformation, LocalizeExportText("common.ok", "OK")
    GoTo CleanUp

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    MsgBox Replace$(Replace$(LocalizeExportText("export.generic.error", "Error {num}: {desc}"), "{num}", CStr(errNum)), "{desc}", errDesc), vbCritical, LocalizeExportText("common.error", "Error")
CleanUp:
    If Not wbNew Is Nothing Then
        On Error Resume Next
        wbNew.Close False
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub CreateRiskExcelReport()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim wbNew As Workbook, wsNew As Worksheet
    Dim wsDSO As Worksheet, wsStaff As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim lichniyNomer As String, fio As String
    Dim rawPeriods() As mdlRiskExport.RiskPeriod
    Dim splitPeriods() As mdlRiskExport.RiskPeriod
    Dim periodCount As Long, k As Long

    Dim colLichniyNomer As Long, colTableNumber As Long
    Dim staffColumnMap As Object
    Dim primaryOrganization As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = LocalizeExportText("export.risk.statusbar", "Building FRP Risk report...")

    Set wsDSO = mdlHelper.GetDsoWorksheet()
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    If wsDSO Is Nothing Or wsStaff Is Nothing Then
        MsgBox LocalizeExportText("export.staff.columns_missing", "Failed to detect required sheets."), vbCritical, LocalizeExportText("common.error", "Error")
        GoTo CleanUp
    End If

    colLichniyNomer = mdlHelper.colLichniyNomer_Global
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)
    Set staffColumnMap = BuildStaffReportColumnMap(wsStaff)

    lastRow = wsDSO.Cells(wsDSO.Rows.Count, 3).End(xlUp).Row
    primaryOrganization = DetectPrimaryOrganizationFromDsoSheet(wsDSO, wsStaff, lastRow, colLichniyNomer, colTableNumber, staffColumnMap)

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = BuildSafeWorksheetName(LocalizeExportText("export.risk.sheet_name", "FRP Risk"))

    WriteRiskHeaders wsNew
    outputRow = 2

    For i = 2 To lastRow
        Dim staffInfo As Object

        lichniyNomer = Trim(CStr(wsDSO.Cells(i, 3).Value))
        fio = Trim(CStr(wsDSO.Cells(i, 2).Value))

        If lichniyNomer <> "" Then
            Set staffInfo = GetStaffExportInfo(wsStaff, lichniyNomer, colLichniyNomer, colTableNumber, staffColumnMap)

            periodCount = CollectRawRiskPeriods_Local(wsDSO, i, rawPeriods)

            If periodCount > 0 Then
                Dim splitCount As Long
                splitCount = SplitPeriodsByMonth_SeparateRows(rawPeriods, periodCount, splitPeriods)

                For k = 1 To splitCount
                    wsNew.Cells(outputRow, 1).Value = outputRow - 1
                    wsNew.Cells(outputRow, 2).Value = fio
                    wsNew.Cells(outputRow, 3).Value = lichniyNomer
                    wsNew.Cells(outputRow, 4).Value = staffInfo("tableNumber")
                    wsNew.Cells(outputRow, 5).Value = splitPeriods(k).StartDate
                    wsNew.Cells(outputRow, 6).Value = splitPeriods(k).EndDate
                    wsNew.Cells(outputRow, 7).Value = splitPeriods(k).daysCount
                    wsNew.Cells(outputRow, 8).Value = splitPeriods(k).PercentValue & "%"
                    wsNew.Cells(outputRow, 9).Value = IIf(splitPeriods(k).IsExpired, LocalizeExportText("common.no", "No"), LocalizeExportText("common.yes", "Yes"))
                    wsNew.Cells(outputRow, 10).Value = staffInfo("eventType")
                    wsNew.Cells(outputRow, 11).Value = staffInfo("eventReason")
                    wsNew.Cells(outputRow, 12).Value = staffInfo("serviceStartText")
                    wsNew.Cells(outputRow, 13).Value = staffInfo("personnelSection")
                    wsNew.Cells(outputRow, 14).Value = BuildValidationMessage(splitPeriods(k).StartDate, staffInfo, primaryOrganization, splitPeriods(k).IsExpired)

                    ApplyValidationFormatting wsNew, outputRow, splitPeriods(k).StartDate, splitPeriods(k).EndDate, 0, staffInfo, primaryOrganization, splitPeriods(k).IsExpired
                    outputRow = outputRow + 1
                Next k
            End If
        End If
    Next i

    FormatExportSheet wsNew, "A1:N1", "A:N"

    Dim filePathRisk As String
    filePathRisk = ThisWorkbook.Path & "\" & LocalizeExportText("export.risk.file_prefix", "Export_Risk_") & Format(Date, "dd.mm.yyyy") & ".xlsx"

    Application.DisplayAlerts = False
    If Dir(filePathRisk) <> "" Then Kill filePathRisk

    wbNew.SaveAs filePathRisk
    wbNew.Close False
    Set wbNew = Nothing

    MsgBox Replace$(LocalizeExportText("export.risk.success", "FRP Risk report created: {path}"), "{path}", filePathRisk), vbInformation, LocalizeExportText("common.ok", "OK")
    GoTo CleanUp

ErrorHandler:
    Dim errNum2 As Long, errDesc2 As String
    errNum2 = Err.Number
    errDesc2 = Err.Description
    MsgBox Replace$(Replace$(LocalizeExportText("export.risk.error", "Risk export error {num}: {desc}"), "{num}", CStr(errNum2)), "{desc}", errDesc2), vbCritical, LocalizeExportText("common.error", "Error")
    Resume CleanUp
CleanUp:
    If Not wbNew Is Nothing Then
        On Error Resume Next
        wbNew.Close False
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Private Sub WriteDsoHeaders(ByVal wsNew As Worksheet)
    wsNew.Cells(1, 1).Value = LocalizeExportText("export.header.number", "No.")
    wsNew.Cells(1, 2).Value = LocalizeExportText("export.header.fio", "Full name")
    wsNew.Cells(1, 3).Value = LocalizeExportText("export.header.personal_number", "Personal number")
    wsNew.Cells(1, 4).Value = LocalizeExportText("export.header.table_number", "Table number")
    wsNew.Cells(1, 5).Value = LocalizeExportText("export.header.period_start", "Period start")
    wsNew.Cells(1, 6).Value = LocalizeExportText("export.header.period_end", "Period end")
    wsNew.Cells(1, 7).Value = LocalizeExportText("export.header.duration_days", "Duration, days")
    wsNew.Cells(1, 8).Value = LocalizeExportText("export.header.rest_days", "Rest days")
    wsNew.Cells(1, 9).Value = LocalizeExportText("export.header.is_actual", "Actual")
    wsNew.Cells(1, 10).Value = LocalizeExportText("export.header.event_type", "Event type")
    wsNew.Cells(1, 11).Value = LocalizeExportText("export.header.event_reason", "Event reason")
    wsNew.Cells(1, 12).Value = LocalizeExportText("export.header.service_start", "Service start")
    wsNew.Cells(1, 13).Value = LocalizeExportText("export.header.personnel_section", "Personnel section")
    wsNew.Cells(1, 14).Value = LocalizeExportText("export.header.validation", "Validation")
End Sub

Private Sub WriteRiskHeaders(ByVal wsNew As Worksheet)
    wsNew.Cells(1, 1).Value = LocalizeExportText("export.header.number", "No.")
    wsNew.Cells(1, 2).Value = LocalizeExportText("export.header.fio", "Full name")
    wsNew.Cells(1, 3).Value = LocalizeExportText("export.header.personal_number", "Personal number")
    wsNew.Cells(1, 4).Value = LocalizeExportText("export.header.table_number", "Table number")
    wsNew.Cells(1, 5).Value = LocalizeExportText("export.header.period_start", "Period start")
    wsNew.Cells(1, 6).Value = LocalizeExportText("export.header.period_end", "Period end")
    wsNew.Cells(1, 7).Value = LocalizeExportText("export.header.days", "Days")
    wsNew.Cells(1, 8).Value = LocalizeExportText("export.header.percent", "Percent")
    wsNew.Cells(1, 9).Value = LocalizeExportText("export.header.is_actual", "Actual")
    wsNew.Cells(1, 10).Value = LocalizeExportText("export.header.event_type", "Event type")
    wsNew.Cells(1, 11).Value = LocalizeExportText("export.header.event_reason", "Event reason")
    wsNew.Cells(1, 12).Value = LocalizeExportText("export.header.service_start", "Service start")
    wsNew.Cells(1, 13).Value = LocalizeExportText("export.header.personnel_section", "Personnel section")
    wsNew.Cells(1, 14).Value = LocalizeExportText("export.header.validation", "Validation")
End Sub

Private Sub FormatExportSheet(ByVal wsNew As Worksheet, ByVal headerRangeAddress As String, ByVal columnsAddress As String)
    With wsNew.Range(headerRangeAddress)
        .Font.Bold = True
        .Interior.Color = HEADER_COLOR
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    wsNew.Columns("E:F").NumberFormat = "dd.mm.yyyy"
    wsNew.Columns("L:L").NumberFormat = "dd.mm.yyyy"
    wsNew.Columns(columnsAddress).AutoFit
End Sub

Private Function BuildStaffReportColumnMap(ByVal wsStaff As Worksheet) As Object
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.CompareMode = 1
    map("eventType") = STAFF_COL_EVENT_TYPE
    map("eventReason") = STAFF_COL_EVENT_REASON
    map("serviceStart") = STAFF_COL_SERVICE_START
    map("personnelSection1") = STAFF_COL_PERSONNEL_SECTION_1
    map("personnelSection2") = STAFF_COL_PERSONNEL_SECTION_2

    Set BuildStaffReportColumnMap = map
End Function

Private Function GetStaffExportInfo(ByVal wsStaff As Worksheet, ByVal lichniyNomer As String, ByVal colLichniyNomer As Long, ByVal colTableNumber As Long, ByVal staffColumnMap As Object) As Object
    Dim info As Object
    Dim staffRow As Long

    Set info = CreateObject("Scripting.Dictionary")
    info.CompareMode = 1
    info("found") = False
    info("tableNumber") = vbNullString
    info("eventType") = vbNullString
    info("eventReason") = vbNullString
    info("serviceStartDate") = 0
    info("serviceStartText") = vbNullString
    info("personnelSection") = vbNullString

    staffRow = mdlHelper.FindStaffRow(wsStaff, lichniyNomer, colLichniyNomer)
    If staffRow <= 0 Then
        Set GetStaffExportInfo = info
        Exit Function
    End If

    info("found") = True

    If colTableNumber > 0 Then
        info("tableNumber") = Trim(CStr(wsStaff.Cells(staffRow, colTableNumber).Value))
    End If
    If staffColumnMap("eventType") > 0 Then
        info("eventType") = Trim(CStr(wsStaff.Cells(staffRow, staffColumnMap("eventType")).Value))
    End If
    If staffColumnMap("eventReason") > 0 Then
        info("eventReason") = Trim(CStr(wsStaff.Cells(staffRow, staffColumnMap("eventReason")).Value))
    End If
    If staffColumnMap("serviceStart") > 0 Then
        info("serviceStartText") = Trim(CStr(wsStaff.Cells(staffRow, staffColumnMap("serviceStart")).Text))
        info("serviceStartDate") = mdlHelper.ParseDateSafe(wsStaff.Cells(staffRow, staffColumnMap("serviceStart")).Value)
    End If

    info("personnelSection") = ResolvePersonnelSection(wsStaff, staffRow, staffColumnMap)
    Set GetStaffExportInfo = info
End Function

Private Function ResolvePersonnelSection(ByVal wsStaff As Worksheet, ByVal staffRow As Long, ByVal staffColumnMap As Object) As String
    Dim sectionText As String

    sectionText = GetTrimmedCellValue(wsStaff, staffRow, CLng(staffColumnMap("personnelSection2")))
    If sectionText = "" Then
        sectionText = GetTrimmedCellValue(wsStaff, staffRow, CLng(staffColumnMap("personnelSection1")))
    End If

    ResolvePersonnelSection = sectionText
End Function

Private Function DetectPrimaryOrganization(ByVal uniquePersons As Collection, ByVal wsStaff As Worksheet, ByVal colLichniyNomer As Long, ByVal colTableNumber As Long, ByVal staffColumnMap As Object) As String
    Dim counts As Object
    Dim idx As Long

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 1

    For idx = 1 To uniquePersons.Count
        Dim personInfo As Collection
        Dim info As Object
        Dim key As String

        Set personInfo = uniquePersons(idx)
        Set info = GetStaffExportInfo(wsStaff, CStr(personInfo("lichniyNomer")), colLichniyNomer, colTableNumber, staffColumnMap)
        key = NormalizeComparisonValue(CStr(info("personnelSection")))
        If key <> "" Then
            If counts.Exists(key) Then
                counts(key) = CLng(counts(key)) + 1
            Else
                counts.Add key, 1
            End If
        End If
    Next idx

    DetectPrimaryOrganization = GetMostCommonDictionaryValue(counts)
End Function

Private Function DetectPrimaryOrganizationFromDsoSheet(ByVal wsDSO As Worksheet, ByVal wsStaff As Worksheet, ByVal lastRow As Long, ByVal colLichniyNomer As Long, ByVal colTableNumber As Long, ByVal staffColumnMap As Object) As String
    Dim counts As Object
    Dim i As Long

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = 1

    For i = 2 To lastRow
        Dim lichniyNomer As String
        Dim info As Object
        Dim key As String

        lichniyNomer = Trim(CStr(wsDSO.Cells(i, 3).Value))
        If lichniyNomer <> "" Then
            Set info = GetStaffExportInfo(wsStaff, lichniyNomer, colLichniyNomer, colTableNumber, staffColumnMap)
            key = NormalizeComparisonValue(CStr(info("personnelSection")))
            If key <> "" Then
                If counts.Exists(key) Then
                    counts(key) = CLng(counts(key)) + 1
                Else
                    counts.Add key, 1
                End If
            End If
        End If
    Next i

    DetectPrimaryOrganizationFromDsoSheet = GetMostCommonDictionaryValue(counts)
End Function

Private Function GetMostCommonDictionaryValue(ByVal counts As Object) As String
    Dim item As Variant
    Dim bestCount As Long

    GetMostCommonDictionaryValue = vbNullString

    For Each item In counts.Keys
        If CLng(counts(item)) > bestCount Then
            bestCount = CLng(counts(item))
            GetMostCommonDictionaryValue = CStr(item)
        End If
    Next item
End Function

Private Function BuildValidationMessage(ByVal periodStart As Date, ByVal staffInfo As Object, ByVal primaryOrganization As String, Optional ByVal isExpired As Boolean = False) As String
    Dim messages As Collection
    Dim detailText As String
    Set messages = New Collection

    If isExpired Then messages.Add LocalizeExportText("export.validation.expired", "Period is outdated")

    If Not CBool(staffInfo("found")) Then
        messages.Add LocalizeExportText("export.validation.staff_missing", "Employee not found on Staff sheet")
    Else
        If CDateSafe(staffInfo("serviceStartDate")) > 0 Then
            If periodStart < CDateSafe(staffInfo("serviceStartDate")) Then
                detailText = LocalizeExportText("export.validation.before_enrollment_detail", _
                                                "Period starts before enrollment ({date})")
                detailText = Replace$(detailText, "{date}", CStr(staffInfo("serviceStartText")))
                messages.Add detailText
            End If
        End If

        If primaryOrganization <> "" And NormalizeComparisonValue(CStr(staffInfo("personnelSection"))) <> "" Then
            If NormalizeComparisonValue(CStr(staffInfo("personnelSection"))) <> NormalizeComparisonValue(primaryOrganization) Then
                detailText = LocalizeExportText("export.validation.other_section_detail", _
                                                "Employee is in another personnel section ({section})")
                detailText = Replace$(detailText, "{section}", CStr(staffInfo("personnelSection")))
                messages.Add detailText
            End If
        End If
    End If

    If messages.Count = 0 Then
        BuildValidationMessage = LocalizeExportText("export.validation.ok", "OK")
    Else
        BuildValidationMessage = JoinCollection(messages, "; ")
    End If
End Function

Private Function LocalizeExportText(ByVal localizationKey As String, ByVal fallback As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim languageCode As String

    On Error GoTo SafeFallback

    Set ws = ThisWorkbook.Worksheets(LOCALIZATION_SHEET_NAME)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For colIndex = 2 To lastCol
        languageCode = LCase$(Trim$(CStr(ws.Cells(1, colIndex).Value)))
        If languageCode = LOCALIZATION_LANGUAGE_CODE Then
            For rowIndex = 2 To lastRow
                If LCase$(Trim$(CStr(ws.Cells(rowIndex, 1).Value))) = LCase$(Trim$(localizationKey)) Then
                    If Len(CStr(ws.Cells(rowIndex, colIndex).Value)) > 0 Then
                        LocalizeExportText = CStr(ws.Cells(rowIndex, colIndex).Value)
                        If Not IsLikelyBrokenExportText(LocalizeExportText) Then
                            Exit Function
                        End If
                        LocalizeExportText = vbNullString
                    End If
                End If
            Next rowIndex
            Exit For
        End If
    Next colIndex

SafeFallback:
    If Len(LocalizeExportText) = 0 Then
        If Not IsLikelyBrokenExportText(fallback) Then
            LocalizeExportText = fallback
        Else
            LocalizeExportText = localizationKey
        End If
    End If
End Function

Private Function IsLikelyBrokenExportText(ByVal value As String) As Boolean
    Dim sample As String

    sample = Trim$(value)
    If Len(sample) = 0 Then Exit Function

    IsLikelyBrokenExportText = _
        InStr(1, sample, "РїС—", vbTextCompare) > 0 Or _
        InStr(1, sample, "Гђ", vbTextCompare) > 0 Or _
        InStr(1, sample, "Г‘", vbTextCompare) > 0 Or _
        InStr(1, sample, "пїЅпїЅпїЅпїЅ", vbTextCompare) > 0
End Function

Private Sub ApplyValidationFormatting(ByVal wsNew As Worksheet, ByVal rowNum As Long, ByVal periodStart As Date, ByVal periodEnd As Date, ByVal cutoffDate As Date, ByVal staffInfo As Object, ByVal primaryOrganization As String, ByVal isExpired As Boolean)
    Dim hasDateWarning As Boolean
    Dim hasOrgWarning As Boolean

    If isExpired Then
        wsNew.Range("A" & rowNum & ":N" & rowNum).Interior.Color = COLOR_EXPIRED
    ElseIf cutoffDate > 0 And periodEnd < cutoffDate Then
        wsNew.Range("A" & rowNum & ":N" & rowNum).Interior.Color = COLOR_EXPIRED
    End If

    If Not CBool(staffInfo("found")) Then
        wsNew.Range("J" & rowNum & ":N" & rowNum).Interior.Color = COLOR_MISSING_STAFF
        Exit Sub
    End If

    If CDateSafe(staffInfo("serviceStartDate")) > 0 Then
        If periodStart < CDateSafe(staffInfo("serviceStartDate")) Then
            hasDateWarning = True
            wsNew.Cells(rowNum, 5).Interior.Color = COLOR_DATE_WARNING
            wsNew.Cells(rowNum, 12).Interior.Color = COLOR_DATE_WARNING
            wsNew.Cells(rowNum, 14).Interior.Color = COLOR_DATE_WARNING
        End If
    End If

    If primaryOrganization <> "" And NormalizeComparisonValue(CStr(staffInfo("personnelSection"))) <> "" Then
        If NormalizeComparisonValue(CStr(staffInfo("personnelSection"))) <> NormalizeComparisonValue(primaryOrganization) Then
            hasOrgWarning = True
            wsNew.Cells(rowNum, 13).Interior.Color = COLOR_ORG_WARNING
            wsNew.Cells(rowNum, 14).Interior.Color = COLOR_ORG_WARNING
        End If
    End If

    If hasDateWarning And hasOrgWarning Then
        wsNew.Cells(rowNum, 14).Interior.Color = COLOR_ORG_WARNING
    End If
End Sub

Private Function CDateSafe(ByVal value As Variant) As Date
    On Error Resume Next
    CDateSafe = 0
    If IsDate(value) Then CDateSafe = CDate(value)
    On Error GoTo 0
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal separator As String) As String
    Dim i As Long
    For i = 1 To items.Count
        If JoinCollection = "" Then
            JoinCollection = CStr(items(i))
        Else
            JoinCollection = JoinCollection & separator & CStr(items(i))
        End If
    Next i
End Function

Private Function NormalizeComparisonValue(ByVal rawValue As String) As String
    NormalizeComparisonValue = LCase$(Trim$(Replace(Replace(rawValue, vbCr, " "), vbLf, " ")))
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String, Optional ByVal occurrence As Long = 1) As Long
    Dim lastCol As Long, i As Long, currentOccurrence As Long
    Dim candidate As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        candidate = Trim$(CStr(ws.Cells(1, i).Value))
        If LCase$(candidate) = LCase$(headerText) Then
            currentOccurrence = currentOccurrence + 1
            If currentOccurrence = occurrence Then
                FindHeaderColumn = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function BuildSafeWorksheetName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    BuildSafeWorksheetName = Left$(rawName, 31)
    invalidChars = Array("[", "]", ":", "*", "?", "/", "\")
    For Each item In invalidChars
        BuildSafeWorksheetName = Replace$(BuildSafeWorksheetName, CStr(item), "_")
    Next item

    If Len(Trim$(BuildSafeWorksheetName)) = 0 Then
        BuildSafeWorksheetName = "Report"
    End If
End Function

Private Function GetTrimmedCellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As String
    If colNum <= 0 Then Exit Function
    GetTrimmedCellValue = Trim$(CStr(ws.Cells(rowNum, colNum).Value))
End Function

Private Function CollectRawRiskPeriods_Local(ws As Worksheet, rowNum As Long, ByRef periods() As mdlRiskExport.RiskPeriod) As Long
    Dim lastCol As Long, j As Long, pCount As Long
    pCount = 0
    lastCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column

    ReDim periods(1 To 50)
    Dim expirationDate As Date
    expirationDate = DateAdd("m", -42, Date)

    For j = 5 To lastCol Step 2
        Dim sVal As Variant, eVal As Variant
        sVal = ws.Cells(rowNum, j).Text
        eVal = ws.Cells(rowNum, j + 1).Text

        Dim sDate As Date, eDate As Date
        sDate = mdlHelper.ParseDateSafe(sVal)
        eDate = mdlHelper.ParseDateSafe(eVal)

        If sDate > 0 And eDate > 0 Then
            If sDate <= eDate Then
                pCount = pCount + 1
                periods(pCount).StartDate = sDate
                periods(pCount).EndDate = eDate
                periods(pCount).IsExpired = (sDate < expirationDate)
            End If
        End If
    Next j

    CollectRawRiskPeriods_Local = pCount
End Function

Private Function SplitPeriodsByMonth_SeparateRows(ByRef rawPeriods() As mdlRiskExport.RiskPeriod, ByVal rawCount As Long, ByRef splitPeriods() As mdlRiskExport.RiskPeriod) As Long
    Dim i As Long, count As Long
    count = 0

    Dim tempSplit() As mdlRiskExport.RiskPeriod
    ReDim tempSplit(1 To rawCount * 10)

    For i = 1 To rawCount
        Dim curDate As Date
        curDate = rawPeriods(i).StartDate

        Do While curDate <= rawPeriods(i).EndDate
            Dim monthEnd As Date
            monthEnd = DateSerial(Year(curDate), Month(curDate) + 1, 0)

            Dim segEnd As Date
            If rawPeriods(i).EndDate < monthEnd Then
                segEnd = rawPeriods(i).EndDate
            Else
                segEnd = monthEnd
            End If

            count = count + 1
            tempSplit(count).StartDate = curDate
            tempSplit(count).EndDate = segEnd
            tempSplit(count).daysCount = DateDiff("d", curDate, segEnd) + 1
            tempSplit(count).MonthYear = Format(curDate, "yyyymm")
            tempSplit(count).IsExpired = rawPeriods(i).IsExpired

            curDate = monthEnd + 1
        Loop
    Next i

    If count = 0 Then
        SplitPeriodsByMonth_SeparateRows = 0
        Exit Function
    End If

    Dim j As Long
    Dim temp As mdlRiskExport.RiskPeriod
    For i = 1 To count - 1
        For j = i + 1 To count
            If tempSplit(i).StartDate > tempSplit(j).StartDate Then
                temp = tempSplit(i)
                tempSplit(i) = tempSplit(j)
                tempSplit(j) = temp
            End If
        Next j
    Next i

    Dim monthlyAccumulator As Object
    Set monthlyAccumulator = CreateObject("Scripting.Dictionary")

    For i = 1 To count
        Dim key As String
        key = tempSplit(i).MonthYear

        Dim currentAccumulated As Double
        If monthlyAccumulator.Exists(key) Then
            currentAccumulated = monthlyAccumulator(key)
        Else
            currentAccumulated = 0
        End If

        Dim periodValue As Double
        periodValue = tempSplit(i).daysCount * 2

        Dim remainingLimit As Double
        remainingLimit = 60 - currentAccumulated
        If remainingLimit < 0 Then remainingLimit = 0

        Dim finalPercent As Double
        If periodValue <= remainingLimit Then
            finalPercent = periodValue
        Else
            finalPercent = remainingLimit
        End If

        tempSplit(i).PercentValue = finalPercent

        If monthlyAccumulator.Exists(key) Then
            monthlyAccumulator(key) = currentAccumulated + finalPercent
        Else
            monthlyAccumulator.Add key, finalPercent
        End If
    Next i

    ReDim splitPeriods(1 To count)
    For i = 1 To count
        splitPeriods(i) = tempSplit(i)
    Next i

    SplitPeriodsByMonth_SeparateRows = count
End Function

Private Sub SortArray(ByRef arr As Variant)
    Dim i As Long, j As Long, temp1, temp2, temp3
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i, 1) > arr(j, 1) Then
                temp1 = arr(i, 1): temp2 = arr(i, 2): temp3 = arr(i, 3)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3)
                arr(j, 1) = temp1: arr(j, 2) = temp2: arr(j, 3) = temp3
            End If
        Next j
    Next i
End Sub
