Attribute VB_Name = "mdlEnrollmentOrderExport"
Option Explicit

Private Const WD_FORMAT_XML_DOCUMENT As Long = 16
Private Const DEFAULT_PREMIUM_PERCENT As String = "25"
Private Const DEFAULT_POSITION_ALLOWANCE_PERCENT As String = "100"
Private Const DEFAULT_SPECIAL_CONDITIONS_PERCENT As String = "50"
Private Const DEFAULT_TARIFF_PERCENT As String = "50"
Private Const DEFAULT_CONTRACT430_PERCENT As String = "60"
Private Const ENROLLMENT_BODY_MARKER_DEFAULT As String = "[ENROLLMENT_ORDER_BODY]"
Public Sub EnsureEnrollmentTemplateAvailable()
    Dim templatePath As String
    Dim legacyTemplatePath As String
    Dim fallbackTemplate As String

    templatePath = ThisWorkbook.Path & "\" & EnrollmentTemplateFileName()
    If Dir$(templatePath) <> "" Then Exit Sub

    legacyTemplatePath = ThisWorkbook.Path & "\EnrollmentOrderTemplate.docx"
    If Dir$(legacyTemplatePath) <> "" Then
        On Error Resume Next
        FileCopy legacyTemplatePath, templatePath
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    End If

    fallbackTemplate = mdlPaymentTypes.GetTemplatePath(mdlPaymentTypes.DEFAULT_TEMPLATE)
    If fallbackTemplate <> "" Then
        On Error Resume Next
        FileCopy fallbackTemplate, templatePath
        On Error GoTo 0
    End If
End Sub

Public Function BuildPreviewPayload(ByVal record As Object) As Object
    Dim payload As Object

    Set payload = CreateObject("Scripting.Dictionary")
    payload.CompareMode = vbTextCompare

    payload("standard_text") = BuildStandardMonthlyText(record)
    payload("personal_text") = BuildPersonalMonthlyText(record)
    payload("section1_text") = BuildSection1ForRecord(record)
    payload("section2_text") = BuildSection2ForRecord(record)

    Set BuildPreviewPayload = payload
End Function

Public Function ExportEnrollmentOrderByRow(ByVal rowNum As Long) As String
    Dim orderDraftId As String

    orderDraftId = GetOrderDraftIdForRow(rowNum)

    If orderDraftId = "" Then
        ExportEnrollmentOrderByRow = ExportEnrollmentOrderByDraftId("", rowNum)
    Else
        ExportEnrollmentOrderByRow = ExportEnrollmentOrderByDraftId(orderDraftId, rowNum)
    End If
End Function

Public Function IsEnrollmentExportErrorResult(ByVal exportResult As String) As Boolean
    IsEnrollmentExportErrorResult = (UCase$(Left$(Trim$(exportResult), 6)) = "ERROR:")
End Function

Public Function GetEnrollmentExportErrorText(ByVal exportResult As String) As String
    If IsEnrollmentExportErrorResult(exportResult) Then
        GetEnrollmentExportErrorText = Trim$(Mid$(Trim$(exportResult), 7))
    Else
        GetEnrollmentExportErrorText = Trim$(exportResult)
    End If
End Function

Public Function GetOrderDraftIdForRow(ByVal rowNum As Long) As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    GetOrderDraftIdForRow = Trim$(CStr(ws.Cells(rowNum, mdlEnrollmentWorkflow.COL_ENROLLMENT_ORDER_DRAFT_ID).Value))
End Function

Public Function GetExportRowCount(ByVal orderDraftId As String, Optional ByVal fallbackRow As Long = 0) As Long
    GetExportRowCount = CollectExportRows(orderDraftId, fallbackRow).Count
End Function

Public Function GetExportScopeText(ByVal orderDraftId As String, Optional ByVal fallbackRow As Long = 0) As String
    Dim rowsCount As Long

    rowsCount = GetExportRowCount(orderDraftId, fallbackRow)

    If Trim$(orderDraftId) = "" Then
        GetExportScopeText = tf("enrollment.word.scope.single", "Строка {row}", "{row}", fallbackRow)
    Else
        GetExportScopeText = tf("enrollment.word.scope.package", "Пакет {draftId}, строк: {count}", "{draftId}", orderDraftId, "{count}", rowsCount)
    End If
End Function

Public Function GetEnrollmentExportBlockingIssues(ByVal orderDraftId As String, Optional ByVal fallbackRow As Long = 0) As String
    Dim rowsToExport As Collection
    Dim readinessIssues As String
    Dim mismatch As String

    On Error GoTo ErrorHandler

    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    Set rowsToExport = CollectExportRows(orderDraftId, fallbackRow)

    If rowsToExport.Count = 0 Then
        GetEnrollmentExportBlockingIssues = L("enrollment.word.error.no_rows", "Не выбраны строки зачисления для экспорта.")
        Exit Function
    End If

    readinessIssues = BuildReadinessIssuesForRows(rowsToExport)
    If readinessIssues <> "" Then
        GetEnrollmentExportBlockingIssues = L("enrollment.word.error.not_ready", "Невозможно сформировать Word-приказ: в выбранных строках есть незакрытые обязательные замечания.") & vbCrLf & readinessIssues
        Exit Function
    End If

    mismatch = BuildGroupHeaderConflictText(rowsToExport)
    If mismatch <> "" Then
        GetEnrollmentExportBlockingIssues = L("enrollment.word.error.group_conflict", "Невозможно собрать один приказ по OrderDraftId из-за конфликтов:") & vbCrLf & mismatch
    End If
    Exit Function

ErrorHandler:
    GetEnrollmentExportBlockingIssues = Err.Description
End Function

Public Function ExportSelectedEnrollmentPackage() As String
    Dim rowNum As Long
    Dim orderDraftId As String

    rowNum = mdlEnrollmentWorkflow.ResolveActiveEnrollmentRow()
    orderDraftId = GetOrderDraftIdForRow(rowNum)
    ExportSelectedEnrollmentPackage = ExportEnrollmentOrderByDraftId(orderDraftId, rowNum)
End Function

Public Function ExportEnrollmentOrderByDraftId(ByVal orderDraftId As String, Optional ByVal fallbackRow As Long = 0) As String
    Dim rowsToExport As Collection
    Dim groupHeader As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim fullText As String
    Dim rowNum As Variant
    Dim itemIndex As Long
    Dim errNumber As Long
    Dim errDescription As String
    Dim section2Text As String
    Dim blockingIssues As String

    mdlEnrollmentWorkflow.EnsureEnrollmentInfrastructure
    EnsureEnrollmentTemplateAvailable
    Set rowsToExport = CollectExportRows(orderDraftId, fallbackRow)
    If rowsToExport.Count = 0 Then
        ExportEnrollmentOrderByDraftId = BuildExportErrorResult(L("enrollment.word.error.no_rows", "Не выбраны строки зачисления для экспорта."))
        Exit Function
    End If

    blockingIssues = GetEnrollmentExportBlockingIssues(orderDraftId, fallbackRow)
    If blockingIssues <> "" Then
        ExportEnrollmentOrderByDraftId = BuildExportErrorResult(blockingIssues)
        Exit Function
    End If

    Set groupHeader = ValidateGroupHeader(rowsToExport)

    On Error GoTo ErrorHandler

    templatePath = ThisWorkbook.Path & "\" & EnrollmentTemplateFileName()
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    If Dir$(templatePath) <> "" Then
        Set wordDoc = wordApp.Documents.Add(templatePath)
    Else
        Set wordDoc = wordApp.Documents.Add
    End If

    fullText = BuildDocumentHeader(groupHeader) & vbCrLf
    fullText = fullText & BuildSectionCaption(1) & vbCrLf & vbCrLf

    itemIndex = 1
    For Each rowNum In rowsToExport
        fullText = fullText & BuildSection1ForRecord(mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowNum)), itemIndex) & vbCrLf & vbCrLf
        itemIndex = itemIndex + 1
    Next rowNum

    If HasSection2Rows(rowsToExport) Then
        fullText = fullText & vbFormFeed
        fullText = fullText & BuildSectionCaption(2) & vbCrLf
        fullText = fullText & BuildSection2Intro() & vbCrLf & vbCrLf
        itemIndex = 1
        For Each rowNum In rowsToExport
            section2Text = BuildSection2ForRecord(mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowNum)), itemIndex)
            If Trim$(section2Text) <> "" Then
                fullText = fullText & section2Text
                fullText = fullText & vbCrLf & vbCrLf
                itemIndex = itemIndex + 1
            End If
        Next rowNum
    End If

    fullText = fullText & BuildSignatureBlock(groupHeader)
    InsertEnrollmentBodyIntoDocument wordDoc, fullText
    ApplyEnrollmentOrderFormatting wordDoc

    outputPath = BuildOutputFilePath(groupHeader("orderDraftId"))
    mdlHelper.SaveWordDocumentSafe wordDoc, outputPath
    wordDoc.Close False
    wordApp.Quit

    ExportEnrollmentOrderByDraftId = outputPath
    Exit Function

ErrorHandler:
    errNumber = Err.Number
    errDescription = Err.Description

    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0

    If errNumber = vbObjectError + 1701 Or errNumber = vbObjectError + 1702 Or errNumber = vbObjectError + 1703 Then
        ExportEnrollmentOrderByDraftId = BuildExportErrorResult(errDescription)
    ElseIf errNumber > 0 Then
        Err.Raise errNumber, "ExportEnrollmentOrderByDraftId", errDescription
    Else
        Err.Raise vbObjectError + 1899, "ExportEnrollmentOrderByDraftId", errDescription
    End If
End Function

Private Function BuildExportErrorResult(ByVal errorText As String) As String
    BuildExportErrorResult = "ERROR: " & Trim$(errorText)
End Function

Public Function GetEnrollmentTemplateBodyMarker() As String
    GetEnrollmentTemplateBodyMarker = EnrollmentTemplateBodyMarker()
End Function

Private Sub InsertEnrollmentBodyIntoDocument(ByVal wordDoc As Object, ByVal bodyText As String)
    Dim markerText As String
    Dim searchRange As Object
    Dim endRange As Object
    Dim existingText As String
    Dim contentText As String
    Dim markerStart As Long

    markerText = EnrollmentTemplateBodyMarker()
    If markerText <> "" Then
        contentText = wordDoc.Content.Text
        markerStart = InStr(1, contentText, markerText, vbBinaryCompare)
        If markerStart > 0 Then
            Set searchRange = wordDoc.Range(wordDoc.Content.Start + markerStart - 1, wordDoc.Content.Start + markerStart - 1 + Len(markerText))
            searchRange.Text = bodyText
            Exit Sub
        End If

        Set searchRange = wordDoc.Content
        With searchRange.Find
            .ClearFormatting
            .Text = markerText
            .Forward = True
            .Wrap = 0
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            If .Execute Then
                searchRange.Text = bodyText
                Exit Sub
            End If
        End With
    End If

    existingText = Trim$(Replace$(Replace$(wordDoc.Content.Text, ChrW$(13), ""), ChrW$(7), ""))
    If existingText = "" Then
        wordDoc.Range.Text = bodyText
    Else
        Set endRange = wordDoc.Range(wordDoc.Content.End - 1, wordDoc.Content.End - 1)
        endRange.InsertAfter vbCrLf & bodyText
    End If
End Sub

Private Sub ApplyEnrollmentOrderFormatting(ByVal wordDoc As Object)
    Const WD_ALIGN_PARAGRAPH_CENTER As Long = 1
    Const WD_ALIGN_PARAGRAPH_JUSTIFY As Long = 3
    Const WD_LINE_SPACE_SINGLE As Long = 0
    Const CM_1 As Single = 28.35
    Const FIRST_LINE_INDENT_POINTS As Single = 35.4

    Dim paragraph As Object
    Dim paragraphText As String
    Dim nextParagraphText As String
    Dim paragraphIndex As Long
    Dim paragraphCount As Long
    Dim currentSection As Long

    On Error Resume Next

    With wordDoc.PageSetup
        .TopMargin = 2 * CM_1
        .BottomMargin = 2 * CM_1
        .LeftMargin = 3 * CM_1
        .RightMargin = 1.5 * CM_1
    End With

    With wordDoc.Content
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_JUSTIFY
        .ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_SINGLE
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.FirstLineIndent = FIRST_LINE_INDENT_POINTS
    End With

    paragraphCount = wordDoc.Paragraphs.Count
    For Each paragraph In wordDoc.Paragraphs
        paragraphIndex = paragraphIndex + 1
        paragraphText = Trim$(Replace$(Replace$(paragraph.Range.Text, ChrW$(13), ""), ChrW$(7), ""))
        If paragraphIndex < paragraphCount Then
            nextParagraphText = Trim$(Replace$(Replace$(wordDoc.Paragraphs(paragraphIndex + 1).Range.Text, ChrW$(13), ""), ChrW$(7), ""))
        Else
            nextParagraphText = vbNullString
        End If

        If IsSectionCaptionText(paragraphText, 1) Then
            currentSection = 1
        ElseIf IsSectionCaptionText(paragraphText, 2) Then
            currentSection = 2
        End If

        If ShouldCenterEnrollmentParagraph(paragraphText, paragraphIndex, paragraphCount) Then
            paragraph.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_CENTER
            paragraph.Range.ParagraphFormat.FirstLineIndent = 0
        Else
            paragraph.Range.ParagraphFormat.Alignment = WD_ALIGN_PARAGRAPH_JUSTIFY
            paragraph.Range.ParagraphFormat.FirstLineIndent = FIRST_LINE_INDENT_POINTS
        End If

        If paragraphIndex <= 2 Or Left$(paragraphText, 1) = ChrW$(167) Then
            paragraph.Range.Font.Bold = True
        End If

        If ShouldKeepEnrollmentParagraphWithNext(paragraphText, currentSection, nextParagraphText) Then
            paragraph.Range.ParagraphFormat.KeepWithNext = True
        End If
        If ShouldKeepEnrollmentParagraphTogether(paragraphText, currentSection) Then
            paragraph.Range.ParagraphFormat.KeepTogether = True
        End If
    Next paragraph

    On Error GoTo 0
End Sub

Private Function ShouldCenterEnrollmentParagraph(ByVal paragraphText As String, ByVal paragraphIndex As Long, ByVal paragraphCount As Long) As Boolean
    If paragraphText = "" Then Exit Function

    If paragraphIndex <= 5 Then
        ShouldCenterEnrollmentParagraph = True
        Exit Function
    End If

    If Left$(paragraphText, 1) = ChrW$(167) Then
        ShouldCenterEnrollmentParagraph = True
        Exit Function
    End If

    If paragraphIndex >= paragraphCount - 2 Then
        ShouldCenterEnrollmentParagraph = True
        Exit Function
    End If
End Function

Private Function ShouldKeepEnrollmentParagraphWithNext(ByVal paragraphText As String, ByVal currentSection As Long, ByVal nextParagraphText As String) As Boolean
    If paragraphText = "" Then Exit Function

    If Left$(paragraphText, 1) = ChrW$(167) Then
        ShouldKeepEnrollmentParagraphWithNext = True
        Exit Function
    End If

    If paragraphText = BuildSection2Intro() Then
        ShouldKeepEnrollmentParagraphWithNext = True
        Exit Function
    End If

    If currentSection = 2 Then
        If IsNumberedEnrollmentParagraph(paragraphText) Then
            ShouldKeepEnrollmentParagraphWithNext = True
            Exit Function
        End If

        If IsBasisParagraph(paragraphText) And IsSignatureParagraph(nextParagraphText) Then
            ShouldKeepEnrollmentParagraphWithNext = True
            Exit Function
        End If

        If IsSignatureParagraph(paragraphText) And IsSignatureParagraph(nextParagraphText) Then
            ShouldKeepEnrollmentParagraphWithNext = True
            Exit Function
        End If
    End If
End Function

Private Function ShouldKeepEnrollmentParagraphTogether(ByVal paragraphText As String, ByVal currentSection As Long) As Boolean
    If paragraphText = "" Then Exit Function

    If currentSection = 2 Then
        If paragraphText = BuildSection2Intro() Or _
           IsNumberedEnrollmentParagraph(paragraphText) Or _
           IsBasisParagraph(paragraphText) Or _
           IsSignatureParagraph(paragraphText) Then
            ShouldKeepEnrollmentParagraphTogether = True
            Exit Function
        End If
    End If
End Function

Private Function IsSectionCaptionText(ByVal paragraphText As String, ByVal sectionNumber As Long) As Boolean
    IsSectionCaptionText = (paragraphText = BuildSectionCaption(sectionNumber))
End Function

Private Function IsNumberedEnrollmentParagraph(ByVal paragraphText As String) As Boolean
    Dim dotPos As Long

    dotPos = InStr(1, paragraphText, ". ", vbBinaryCompare)
    If dotPos <= 1 Or dotPos > 4 Then Exit Function

    IsNumberedEnrollmentParagraph = IsNumeric(Left$(paragraphText, dotPos - 1))
End Function

Private Function IsBasisParagraph(ByVal paragraphText As String) As Boolean
    Dim basisPrefix As String

    basisPrefix = L("enrollment.word.requisites.basis", mdlHelper.Ru(1054, 1089, 1085, 1086, 1074, 1072, 1085, 1080, 1077, 58))
    IsBasisParagraph = (Left$(paragraphText, Len(basisPrefix)) = basisPrefix)
End Function

Private Function IsSignatureParagraph(ByVal paragraphText As String) As Boolean
    If paragraphText = "" Then Exit Function

    If InStr(1, paragraphText, L("enrollment.word.signature.position_marker", mdlHelper.Ru(1050, 1054, 1052, 1040, 1053, 1044, 1048, 1056, 1040)), vbTextCompare) > 0 Then
        IsSignatureParagraph = True
        Exit Function
    End If

    If InStr(1, paragraphText, L("enrollment.word.signature.rank.major", mdlHelper.Ru(1084, 1072, 1081, 1086, 1088)), vbTextCompare) > 0 Then
        IsSignatureParagraph = True
        Exit Function
    End If
End Function

Private Sub ValidateExportReadiness(ByVal rowsToExport As Collection)
    Dim readinessIssues As String

    readinessIssues = BuildReadinessIssuesForRows(rowsToExport)
    If readinessIssues <> "" Then
        Err.Raise vbObjectError + 1703, "ValidateExportReadiness", L("enrollment.word.error.not_ready", "Невозможно сформировать Word-приказ: в выбранных строках есть незакрытые обязательные замечания.") & vbCrLf & readinessIssues
    End If
End Sub

Private Function BuildReadinessIssuesForRows(ByVal rowsToExport As Collection) As String
    Dim rowNum As Variant
    Dim record As Object
    Dim readinessIssues As String
    Dim rowIssues As String

    For Each rowNum In rowsToExport
        Set record = mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowNum))
        rowIssues = BuildExportReadinessIssue(record, CLng(rowNum))
        If rowIssues <> "" Then
            If readinessIssues <> "" Then readinessIssues = readinessIssues & vbCrLf
            readinessIssues = readinessIssues & rowIssues
        End If
    Next rowNum

    BuildReadinessIssuesForRows = readinessIssues
End Function

Private Function BuildExportReadinessIssue(ByVal record As Object, ByVal rowNum As Long) As String
    Dim issuesText As String
    Dim rowHeader As String

    If NormalizeExportReadyValue(SafeText(record("word_ready"))) = "YES" Then Exit Function

    rowHeader = tf("enrollment.word.error.row_prefix", "Строка {row}", "{row}", rowNum)
    If SafeText(record("fio")) <> "" Then
        rowHeader = rowHeader & " (" & SafeText(record("fio")) & ")"
    End If

    issuesText = SafeText(record("validation_issues"))
    If Trim$(issuesText) = "" Then
        BuildExportReadinessIssue = "- " & rowHeader
    Else
        BuildExportReadinessIssue = "- " & rowHeader & ": " & Replace$(issuesText, vbCrLf, "; ")
    End If
End Function

Private Function NormalizeExportReadyValue(ByVal rawValue As String) As String
    Select Case UCase$(Trim$(rawValue))
        Case "YES", "TRUE", "1", "ДА"
            NormalizeExportReadyValue = "YES"
        Case Else
            NormalizeExportReadyValue = "NO"
    End Select
End Function

Private Function CollectExportRows(ByVal orderDraftId As String, ByVal fallbackRow As Long) As Collection
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim currentDraftId As String

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    Set CollectExportRows = New Collection
    lastRow = ws.Cells(ws.Rows.Count, mdlEnrollmentWorkflow.COL_ENROLLMENT_ORDER_DRAFT_ID).End(xlUp).Row

    If Trim$(orderDraftId) = "" Then
        If fallbackRow >= 2 Then
            CollectExportRows.Add fallbackRow
        End If
        Exit Function
    End If

    For rowNum = 2 To lastRow
        currentDraftId = Trim$(CStr(ws.Cells(rowNum, mdlEnrollmentWorkflow.COL_ENROLLMENT_ORDER_DRAFT_ID).Value))
        If StrComp(currentDraftId, orderDraftId, vbTextCompare) = 0 Then
            CollectExportRows.Add rowNum
        End If
    Next rowNum
End Function

Private Function BuildGroupHeaderConflictText(ByVal rowsToExport As Collection) As String
    Dim firstRecord As Object
    Dim currentRecord As Object
    Dim rowNum As Variant
    Dim mismatch As String
    Dim orderDate As String
    Dim orderNumber As String
    Dim orderIssuer As String
    Dim militaryUnit As String

    If rowsToExport.Count = 0 Then Exit Function

    Set firstRecord = mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowsToExport(1)))
    orderDate = RecordFieldText(firstRecord, "order_date")
    orderNumber = RecordFieldText(firstRecord, "order_number")
    orderIssuer = RecordFieldText(firstRecord, "order_issuer")
    militaryUnit = RecordFieldText(firstRecord, "military_unit")

    For Each rowNum In rowsToExport
        Set currentRecord = mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowNum))
        mismatch = CompareHeaderField(mismatch, orderDate, RecordFieldText(currentRecord, "order_date"), L("enrollment.word.conflict.order_date", "Дата приказа"))
        mismatch = CompareHeaderField(mismatch, orderNumber, RecordFieldText(currentRecord, "order_number"), L("enrollment.word.conflict.order_number", "Номер приказа"))
        mismatch = CompareHeaderField(mismatch, orderIssuer, RecordFieldText(currentRecord, "order_issuer"), L("enrollment.word.conflict.order_issuer", "Кем издан приказ"))
        mismatch = CompareHeaderField(mismatch, militaryUnit, RecordFieldText(currentRecord, "military_unit"), L("enrollment.word.conflict.unit", "Воинская часть в строке"))
    Next rowNum

    BuildGroupHeaderConflictText = mismatch
End Function

Private Function ValidateGroupHeader(ByVal rowsToExport As Collection) As Object
    Dim header As Object
    Dim firstRecord As Object
    Dim currentRecord As Object
    Dim rowNum As Variant
    Dim mismatch As String

    Set header = CreateObject("Scripting.Dictionary")
    header.CompareMode = vbTextCompare

    Set firstRecord = mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowsToExport(1)))
    header("partNumber") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.unit_number", "81510")
    header("city") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.city", mdlHelper.Ru(1043, 1088, 1086, 1079, 1085, 1099, 1081))
    header("signatoryName") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.signatory_name", mdlHelper.Ru(1045) & "." & mdlHelper.Ru(1050, 1086, 1088, 1086, 1087, 1077, 1094))
    header("signatoryRank") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.signatory_rank", mdlHelper.Ru(1084, 1072, 1081, 1086, 1088))
    header("signatoryPosition") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.signatory_position", mdlHelper.Ru(1042, 1056, 1048, 1054) & " " & mdlHelper.Ru(1050, 1054, 1052, 1040, 1053, 1044, 1048, 1056, 1040) & " " & mdlHelper.Ru(1042, 1054, 1049, 1057, 1050, 1054, 1049) & " " & mdlHelper.Ru(1063, 1040, 1057, 1058, 1048) & " 81510")
    header("headerText") = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.header_text", mdlHelper.Ru(1055, 1056, 1054, 1045, 1050, 1058) & " " & mdlHelper.Ru(1055, 1056, 1048, 1050, 1040, 1047, 1040) & "|" & mdlHelper.Ru(1050, 1054, 1052, 1040, 1053, 1044, 1048, 1056, 1040) & " " & mdlHelper.Ru(1042, 1054, 1049, 1057, 1050, 1054, 1049) & " " & mdlHelper.Ru(1063, 1040, 1057, 1058, 1048) & " {unit}|(" & mdlHelper.Ru(1087, 1086, 32, 1089, 1090, 1088, 1086, 1077, 1074, 1086, 1081, 32, 1095, 1072, 1089, 1090, 1080) & ")")
    header("orderDate") = RecordFieldText(firstRecord, "order_date")
    header("orderNumber") = RecordFieldText(firstRecord, "order_number")
    header("orderIssuer") = RecordFieldText(firstRecord, "order_issuer")
    header("militaryUnit") = RecordFieldText(firstRecord, "military_unit")
    header("orderDraftId") = RecordFieldText(firstRecord, "order_draft_id")

    mismatch = BuildGroupHeaderConflictText(rowsToExport)

    If mismatch <> "" Then
        Err.Raise vbObjectError + 1702, "ValidateGroupHeader", L("enrollment.word.error.group_conflict", "Невозможно собрать один приказ по OrderDraftId из-за конфликтов:") & vbCrLf & mismatch
    End If

    Set ValidateGroupHeader = header
End Function

Private Function CompareHeaderField(ByVal currentIssues As String, ByVal leftValue As String, ByVal rightValue As String, ByVal fieldLabel As String) As String
    If Trim$(leftValue) = Trim$(rightValue) Then
        CompareHeaderField = currentIssues
    Else
        If currentIssues <> "" Then currentIssues = currentIssues & vbCrLf
        CompareHeaderField = currentIssues & "- " & fieldLabel & ": '" & leftValue & "' <> '" & rightValue & "'"
    End If
End Function

Private Function BuildDocumentHeader(ByVal header As Object) As String
    Dim lines() As String
    Dim i As Long
    Dim resultText As String

    lines = Split(Replace$(SafeText(header("headerText")), "{unit}", SafeText(header("partNumber"))), "|")
    For i = LBound(lines) To UBound(lines)
        resultText = resultText & Trim$(lines(i)) & vbCrLf
    Next i

    resultText = resultText & BuildQuotePlaceholder() & " " & SafeText(header("orderDate")) & " " & L("enrollment.word.header.year_short", "г.") & " " & L("enrollment.word.header.no", "№") & " " & SafeText(header("orderNumber")) & vbCrLf
    resultText = resultText & L("enrollment.word.header.city_prefix", "г.") & " " & SafeText(header("city"))
    BuildDocumentHeader = resultText
End Function

Private Function BuildSection1ForRecord(ByVal record As Object, Optional ByVal itemNumber As Long = 1) As String
    Dim resultText As String
    Dim monthlyText As String
    Dim personalText As String
    Dim premiumText As String

    resultText = CStr(itemNumber) & ". " & BuildCoreEnrollmentSentence(record) & vbCrLf
    resultText = resultText & BuildSalaryParagraph(record) & vbCrLf

    If NormalizeYesNo(record("preferential_enabled")) = "YES" Then
        resultText = resultText & BuildPreferentialParagraph(record) & vbCrLf
    End If

    monthlyText = BuildStandardMonthlyText(record)
    If monthlyText <> "" Then resultText = resultText & monthlyText & vbCrLf

    personalText = BuildPersonalMonthlyText(record)
    If personalText <> "" Then resultText = resultText & personalText & vbCrLf

    resultText = resultText & BuildOneTimeText(record)

    premiumText = BuildPremiumText(record)
    If premiumText <> "" Then resultText = resultText & vbCrLf & premiumText

    resultText = resultText & vbCrLf & BuildRequisitesText(record)
    BuildSection1ForRecord = resultText
End Function

Private Function BuildCoreEnrollmentSentence(ByVal record As Object) As String
    Dim definition As Object
    Dim templateText As String
    Dim renderedText As String

    Set definition = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("core")
    templateText = SafeText(definition("text_template"))
    If templateText <> "" And StrComp(templateText, SafeText(definition("label")), vbTextCompare) <> 0 Then
        renderedText = RenderDefinitionTemplate(templateText, record, definition)
        If renderedText <> "" Then
            BuildCoreEnrollmentSentence = renderedText
            Exit Function
        End If
    End If

    BuildCoreEnrollmentSentence = BuildDefaultCoreEnrollmentSentence(record)
End Function

Private Function BuildDefaultCoreEnrollmentSentence(ByVal record As Object) As String
    Dim resultText As String
    Dim orderReference As String
    Dim acceptText As String

    resultText = SafeText(record("rank")) & " " & SafeText(record("fio")) & ", " & SafeText(record("personal_number"))
    If NormalizeYesNo(record("arrival_details_enabled")) = "YES" And SafeText(record("arrival_source")) <> "" Then
        resultText = resultText & ", " & L("enrollment.word.core.arrived_from", "прибывшего") & " " & SafeText(record("arrival_source"))
    End If

    orderReference = BuildOrderReference(record)
    If orderReference <> "" Then
        resultText = resultText & ", " & L("enrollment.word.core.assigned_by", "назначенного") & " " & orderReference
    End If

    If SafeText(record("position")) <> "" Then
        resultText = resultText & " " & L("enrollment.word.core.to_position", "на воинскую должность") & " " & SafeText(record("position"))
    End If
    If SafeText(record("military_unit")) <> "" Then
        resultText = resultText & " " & SafeText(record("military_unit"))
    End If
    If SafeText(record("vus")) <> "" Then
        resultText = resultText & ", " & L("enrollment.word.core.vus", "ВУС") & " - " & SafeText(record("vus"))
    End If

    resultText = resultText & ", " & L("enrollment.word.core.enroll_prefix", "с") & " " & SafeText(record("enroll_date")) & " " & L("enrollment.word.header.year_short", "г.") & " " & L("enrollment.word.core.enroll_text", "зачислить в списки личного состава") & " " & SafeText(record("military_unit"))

    If SafeText(record("accept_date")) <> "" Then
        acceptText = L("enrollment.word.core.accept_text", "с {date} полагать принявшим дела и должность")
        acceptText = Replace$(acceptText, "{date}", SafeText(record("accept_date")))
        If SafeText(record("position")) <> "" Then
            acceptText = acceptText & " " & SafeText(record("position"))
        End If
        resultText = resultText & ", " & acceptText
    End If

    If SafeText(record("duty_start_date")) <> "" Then
        resultText = resultText & ", " & Replace$(L("enrollment.word.core.duty_text", "и вступившим в исполнение обязанностей с {date}"), "{date}", SafeText(record("duty_start_date")))
    End If

    resultText = resultText & "."
    BuildDefaultCoreEnrollmentSentence = resultText
End Function

Private Function BuildStandardMonthlyText(ByVal record As Object) As String
    Dim definition As Object
    Dim resultText As String

    For Each definition In mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock("Section1MonthlyStandard")
        resultText = AppendParagraph(resultText, BuildConfiguredEnrollmentParagraph(record, definition))
    Next definition

    BuildStandardMonthlyText = resultText
End Function

Private Function BuildPersonalMonthlyText(ByVal record As Object) As String
    Dim definition As Object
    Dim resultText As String

    For Each definition In mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock("Section1MonthlyPersonal")
        If StrComp(SafeText(definition("code")), "extra_monthly", vbTextCompare) <> 0 Then
            resultText = AppendParagraph(resultText, BuildConfiguredEnrollmentParagraph(record, definition))
        End If
    Next definition

    resultText = AppendParagraph(resultText, BuildExtraMonthlyText(record))

    BuildPersonalMonthlyText = resultText
End Function

Private Function BuildOneTimeText(ByVal record As Object) As String
    Dim definition As Object
    Dim resultText As String

    If HasConfiguredParagraph(record, "Section1OneTime") Or BuildExtraOneTimeText(record) <> "" Then
        resultText = L("enrollment.word.onetime.header", "Выплатить:")
    End If
    For Each definition In mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock("Section1OneTime")
        If StrComp(SafeText(definition("code")), "extra_onetime", vbTextCompare) <> 0 Then
            resultText = AppendParagraph(resultText, BuildConfiguredEnrollmentParagraph(record, definition))
        End If
    Next definition
    resultText = AppendParagraph(resultText, BuildExtraOneTimeText(record))

    BuildOneTimeText = resultText
End Function

Private Function BuildConfiguredEnrollmentParagraph(ByVal record As Object, ByVal definition As Object) As String
    Dim templateText As String

    templateText = SafeText(definition("text_template"))
    If templateText <> "" Then
        If IsConfiguredPaymentActive(record, definition) Then
            BuildConfiguredEnrollmentParagraph = RenderDefinitionTemplate(templateText, record, definition)
            If StrComp(SafeText(definition("code")), "per_diem", vbTextCompare) = 0 Then
                BuildConfiguredEnrollmentParagraph = AppendPerDiemDaysText(BuildConfiguredEnrollmentParagraph, record, templateText)
            End If
            Exit Function
        End If
    End If

    Select Case SafeText(definition("code"))
        Case "std_duty"
            If NormalizeYesNo(record("std_duty_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = Replace$(L("enrollment.word.monthly.std_duty", "Установить ежемесячную надбавку к денежному довольствию в размере {percent}% оклада по воинской должности."), "{percent}", ValueOrPlaceholder(record("std_duty_percent"), DEFAULT_POSITION_ALLOWANCE_PERCENT))
            End If
        Case "std_special"
            If NormalizeYesNo(record("std_special_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = Replace$(L("enrollment.word.monthly.std_special", "Установить ежемесячную надбавку за особые условия военной службы в размере {percent}% оклада по воинской должности."), "{percent}", ValueOrPlaceholder(record("std_special_percent"), DEFAULT_SPECIAL_CONDITIONS_PERCENT))
            End If
        Case "std_tariff"
            If NormalizeYesNo(record("std_tariff_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = Replace$(L("enrollment.word.monthly.std_tariff", "Установить ежемесячную надбавку по должностям 1-4 тарифных разрядов в размере {percent}% оклада по воинской должности."), "{percent}", ValueOrPlaceholder(record("std_tariff_percent"), DEFAULT_TARIFF_PERCENT))
            End If
        Case "std_contract430"
            If NormalizeYesNo(record("std_contract430_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = Replace$(L("enrollment.word.monthly.std_contract430", "Установить ежемесячную надбавку за особые достижения в службе по 430 ДСП в размере {percent}%."), "{percent}", ValueOrPlaceholder(record("std_contract430_percent"), DEFAULT_CONTRACT430_PERCENT))
            End If
        Case "class"
            If NormalizeYesNo(record("class_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = L("enrollment.word.personal.class", "Установить ежемесячную надбавку за классную квалификацию") & " (" & SafeText(record("class_param")) & ") " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(record("class_percent"), "0") & "%."
            End If
        Case "fizo"
            If NormalizeYesNo(record("fizo_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = L("enrollment.word.personal.fizo", "Установить ежемесячную надбавку за уровень физической подготовленности") & " (" & SafeText(record("fizo_param")) & ") " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(record("fizo_percent"), "0") & "%."
            End If
        Case "secrecy"
            If NormalizeYesNo(record("secrecy_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = L("enrollment.word.personal.secrecy", "Установить ежемесячную надбавку за работу со сведениями, составляющими государственную тайну") & " (" & SafeText(record("secrecy_param")) & ") " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(record("secrecy_percent"), "0") & "%."
            End If
        Case "achievement"
            If NormalizeYesNo(record("achievement_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = L("enrollment.word.personal.achievement", "Установить ежемесячную надбавку за особые достижения в службе / медаль") & " (" & SafeText(record("achievement_param")) & ") " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(record("achievement_amount"), "0") & IIf(IsNumeric(SafeText(record("achievement_amount"))), "%", "") & "."
            End If
        Case "lift"
            If NormalizeYesNo(record("lift_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = "    " & L("enrollment.word.onetime.lift", "подъёмное пособие в размере") & " " & ValueOrPlaceholder(record("lift_amount"), DefaultLiftAmountText()) & "."
            End If
        Case "per_diem"
            If NormalizeYesNo(record("per_diem_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = "    " & L("enrollment.word.onetime.per_diem", "суточные в размере") & " " & ValueOrPlaceholder(record("per_diem_amount"), DefaultPerDiemAmountText())
                BuildConfiguredEnrollmentParagraph = AppendPerDiemDaysText(BuildConfiguredEnrollmentParagraph & ".", record, "")
            End If
        Case "premium"
            If NormalizeYesNo(record("premium_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = L("enrollment.word.section1.premium_text", "Достоин выплаты ежемесячной премии в размере") & " " & ValueOrPlaceholder(record("premium_percent"), DEFAULT_PREMIUM_PERCENT) & "% " & L("enrollment.word.section1.premium_from", "с") & " " & SafeText(record("premium_start")) & " " & L("enrollment.word.section1.premium_to", "по") & " " & SafeText(record("premium_end")) & "."
            End If
        Case "edv"
            If NormalizeYesNo(record("edv_enabled")) = "YES" Then
                BuildConfiguredEnrollmentParagraph = SafeText(record("rank")) & " " & SafeText(record("fio")) & ", " & SafeText(record("personal_number")) & _
                    ", " & SafeText(record("position")) & ", " & L("enrollment.word.section2.edv_text", "выплатить единовременную денежную выплату в размере") & " " & _
                    ValueOrPlaceholder(record("edv_amount"), "400000") & " " & L("enrollment.word.section2.currency", "рублей.")
            End If
    End Select
End Function

Private Function HasConfiguredParagraph(ByVal record As Object, ByVal wordBlockTarget As String) As Boolean
    Dim definition As Object

    For Each definition In mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock(wordBlockTarget)
        If Trim$(BuildConfiguredEnrollmentParagraph(record, definition)) <> "" Then
            HasConfiguredParagraph = True
            Exit Function
        End If
    Next definition
End Function

Private Function BuildExtraMonthlyText(ByVal record As Object) As String
    Dim i As Long
    Dim resultText As String
    Dim paymentName As String
    Dim paymentParam As String
    Dim paymentAmount As String
    Dim startDate As String
    Dim paymentBasis As String

    For i = 1 To 4
        If NormalizeYesNo(record("extra_monthly" & CStr(i) & "_enabled")) = "YES" Then
            paymentName = SafeText(record("extra_monthly" & CStr(i) & "_name"))
            paymentParam = SafeText(record("extra_monthly" & CStr(i) & "_param"))
            paymentAmount = SafeText(record("extra_monthly" & CStr(i) & "_amount"))
            startDate = SafeText(record("extra_monthly" & CStr(i) & "_start"))
            paymentBasis = SafeText(record("extra_monthly" & CStr(i) & "_basis"))
            If paymentBasis <> "" Then paymentParam = Trim$(paymentParam & "; основание: " & paymentBasis)

            If paymentName <> "" Then
                BuildExtraMonthlyText = AppendParagraph(BuildExtraMonthlyText, L("enrollment.word.extra.monthly_prefix", "Установить ежемесячную выплату") & " """ & paymentName & """" & BuildExtraParamSuffix(paymentParam) & " " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(paymentAmount, "0") & BuildExtraStartSuffix(startDate) & ".")
            End If
        End If
    Next i

    resultText = BuildExtraMonthlyText
    BuildExtraMonthlyText = resultText
End Function

Private Function BuildExtraOneTimeText(ByVal record As Object) As String
    Dim i As Long
    Dim paymentName As String
    Dim paymentAmount As String
    Dim paymentDate As String
    Dim paymentBasis As String

    For i = 1 To 3
        If NormalizeYesNo(record("extra_one_time" & CStr(i) & "_enabled")) = "YES" Then
            paymentName = SafeText(record("extra_one_time" & CStr(i) & "_name"))
            paymentAmount = SafeText(record("extra_one_time" & CStr(i) & "_amount"))
            paymentDate = SafeText(record("extra_one_time" & CStr(i) & "_date"))
            paymentBasis = SafeText(record("extra_one_time" & CStr(i) & "_basis"))
            If paymentBasis <> "" Then paymentName = paymentName & " (" & paymentBasis & ")"

            If paymentName <> "" Then
                BuildExtraOneTimeText = AppendParagraph(BuildExtraOneTimeText, "    " & L("enrollment.word.extra.onetime_prefix", "дополнительную разовую выплату") & " """ & paymentName & """" & " " & L("enrollment.word.personal.amount_prefix", "в размере") & " " & ValueOrPlaceholder(paymentAmount, "0") & BuildExtraStartSuffix(paymentDate) & ".")
            End If
        End If
    Next i
End Function

Private Function BuildExtraParamSuffix(ByVal paramValue As String) As String
    If Trim$(paramValue) = "" Then Exit Function
    BuildExtraParamSuffix = " (" & paramValue & ")"
End Function

Private Function BuildExtraStartSuffix(ByVal dateText As String) As String
    If Trim$(dateText) = "" Then Exit Function
    BuildExtraStartSuffix = " " & L("enrollment.word.section1.premium_from", "с") & " " & dateText
End Function

Private Function BuildRequisitesText(ByVal record As Object) As String
    Dim definition As Object
    Dim templateText As String
    Dim renderedText As String

    Set definition = mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("requisites")
    templateText = SafeText(definition("text_template"))
    If templateText <> "" And StrComp(templateText, SafeText(definition("label")), vbTextCompare) <> 0 Then
        renderedText = RenderDefinitionTemplate(templateText, record, definition)
        If renderedText <> "" Then
            BuildRequisitesText = renderedText
            Exit Function
        End If
    End If

    BuildRequisitesText = BuildDefaultRequisitesText(record)
End Function

Private Function BuildDefaultRequisitesText(ByVal record As Object) As String
    Dim resultText As String
    Dim personalText As String

    If NormalizeYesNo(record("personal_details_enabled")) = "YES" And (SafeText(record("birth_date")) <> "" Or SafeText(record("birth_place")) <> "" Or SafeText(record("citizenship")) <> "") Then
        If SafeText(record("birth_date")) <> "" Then
            personalText = L("enrollment.word.requisites.birth_date", "Дата рождения:") & " " & SafeText(record("birth_date"))
        End If
        If SafeText(record("birth_place")) <> "" Then
            If personalText <> "" Then personalText = personalText & ", "
            personalText = personalText & L("enrollment.word.requisites.birth_place", "место рождения:") & " " & SafeText(record("birth_place"))
        End If
        If SafeText(record("citizenship")) <> "" Then
            If personalText <> "" Then personalText = personalText & ", "
            personalText = personalText & L("enrollment.word.requisites.citizenship", "гражданство:") & " " & SafeText(record("citizenship"))
        End If
        resultText = personalText & "." & vbCrLf
    End If

    If NormalizeYesNo(record("personal_details_enabled")) = "YES" Then
        resultText = resultText & L("enrollment.word.requisites.passport_series", "Паспорт серия") & " " & SafeText(record("passport_series")) & " " & L("enrollment.word.header.no", "№") & " " & SafeText(record("passport_number"))
        resultText = resultText & ", " & L("enrollment.word.requisites.issued_by", "выдан") & " " & SafeText(record("passport_issuer"))
        If SafeText(record("passport_issue_date")) <> "" Then resultText = resultText & ", " & L("enrollment.word.requisites.issue_date", "дата выдачи:") & " " & SafeText(record("passport_issue_date"))
        If SafeText(record("passport_code")) <> "" Then resultText = resultText & ", " & L("enrollment.word.requisites.passport_code", "код подразделения") & " " & SafeText(record("passport_code"))
        resultText = resultText & "." & vbCrLf
        resultText = resultText & L("enrollment.word.requisites.inn", "ИНН") & " - " & SafeText(record("inn")) & ", " & L("enrollment.word.requisites.snils", "СНИЛС") & " - " & SafeText(record("snils")) & "." & vbCrLf
    End If
    If NormalizeYesNo(record("bank_details_enabled")) = "YES" Then
        resultText = resultText & L("enrollment.word.requisites.bank", "Лицевой счёт / банк:") & " " & SafeText(record("bank_name")) & " - " & SafeText(record("bank_account"))
        If SafeText(record("bank_bik")) <> "" Then resultText = resultText & "; БИК " & SafeText(record("bank_bik"))
        resultText = resultText & "." & vbCrLf
        If SafeText(record("requisites_note")) <> "" Then
            resultText = resultText & SafeText(record("requisites_note")) & "." & vbCrLf
        End If
    End If
    resultText = resultText & L("enrollment.word.requisites.basis", "Основание:") & " " & SafeText(record("basis_section1")) & "."

    BuildDefaultRequisitesText = resultText
End Function

Private Function BuildSection2ForRecord(ByVal record As Object, Optional ByVal itemNumber As Long = 1) As String
    If NormalizeYesNo(record("edv_enabled")) <> "YES" Then Exit Function

    BuildSection2ForRecord = CStr(itemNumber) & ". " & BuildConfiguredEnrollmentParagraph(record, mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinition("edv")) & vbCrLf & _
        L("enrollment.word.requisites.basis", "Основание:") & " " & SafeText(record("basis_section2")) & "."
End Function

Private Function BuildPremiumText(ByVal record As Object) As String
    Dim definition As Object

    For Each definition In mdlEnrollmentWorkflow.GetEnrollmentPaymentDefinitionsByBlock("Section1Premium")
        BuildPremiumText = AppendParagraph(BuildPremiumText, BuildConfiguredEnrollmentParagraph(record, definition))
    Next definition
End Function

Private Function RenderDefinitionTemplate(ByVal templateText As String, ByVal record As Object, ByVal definition As Object) As String
    Dim resultText As String

    resultText = templateText
    resultText = Replace$(resultText, "{label}", SafeText(definition("label")))
    resultText = Replace$(resultText, "{rank}", SafeText(record("rank")))
    resultText = Replace$(resultText, "{fio}", SafeText(record("fio")))
    resultText = Replace$(resultText, "{personal_number}", SafeText(record("personal_number")))
    resultText = Replace$(resultText, "{position}", SafeText(record("position")))
    resultText = Replace$(resultText, "{military_unit}", SafeText(record("military_unit")))
    resultText = Replace$(resultText, "{vus}", SafeText(record("vus")))
    resultText = Replace$(resultText, "{basis}", ResolveDefinitionFieldValue(record, definition, "basis"))
    resultText = Replace$(resultText, "{param}", ResolveDefinitionFieldValue(record, definition, "param"))
    resultText = Replace$(resultText, "{date}", ResolveDefinitionFieldValue(record, definition, "date"))
    resultText = Replace$(resultText, "{start}", ResolveDefinitionFieldValue(record, definition, "start"))
    resultText = Replace$(resultText, "{end}", ResolveDefinitionFieldValue(record, definition, "premium_end"))
    resultText = Replace$(resultText, "{percent}", ResolveDefinitionPercentValue(record, definition))
    resultText = Replace$(resultText, "{amount}", ResolveDefinitionAmountValue(record, definition))
    resultText = Replace$(resultText, "{amount_with_unit}", ResolveDefinitionAmountWithUnit(record, definition))
    resultText = Replace$(resultText, "{days}", ResolveDefinitionFieldValue(record, definition, "days"))
    resultText = ReplaceRecordFieldTokens(resultText, record, definition)
    resultText = Replace$(resultText, "{core_text}", BuildDefaultCoreEnrollmentSentence(record))
    resultText = Replace$(resultText, "{requisites_text}", BuildDefaultRequisitesText(record))

    RenderDefinitionTemplate = CleanupTemplateTokens(resultText)
End Function

Private Function ReplaceRecordFieldTokens(ByVal templateText As String, ByVal record As Object, ByVal definition As Object) As String
    Dim token As Variant

    For Each token In Array("fio", "rank", "position", "section", "vus", "order_date", "order_number", _
                            "order_issuer", "arrival_source", "contract_basis", "military_unit", _
                            "personal_number", "table_number", "service_category", "contract_kind", _
                            "tariff_rank", "position_salary", "rank_salary", "passport_series", _
                            "passport_number", "passport_issuer", "passport_issue_date", "passport_code", _
                            "inn", "snils", "bank_account", "bank_name", "bank_bik", "basis_section1", "basis_section2", _
                            "enroll_date", "accept_date", "duty_start_date", "standard_start_date", _
                            "preferential_start_date", "requisites_note", "birth_date", "birth_place", _
                            "citizenship")
        templateText = Replace$(templateText, "{" & CStr(token) & "}", ResolveDefinitionFieldValue(record, definition, CStr(token)))
    Next token

    ReplaceRecordFieldTokens = templateText
End Function

Private Function ResolveDefinitionFieldValue(ByVal record As Object, ByVal definition As Object, ByVal token As String) As String
    Dim fieldKey As String

    fieldKey = ResolveDefinitionFieldKey(definition, token)
    If fieldKey = "" Then Exit Function
    If Not record.Exists(fieldKey) Then Exit Function
    ResolveDefinitionFieldValue = SafeText(record(fieldKey))
End Function

Private Function ResolveDefinitionFieldKey(ByVal definition As Object, ByVal token As String) As String
    Dim bindingKey As String

    bindingKey = SafeText(definition("journal_binding"))
    Select Case LCase$(token)
        Case "fio", "rank", "position", "section", "vus", "order_date", "order_number", _
             "order_issuer", "arrival_source", "contract_basis", "military_unit", _
             "personal_number", "table_number", "service_category", "contract_kind", _
             "tariff_rank", "position_salary", "rank_salary", "passport_series", _
             "passport_number", "passport_issuer", "passport_issue_date", "passport_code", _
             "inn", "snils", "bank_account", "bank_name", "bank_bik", "basis_section1", "basis_section2", _
             "enroll_date", "accept_date", "duty_start_date", "standard_start_date", _
             "preferential_start_date", "requisites_note", "birth_date", "birth_place", _
             "citizenship"
            ResolveDefinitionFieldKey = LCase$(token)
        Case "days"
            ResolveDefinitionFieldKey = "per_diem_days"
        Case "unit"
            ResolveDefinitionFieldKey = "military_unit"
        Case "issuer"
            ResolveDefinitionFieldKey = "order_issuer"
        Case "basis"
            ResolveDefinitionFieldKey = bindingKey & "_basis"
        Case "param"
            ResolveDefinitionFieldKey = bindingKey & "_param"
        Case "amount"
            ResolveDefinitionFieldKey = bindingKey & "_amount"
        Case "date"
            ResolveDefinitionFieldKey = bindingKey & "_date"
        Case "start"
            If SafeText(definition("code")) = "premium" Then
                ResolveDefinitionFieldKey = "premium_start"
            Else
                ResolveDefinitionFieldKey = bindingKey & "_start"
            End If
        Case "premium_end", "end"
            ResolveDefinitionFieldKey = "premium_end"
    End Select
End Function

Private Function ResolveDefinitionPercentValue(ByVal record As Object, ByVal definition As Object) As String
    Select Case LCase$(SafeText(definition("code")))
        Case "std_duty"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("std_duty_percent"), DEFAULT_POSITION_ALLOWANCE_PERCENT)
        Case "std_special"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("std_special_percent"), DEFAULT_SPECIAL_CONDITIONS_PERCENT)
        Case "std_tariff"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("std_tariff_percent"), DEFAULT_TARIFF_PERCENT)
        Case "std_contract430"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("std_contract430_percent"), DEFAULT_CONTRACT430_PERCENT)
        Case "class"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("class_percent"), "0")
        Case "fizo"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("fizo_percent"), "0")
        Case "secrecy"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("secrecy_percent"), "0")
        Case "premium"
            ResolveDefinitionPercentValue = ValueOrPlaceholder(record("premium_percent"), DEFAULT_PREMIUM_PERCENT)
    End Select
End Function

Private Function ResolveDefinitionAmountValue(ByVal record As Object, ByVal definition As Object) As String
    Select Case LCase$(SafeText(definition("code")))
        Case "achievement"
            ResolveDefinitionAmountValue = ValueOrPlaceholder(record("achievement_amount"), "0")
        Case "lift"
            ResolveDefinitionAmountValue = ValueOrPlaceholder(record("lift_amount"), DefaultLiftAmountText())
        Case "per_diem"
            ResolveDefinitionAmountValue = ValueOrPlaceholder(record("per_diem_amount"), DefaultPerDiemAmountText())
        Case "edv"
            ResolveDefinitionAmountValue = ValueOrPlaceholder(record("edv_amount"), "400000")
        Case Else
            ResolveDefinitionAmountValue = ResolveDefinitionPercentValue(record, definition)
    End Select
End Function

Private Function ResolveDefinitionAmountWithUnit(ByVal record As Object, ByVal definition As Object) As String
    Dim amountText As String

    amountText = ResolveDefinitionAmountValue(record, definition)
    Select Case LCase$(SafeText(definition("code")))
        Case "lift", "per_diem", "edv"
            ResolveDefinitionAmountWithUnit = amountText
        Case Else
            If amountText <> "" And Right$(amountText, 1) <> "%" Then
                ResolveDefinitionAmountWithUnit = amountText & "%"
            Else
                ResolveDefinitionAmountWithUnit = amountText
            End If
    End Select
End Function

Private Function CleanupTemplateTokens(ByVal textValue As String) As String
    textValue = Replace$(textValue, " ()", "")
    textValue = Replace$(textValue, "( )", "")
    textValue = Replace$(textValue, " ,", ",")
    textValue = Replace$(textValue, " .", ".")
    textValue = Replace$(textValue, "  .", ".")
    textValue = Replace$(textValue, "  ,", ",")
    CleanupTemplateTokens = textValue
End Function

Private Function BuildSection2Intro() As String
    BuildSection2Intro = L("enrollment.word.section2.intro", "Нижепоименованным военнослужащим выплатить единовременную денежную выплату:")
End Function

Private Function BuildSignatureBlock(ByVal header As Object) As String
    BuildSignatureBlock = SafeText(header("signatoryPosition")) & vbCrLf & _
                          SafeText(header("signatoryRank")) & " " & SafeText(header("signatoryName"))
End Function

Private Function HasSection2Rows(ByVal rowsToExport As Collection) As Boolean
    Dim rowNum As Variant
    Dim record As Object

    For Each rowNum In rowsToExport
        Set record = mdlEnrollmentWorkflow.GetResolvedEnrollmentRecordByRow(CLng(rowNum))
        If NormalizeYesNo(record("edv_enabled")) = "YES" Then
            HasSection2Rows = True
            Exit Function
        End If
    Next rowNum
End Function

Private Function BuildOutputFilePath(ByVal orderDraftId As String) As String
    Dim templateText As String
    Dim fileName As String

    templateText = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.filename_template", mdlHelper.Ru(1055, 1088, 1080, 1082, 1072, 1079) & "_" & mdlHelper.Ru(1086) & "_" & mdlHelper.Ru(1079, 1072, 1095, 1080, 1089, 1083, 1077, 1085, 1080, 1080) & "_{orderDraftId}_{date}")
    fileName = Replace$(templateText, "{orderDraftId}", IIf(Trim$(orderDraftId) = "", "single", orderDraftId))
    fileName = Replace$(fileName, "{date}", Format$(Date, "dd.MM.yyyy"))
    fileName = SanitizeFileName(fileName) & ".docx"
    BuildOutputFilePath = ThisWorkbook.Path & "\" & fileName
End Function

Private Function BuildSalaryParagraph(ByVal record As Object) As String
    Dim resultText As String

    resultText = L("enrollment.word.section1.salary_prefix", "С") & " " & SafeText(record("enroll_date")) & " " & L("enrollment.word.header.year_short", "г.") & " " & _
                 L("enrollment.word.section1.salary_text", "установить оклад по воинской должности в размере") & " " & _
                 ValueOrPlaceholder(record("position_salary"), L("enrollment.word.placeholder.position_salary", "[указать оклад]"))
    If SafeText(record("tariff_rank")) <> "" Then
        resultText = resultText & " " & L("enrollment.word.section1.salary_suffix", "руб. в месяц") & " (" & SafeText(record("tariff_rank")) & " " & L("enrollment.word.section1.tariff_suffix", "тарифный разряд") & ")."
    Else
        resultText = resultText & " " & L("enrollment.word.section1.salary_suffix", "руб. в месяц") & "."
    End If

    If SafeText(record("rank_salary")) <> "" Then
        resultText = resultText & vbCrLf & L("enrollment.word.section1.salary_prefix", "С") & " " & SafeText(record("enroll_date")) & " " & L("enrollment.word.header.year_short", "г.") & " " & _
                     L("enrollment.word.section1.rank_salary_text", "установить оклад по воинскому званию в размере") & " " & _
                     ValueOrPlaceholder(record("rank_salary"), L("enrollment.word.placeholder.rank_salary", "[указать оклад по званию]")) & " " & _
                     L("enrollment.word.section1.salary_suffix", "руб. в месяц") & "."
    End If

    BuildSalaryParagraph = resultText
End Function

Private Function AppendPerDiemDaysText(ByVal paragraphText As String, ByVal record As Object, ByVal templateText As String) As String
    Dim resultText As String
    Dim daysText As String

    resultText = Trim$(paragraphText)
    daysText = SafeText(record("per_diem_days"))
    If daysText = "" Then
        AppendPerDiemDaysText = resultText
        Exit Function
    End If
    If InStr(1, templateText, "{days}", vbTextCompare) > 0 Then
        AppendPerDiemDaysText = resultText
        Exit Function
    End If

    If Right$(resultText, 1) = "." Then
        resultText = Left$(resultText, Len(resultText) - 1)
    End If
    AppendPerDiemDaysText = resultText & " " & L("enrollment.word.onetime.per_diem_days_prefix", "за") & " " & daysText & " " & L("enrollment.word.onetime.per_diem_days_suffix", "сут.") & "."
End Function

Private Function BuildPreferentialParagraph(ByVal record As Object) As String
    BuildPreferentialParagraph = L("enrollment.word.section1.preferential_prefix", "С") & " " & SafeText(record("preferential_start_date")) & " " & _
                                 L("enrollment.word.header.year_short", "г.") & " " & _
                                 L("enrollment.word.section1.preferential_text", "исчислять выслугу лет на пенсию на льготных условиях, из расчёта один месяц военной службы за") & " " & _
                                 ValueOrPlaceholder(record("preferential_coeff"), "1.5") & " " & _
                                 L("enrollment.word.section1.preferential_suffix", "месяца.")
End Function

Private Function BuildOrderReference(ByVal record As Object) As String
    Dim resultText As String
    Dim issuerText As String

    issuerText = SafeText(record("order_issuer"))

    If issuerText <> "" Then
        If LCase$(Left$(issuerText, 8)) = LCase$(L("enrollment.word.core.order_word", "приказом")) Then
            resultText = issuerText
        Else
            resultText = L("enrollment.word.core.order_word", "приказом") & " " & issuerText
        End If
    ElseIf SafeText(record("order_date")) <> "" Or SafeText(record("order_number")) <> "" Then
        resultText = L("enrollment.word.core.order_word", "приказом")
    End If

    If SafeText(record("order_date")) <> "" Then
        resultText = Trim$(resultText & " " & L("enrollment.word.core.from", "от") & " " & SafeText(record("order_date")))
    End If
    If SafeText(record("order_number")) <> "" Then
        resultText = Trim$(resultText & " " & L("enrollment.word.header.no", "№") & " " & SafeText(record("order_number")))
    End If

    BuildOrderReference = Trim$(resultText)
End Function

Private Function SanitizeFileName(ByVal fileName As String) As String
    Dim invalidChars As Variant
    Dim i As Long

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        fileName = Replace$(fileName, CStr(invalidChars(i)), "_")
    Next i
    SanitizeFileName = fileName
End Function

Private Function ValueOrPlaceholder(ByVal valueText As Variant, ByVal fallbackValue As String) As String
    If Trim$(CStr(valueText)) = "" Then
        ValueOrPlaceholder = fallbackValue
    Else
        ValueOrPlaceholder = Trim$(CStr(valueText))
    End If
End Function

Private Function AppendParagraph(ByVal sourceText As String, ByVal paragraphText As String) As String
    If Trim$(paragraphText) = "" Then
        AppendParagraph = sourceText
    ElseIf Trim$(sourceText) = "" Then
        AppendParagraph = paragraphText
    Else
        AppendParagraph = sourceText & vbCrLf & paragraphText
    End If
End Function

Private Function NormalizeYesNo(ByVal valueText As Variant) As String
    Dim normalized As String

    normalized = UCase$(Trim$(CStr(valueText)))
    Select Case normalized
        Case "1", "TRUE", "YES", UCase$(mdlHelper.Ru(1044, 1072)), "Y"
            NormalizeYesNo = "YES"
        Case Else
            NormalizeYesNo = "NO"
    End Select
End Function

Private Function SafeText(ByVal valueText As Variant) As String
    If IsError(valueText) Then Exit Function
    If IsNull(valueText) Then Exit Function
    SafeText = Trim$(CStr(valueText))
End Function

Private Function RecordFieldText(ByVal record As Object, ByVal fieldKey As String) As String
    On Error Resume Next
    If record Is Nothing Then Exit Function
    If record.Exists(fieldKey) Then
        RecordFieldText = SafeText(record(fieldKey))
    End If
    On Error GoTo 0
End Function

Private Function EnrollmentTemplateFileName() As String
    EnrollmentTemplateFileName = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.template_file", EnrollmentDefaultTemplateFileName())
End Function

Private Function EnrollmentTemplateBodyMarker() As String
    EnrollmentTemplateBodyMarker = mdlEnrollmentWorkflow.GetEnrollmentSetting("enrollment.template_body_marker", ENROLLMENT_BODY_MARKER_DEFAULT)
End Function

Private Function EnrollmentDefaultTemplateFileName() As String
    EnrollmentDefaultTemplateFileName = mdlHelper.Ru(1064, 1072, 1073, 1083, 1086, 1085, 95, 1047, 1072, 1095, 1080, 1089, 1083, 1077, 1085, 1080, 1077, 46, 100, 111, 99, 120)
End Function

Private Function L(ByVal key As String, ByVal fallback As String) As String
    Dim resolvedText As String
    Dim safeFallback As String

    resolvedText = t(key, fallback)
    If StrComp(resolvedText, key, vbTextCompare) = 0 Then
        safeFallback = EnrollmentWordSafeFallback(key)
        If safeFallback <> "" Then
            resolvedText = safeFallback
        End If
    End If

    L = resolvedText
End Function

Private Function EnrollmentWordSafeFallback(ByVal key As String) As String
    Select Case LCase$(Trim$(key))
        Case "enrollment.word.monthly.std_duty"
            EnrollmentWordSafeFallback = BuildUnicodeText(1059, 1089, 1090, 1072, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1077, 1078, 1077, 1084, 1077, 1089, 1103, 1095, 1085, 1091, 1102, 32, 1085, 1072, 1076, 1073, 1072, 1074, 1082, 1091, 32, 1082, 32, 1076, 1077, 1085, 1077, 1078, 1085, 1086, 1084, 1091, 32, 1076, 1086, 1074, 1086, 1083, 1100, 1089, 1090, 1074, 1080, 1102, 32, 1074, 32, 1088, 1072, 1079, 1084, 1077, 1088, 1077, 32, 123, 112, 101, 114, 99, 101, 110, 116, 125, 37, 32, 1086, 1082, 1083, 1072, 1076, 1072, 32, 1087, 1086, 32, 1074, 1086, 1080, 1085, 1089, 1082, 1086, 1081, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1080, 46)
        Case "enrollment.word.monthly.std_special"
            EnrollmentWordSafeFallback = BuildUnicodeText(1059, 1089, 1090, 1072, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1077, 1078, 1077, 1084, 1077, 1089, 1103, 1095, 1085, 1091, 1102, 32, 1085, 1072, 1076, 1073, 1072, 1074, 1082, 1091, 32, 1079, 1072, 32, 1086, 1089, 1086, 1073, 1099, 1077, 32, 1091, 1089, 1083, 1086, 1074, 1080, 1103, 32, 1074, 1086, 1077, 1085, 1085, 1086, 1081, 32, 1089, 1083, 1091, 1078, 1073, 1099, 32, 1074, 32, 1088, 1072, 1079, 1084, 1077, 1088, 1077, 32, 123, 112, 101, 114, 99, 101, 110, 116, 125, 37, 32, 1086, 1082, 1083, 1072, 1076, 1072, 32, 1087, 1086, 32, 1074, 1086, 1080, 1085, 1089, 1082, 1086, 1081, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1080, 46)
        Case "enrollment.word.monthly.std_tariff"
            EnrollmentWordSafeFallback = BuildUnicodeText(1059, 1089, 1090, 1072, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1077, 1078, 1077, 1084, 1077, 1089, 1103, 1095, 1085, 1091, 1102, 32, 1085, 1072, 1076, 1073, 1072, 1074, 1082, 1091, 32, 1087, 1086, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1103, 1084, 32, 49, 45, 52, 32, 1090, 1072, 1088, 1080, 1092, 1085, 1099, 1093, 32, 1088, 1072, 1079, 1088, 1103, 1076, 1086, 1074, 32, 1074, 32, 1088, 1072, 1079, 1084, 1077, 1088, 1077, 32, 123, 112, 101, 114, 99, 101, 110, 116, 125, 37, 32, 1086, 1082, 1083, 1072, 1076, 1072, 32, 1087, 1086, 32, 1074, 1086, 1080, 1085, 1089, 1082, 1086, 1081, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1080, 46)
        Case "enrollment.word.monthly.std_contract430"
            EnrollmentWordSafeFallback = BuildUnicodeText(1059, 1089, 1090, 1072, 1085, 1086, 1074, 1080, 1090, 1100, 32, 1077, 1078, 1077, 1084, 1077, 1089, 1103, 1095, 1085, 1091, 1102, 32, 1085, 1072, 1076, 1073, 1072, 1074, 1082, 1091, 32, 1079, 1072, 32, 1086, 1089, 1086, 1073, 1099, 1077, 32, 1076, 1086, 1089, 1090, 1080, 1078, 1077, 1085, 1080, 1103, 32, 1074, 32, 1089, 1083, 1091, 1078, 1073, 1077, 32, 1087, 1086, 32, 52, 51, 48, 32, 1044, 1057, 1055, 32, 1074, 32, 1088, 1072, 1079, 1084, 1077, 1088, 1077, 32, 123, 112, 101, 114, 99, 101, 110, 116, 125, 37, 46)
    End Select
End Function

Private Function BuildUnicodeText(ParamArray codePoints() As Variant) As String
    Dim i As Long

    For i = LBound(codePoints) To UBound(codePoints)
        BuildUnicodeText = BuildUnicodeText & ChrW$(CLng(codePoints(i)))
    Next i
End Function

Private Function BuildQuotePlaceholder() As String
    BuildQuotePlaceholder = ChrW$(171) & "___" & ChrW$(187)
End Function

Private Function BuildSectionCaption(ByVal sectionNumber As Long) As String
    BuildSectionCaption = ChrW$(167) & " " & CStr(sectionNumber)
End Function

Private Function DefaultLiftAmountText() As String
    DefaultLiftAmountText = L("enrollment.word.onetime.default_lift_amount", "1 ОДС")
End Function

Private Function DefaultPerDiemAmountText() As String
    DefaultPerDiemAmountText = L("enrollment.word.onetime.default_per_diem_amount", "1 сутки")
End Function
