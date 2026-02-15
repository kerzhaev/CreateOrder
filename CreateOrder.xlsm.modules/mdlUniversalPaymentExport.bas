Attribute VB_Name = "mdlUniversalPaymentExport"
' ===============================================================================
' Module mdlUniversalPaymentExport
' Version: 1.0.0
' Date: 14.02.2026
' Description: Universal export of allowances without periods to Word
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for mass import of employees by numbers
' =============================================
Public Sub ImportEmployeesByNumbers()
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim selectedRange As Range
    Dim reportText As String
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Массовое добавление сотрудников..."
    
    ' Check active sheet
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox "Активен не тот лист. Перейдите на лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, "Ошибка"
        GoTo CleanUp
    End If
    
    Set wsPayments = ActiveSheet
    
    ' Get selected range
    Set selectedRange = Selection
    If selectedRange Is Nothing Then
        MsgBox "Выделите ячейки с номерами в колонке D (личный номер).", vbExclamation, "Ошибка"
        GoTo CleanUp
    End If
    
    ' Check if range is in column D
    If selectedRange.Column <> mdlPaymentValidation.COL_LICHNIY_NOMER Then
        MsgBox "Выделенный диапазон должен находиться в колонке D (личный номер).", vbExclamation, "Ошибка"
        GoTo CleanUp
    End If
    
    ' Process range
    reportText = ProcessSelectedRangeForImport(wsPayments, selectedRange)
    
    ' Show report
    MsgBox reportText, vbInformation, "Массовое добавление завершено"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Ошибка при массовом добавлении сотрудников: " & Err.Description, vbCritical, "Ошибка"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Process selected range and fill employee data
' @param wsPayments As Worksheet - sheet "Выплаты_Без_Периодов"
' @param selectedRange As Range - selected range
' @return String - report on results
' =============================================
Public Function ProcessSelectedRangeForImport(wsPayments As Worksheet, selectedRange As Range) As String
    On Error GoTo ErrorHandler
    
    Dim cell As Range
    Dim numberValue As String
    Dim staffData As Object
    Dim foundCount As Long
    Dim notFoundCount As Long
    Dim notFoundList As String
    Dim reportText As String
    Dim lastRow As Long
    Dim i As Long
    
    foundCount = 0
    notFoundCount = 0
    notFoundList = ""
    
    Application.ScreenUpdating = False
    
    ' Process each cell in range
    For Each cell In selectedRange.Cells
        numberValue = Trim(CStr(cell.value))
        
        ' Skip empty cells
        If numberValue = "" Then
            GoTo NextCell
        End If
        
        ' Find employee by number (personal or table)
        Set staffData = mdlHelper.FindEmployeeByAnyNumber(numberValue)
        
        If staffData.count > 0 Then
            ' Employee found - fill data
            cell.value = CStr(staffData("Личный номер")) ' Update personal number (normalization)
            wsPayments.Cells(cell.Row, mdlPaymentValidation.COL_FIO).value = CStr(staffData("Лицо")) ' Fill FIO
            foundCount = foundCount + 1
        Else
            ' Not found - add to list
            notFoundCount = notFoundCount + 1
            If notFoundList <> "" Then
                notFoundList = notFoundList & ", "
            End If
            notFoundList = notFoundList & numberValue
        End If
        
NextCell:
    Next cell
    
    ' Automatic numbering in column A
    lastRow = wsPayments.Cells(wsPayments.Rows.count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_NUMBER).value)) = "" Then
            wsPayments.Cells(i, mdlPaymentValidation.COL_NUMBER).value = i - 1
        End If
    Next i
    
    ' Generate report
    reportText = "Результаты массового добавления:" & vbCrLf & vbCrLf
    reportText = reportText & "Найдено и добавлено: " & foundCount & vbCrLf
    reportText = reportText & "Не найдено: " & notFoundCount
    
    If notFoundCount > 0 And Len(notFoundList) < 200 Then
        reportText = reportText & vbCrLf & "Номера, которые не найдены: " & notFoundList
    ElseIf notFoundCount > 0 Then
        reportText = reportText & vbCrLf & "(Список не найденных номеров слишком длинный для отображения)"
    End If
    
    ProcessSelectedRangeForImport = reportText
    Exit Function
    
ErrorHandler:
    ProcessSelectedRangeForImport = "Ошибка при обработке диапазона: " & Err.Description
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Convert PaymentWithoutPeriod to Dictionary (for storing in Collection)
' @param payment As PaymentWithoutPeriod - payment data
' @return Object (Dictionary) - dictionary with payment data
' =============================================
Private Function PaymentToDictionary(ByRef payment As PaymentWithoutPeriod) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict("fio") = payment.fio
    dict("lichniyNomer") = payment.lichniyNomer
    dict("Rank") = payment.Rank
    dict("Position") = payment.Position
    dict("VoinskayaChast") = payment.VoinskayaChast
    dict("paymentType") = payment.paymentType
    dict("amount") = payment.amount
    dict("foundation") = payment.foundation
    Set PaymentToDictionary = dict
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Convert Dictionary back to PaymentWithoutPeriod
' @param dict As Object (Dictionary) - dictionary with payment data
' @return PaymentWithoutPeriod - payment data
' =============================================
Private Function DictionaryToPayment(ByRef dict As Object) As PaymentWithoutPeriod
    Dim payment As PaymentWithoutPeriod
    If dict.count > 0 Then
        payment.fio = CStr(dict("fio"))
        payment.lichniyNomer = CStr(dict("lichniyNomer"))
        payment.Rank = CStr(dict("Rank"))
        payment.Position = CStr(dict("Position"))
        payment.VoinskayaChast = CStr(dict("VoinskayaChast"))
        payment.paymentType = CStr(dict("paymentType"))
        payment.amount = CStr(dict("amount"))
        payment.foundation = CStr(dict("foundation"))
    End If
    DictionaryToPayment = payment
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for exporting allowances
' =============================================
Public Sub ExportPaymentsWithoutPeriods()
    On Error GoTo ErrorHandler
    
    Dim payments As Collection
    Dim groupedPayments As Object
    Dim paymentType As Variant
    Dim paymentList As Collection
    Dim successCount As Long
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Сбор данных о надбавках..."
    
    ' Collect all payment data
    Set payments = CollectPaymentsData()
    
    If payments.count = 0 Then
        MsgBox "Нет данных для экспорта в листе 'Выплаты_Без_Периодов'.", vbExclamation, "Экспорт"
        GoTo CleanUp
    End If
    
    Application.StatusBar = "Группировка по типам выплат..."
    
    ' Group by payment type
    Set groupedPayments = GroupPaymentsByType(payments)
    
    Application.StatusBar = "Генерация приказов..."
    
    ' Generate orders for each payment type
    successCount = 0
    For Each paymentType In groupedPayments.keys
        Set paymentList = groupedPayments(paymentType)
        If GeneratePaymentOrder(CStr(paymentType), paymentList) Then
            successCount = successCount + 1
        End If
    Next paymentType
    
    MsgBox "Экспорт завершен. Создано приказов: " & successCount & " из " & groupedPayments.count, vbInformation, "Экспорт"
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Ошибка при экспорте надбавок: " & Err.Description, vbCritical, "Ошибка"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Collect all payment data from sheet "Выплаты_Без_Периодов"
' @return Collection - collection of PaymentWithoutPeriod objects
' =============================================
Public Function CollectPaymentsData() As Collection
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As Collection
    Dim payment As PaymentWithoutPeriod
    Dim staffData As Object
    
    Set result = New Collection
    
    ' Find sheet "Выплаты_Без_Периодов"
    Set wsPayments = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
            Set wsPayments = ws
            Exit For
        End If
    Next ws
    
    If wsPayments Is Nothing Then
        Set CollectPaymentsData = result
        Exit Function
    End If
    
    lastRow = wsPayments.Cells(wsPayments.Rows.count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row
    
    If lastRow < 2 Then
        Set CollectPaymentsData = result
        Exit Function
    End If
    
    ' Collect data from each row
    For i = 2 To lastRow
        payment.fio = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_FIO).value))
        payment.lichniyNomer = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_LICHNIY_NOMER).value))
        payment.paymentType = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_PAYMENT_TYPE).value))
        payment.amount = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_AMOUNT).value))
        payment.foundation = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_FOUNDATION).value))
        
        ' Skip empty rows
        If payment.lichniyNomer = "" Then
            GoTo NextRow
        End If
        
        ' Get serviceman data from "Staff" sheet
        Set staffData = mdlHelper.GetStaffData(payment.lichniyNomer, True)
        If staffData.count > 0 Then
            payment.Rank = CStr(staffData("Воинское звание"))
            payment.Position = CStr(staffData("Штатная должность"))
            payment.VoinskayaChast = mdlHelper.ExtractVoinskayaChast(CStr(staffData("Часть")))
        Else
            ' If data not found, use values from table or placeholders
            payment.Rank = "Звание не найдено"
            payment.Position = "Должность не найдена"
            payment.VoinskayaChast = ""
        End If
        
        ' If FIO is not specified in table, take from staff
        If payment.fio = "" And staffData.count > 0 Then
            payment.fio = CStr(staffData("Лицо"))
        End If
        
        ' Convert UDT to Dictionary for storage in Collection
        result.Add PaymentToDictionary(payment)
        
NextRow:
    Next i
    
    Set CollectPaymentsData = result
    Exit Function
    
ErrorHandler:
    Set CollectPaymentsData = New Collection
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Group payments by type
' @param payments As Collection - collection of payments
' @return Object (Dictionary) - dictionary where key = payment type, value = collection of payments
' =============================================
Public Function GroupPaymentsByType(ByVal payments As Collection) As Object
    On Error GoTo ErrorHandler
    
    Dim result As Object
    Dim payment As PaymentWithoutPeriod
    Dim paymentType As String
    Dim paymentList As Collection
    
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim paymentDict As Object
    For i = 1 To payments.count
        ' Extract Dictionary from Collection and convert to UDT
        Set paymentDict = payments(i)
        payment = DictionaryToPayment(paymentDict)
        paymentType = Trim(LCase(payment.paymentType))
        
        If paymentType = "" Then
            paymentType = "Не указан"
        End If
        
        If Not result.exists(paymentType) Then
            Set paymentList = New Collection
            result.Add paymentType, paymentList
        Else
            Set paymentList = result(paymentType)
        End If
        
        ' Add Dictionary back to paymentList
        paymentList.Add paymentDict
    Next i
    
    Set GroupPaymentsByType = result
    Exit Function
    
ErrorHandler:
    Set GroupPaymentsByType = CreateObject("Scripting.Dictionary")
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate Word order for specific payment type
' @param paymentType As String - payment type
' @param payments As Collection - collection of payments of this type
' @return Boolean - True if order successfully created
' =============================================
Public Function GeneratePaymentOrder(ByVal paymentType As String, ByVal payments As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim doc As Object
    Dim templateDoc As Object
    Dim config As PaymentTypeConfig
    Dim templatePath As String
    Dim payment As PaymentWithoutPeriod
    Dim i As Long
    Dim fileName As String
    Dim savePath As String
    Dim wordWasNotRunning As Boolean
    Dim successCount As Long
    Dim paymentDict As Object
    Dim templateRange As Object
    Dim newRange As Object
    Dim endRange As Object
    
    ' Get payment type configuration
    config = mdlPaymentTypes.GetPaymentTypeConfig(paymentType)
    
    ' Get template path with priority
    templatePath = mdlPaymentTypes.GetTemplatePathWithFallback(config)
    
    ' Create Word application
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        wordWasNotRunning = True
    Else
        wordWasNotRunning = False
    End If
    On Error GoTo ErrorHandler
    
    wordApp.Visible = True
    
    ' Create one document from template
    If templatePath <> "" Then
        ' Open template to copy content
        Set templateDoc = wordApp.Documents.Open(templatePath)
        ' Create new document
        Set doc = wordApp.Documents.Add
        ' Copy template content to new document (for first record)
        templateDoc.content.Copy
        doc.content.Paste
        ' Close template
        templateDoc.Close False
        Set templateDoc = Nothing
    Else
        Set doc = wordApp.Documents.Add
        ' Set default font
        With doc.Styles(1).Font
            .Name = "Times New Roman"
            .Size = 12
        End With
    End If
    
    successCount = 0
    
    ' Add all records to one document
    For i = 1 To payments.count
        ' Extract Dictionary from Collection and convert to UDT
        Set paymentDict = payments(i)
        payment = DictionaryToPayment(paymentDict)
        
        If templatePath <> "" Then
            ' For each record create a copy of template with marker replacement
            If i = 1 Then
                ' First record - use already created document
                If FillPaymentTemplate(doc, payment) Then
                    successCount = successCount + 1
                End If
            Else
                ' For other records open template, replace markers and add to document
                Set templateDoc = wordApp.Documents.Open(templatePath)
                
                ' Replace markers in template
                If FillPaymentTemplate(templateDoc, payment) Then
                    ' Copy template content
                    templateDoc.content.Copy
                    
                    ' Paste at the end of main document
                    Set endRange = doc.Range
                    endRange.Collapse Direction:=0 ' wdCollapseEnd
                    ' Add break between records
                    If i > 1 Then
                        endRange.InsertAfter vbCrLf & vbCrLf
                        endRange.Collapse Direction:=0
                    End If
                    endRange.Paste
                    
                    successCount = successCount + 1
                End If
                
                ' Close template without saving
                templateDoc.Close False
                Set templateDoc = Nothing
            End If
        Else
            ' If no template, add text directly
            If i > 1 Then
                Set endRange = doc.Range
                endRange.Collapse Direction:=0
                endRange.InsertAfter vbCrLf & vbCrLf
            End If
            
            If GeneratePaymentTextDirectly(doc, payment) Then
                successCount = successCount + 1
            End If
        End If
    Next i
    
    ' Generate filename for order
    Dim cleanTypeName As String
    cleanTypeName = Replace(Replace(Replace(paymentType, " ", "_"), "/", "_"), "\", "_")
    fileName = "Приказ_" & cleanTypeName & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    If config.TypeCode <> "" Then
        fileName = "Приказ_" & config.TypeCode & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    End If
    savePath = ThisWorkbook.Path & "\" & fileName
    
    ' Save document
    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
    doc.Activate
    
    ' Close Word only if we created it
    If wordWasNotRunning And Not wordApp Is Nothing Then
        ' Leave document open, but do not close Word
    End If
    
    MsgBox "Создан приказ с " & successCount & " записями из " & payments.count, vbInformation, "Экспорт завершен"
    
    GeneratePaymentOrder = (successCount > 0)
    Exit Function
    
ErrorHandler:
    GeneratePaymentOrder = False
    If Not templateDoc Is Nothing Then templateDoc.Close False
    If Not doc Is Nothing Then doc.Close False
    If wordWasNotRunning And Not wordApp Is Nothing Then wordApp.Quit False
    MsgBox "Ошибка при создании приказа: " & Err.Description, vbCritical, "Ошибка"
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Fill Word template with payment data
' @param doc As Object - Word document
' @param payment As PaymentWithoutPeriod - payment data
' @return Boolean - True if successful
' =============================================
Public Function FillPaymentTemplate(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Object
    
    ' Replace placeholders in template
    With doc.content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' [ФИО]
        .Text = "[ФИО]"
        .Replacement.Text = payment.fio
        .Execute Replace:=2
        
        ' [ФИО_ИМЕНИТЕЛЬНЫЙ]
        .Text = "[ФИО_ИМЕНИТЕЛЬНЫЙ]"
        .Replacement.Text = mdlHelper.SklonitFIO(payment.fio)
        .Execute Replace:=2
        
        ' [ЗВАНИЕ]
        .Text = "[ЗВАНИЕ]"
        .Replacement.Text = payment.Rank
        .Execute Replace:=2
        
        ' [ЗВАНИЕ_СКЛОНЕННОЕ]
        .Text = "[ЗВАНИЕ_СКЛОНЕННОЕ]"
        .Replacement.Text = mdlHelper.SklonitZvanie(payment.Rank)
        .Execute Replace:=2
        
        ' [ЛИЧНЫЙ_НОМЕР]
        .Text = "[ЛИЧНЫЙ_НОМЕР]"
        .Replacement.Text = payment.lichniyNomer
        .Execute Replace:=2
        
        ' [ДОЛЖНОСТЬ]
        .Text = "[ДОЛЖНОСТЬ]"
        .Replacement.Text = payment.Position
        .Execute Replace:=2
        
        ' [ДОЛЖНОСТЬ_СКЛОНЕННАЯ]
        .Text = "[ДОЛЖНОСТЬ_СКЛОНЕННАЯ]"
        .Replacement.Text = mdlHelper.SklonitDolzhnost(payment.Position, payment.VoinskayaChast)
        .Execute Replace:=2
        
        ' [РАЗМЕР]
        .Text = "[РАЗМЕР]"
        .Replacement.Text = payment.amount
        .Execute Replace:=2
        
        ' [ОСНОВАНИЕ]
        .Text = "[ОСНОВАНИЕ]"
        .Replacement.Text = payment.foundation
        .Execute Replace:=2
    End With
    
    FillPaymentTemplate = True
    Exit Function
    
ErrorHandler:
    FillPaymentTemplate = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate order text directly in Word without template
' @param doc As Object - Word document
' @param payment As PaymentWithoutPeriod - payment data
' @return Boolean - True if successful
' =============================================
Public Function GeneratePaymentTextDirectly(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Object
    Dim textLine As String
    
    ' Form order text
    textLine = mdlHelper.SklonitZvanie(payment.Rank) & " " & _
               mdlHelper.SklonitFIO(payment.fio) & ", личный номер " & payment.lichniyNomer & ", " & _
               mdlHelper.SklonitDolzhnost(payment.Position, payment.VoinskayaChast) & vbCrLf
    textLine = textLine & "Размер: " & payment.amount & vbCrLf
    textLine = textLine & "Основание: " & payment.foundation & vbCrLf & vbCrLf
    
    ' Insert text into document
    Set rng = doc.Range
    rng.Collapse Direction:=0
    rng.Text = textLine
    rng.Font.Name = "Times New Roman"
    rng.Font.Size = 12
    
    GeneratePaymentTextDirectly = True
    Exit Function
    
ErrorHandler:
    GeneratePaymentTextDirectly = False
End Function

