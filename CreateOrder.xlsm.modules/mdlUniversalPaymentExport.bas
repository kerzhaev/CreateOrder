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

    If modActivation.GetLicenseStatus() = 1 Then Exit Sub
    
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
    Dim endRange As Object
    Dim isListTemplate As Boolean
    Dim listText As String
    
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
    
    ' Determine if we are using a List Template
    isListTemplate = False
    If templatePath <> "" Then
        Set doc = wordApp.Documents.Add(templatePath)
        ' Проверяем наличие маркера списка
        With doc.content.Find
            .Text = "[СПИСОК_ВОЕННОСЛУЖАЩИХ]"
            If .Execute Then
                isListTemplate = True
            End If
        End With
        
        If Not isListTemplate Then
            ' Если это не списочный шаблон, используем старую логику постраничного копирования
            doc.Close False
            Set doc = wordApp.Documents.Add
            
            Set templateDoc = wordApp.Documents.Open(templatePath)
            templateDoc.content.Copy
            doc.content.Paste
            templateDoc.Close False
            Set templateDoc = Nothing
        End If
    Else
        Set doc = wordApp.Documents.Add
        With doc.Styles(1).Font
            .Name = "Times New Roman"
            .Size = 12
        End With
    End If
    
    successCount = 0
    
    If isListTemplate Then
        ' ==========================================
        ' НОВАЯ ЛОГИКА: ШАБЛОН СО СПИСКОМ
        ' ==========================================
        listText = ""
        For i = 1 To payments.count
            Set paymentDict = payments(i)
            payment = DictionaryToPayment(paymentDict)
            
            ' Используем нашу новую функцию форматирования
            Dim empText As String
            empText = FormatEmployeePaymentText(payment, i)
            
            ' Разделитель между военнослужащими (пустая строка)
            If i < payments.count Then
                empText = empText & vbCrLf
            End If
            
            listText = listText & empText
        Next i
        
        ' Вставляем готовый список на место маркера, обходя ограничение в 255 символов
        Dim rngFind As Object
        Set rngFind = doc.content
        With rngFind.Find
            .ClearFormatting
            .Text = "[СПИСОК_ВОЕННОСЛУЖАЩИХ]"
            .Forward = True
            .Wrap = 0 ' wdFindStop
            If .Execute Then
                rngFind.Text = listText
            End If
        End With
        
        successCount = payments.count
        
    Else
        ' ==========================================
        ' СТАРАЯ ЛОГИКА: ИНДИВИДУАЛЬНЫЕ МАРКЕРЫ (ИЛИ БЕЗ ШАБЛОНА)
        ' ==========================================
        For i = 1 To payments.count
            Set paymentDict = payments(i)
            payment = DictionaryToPayment(paymentDict)
            
            If templatePath <> "" Then
                If i = 1 Then
                    If FillPaymentTemplate(doc, payment) Then
                        successCount = successCount + 1
                    End If
                Else
                    Set templateDoc = wordApp.Documents.Open(templatePath)
                    If FillPaymentTemplate(templateDoc, payment) Then
                        templateDoc.content.Copy
                        Set endRange = doc.Range
                        endRange.Collapse Direction:=0
                        If i > 1 Then
                            endRange.InsertAfter vbCrLf & vbCrLf
                            endRange.Collapse Direction:=0
                        End If
                        endRange.Paste
                        successCount = successCount + 1
                    End If
                    templateDoc.Close False
                    Set templateDoc = Nothing
                End If
            Else
                If i > 1 Then
                    Set endRange = doc.Range
                    endRange.Collapse Direction:=0
                    endRange.InsertAfter vbCrLf & vbCrLf
                End If
                
                ' Передаем i для нумерации
                If GeneratePaymentTextDirectly(doc, payment, i) Then
                    successCount = successCount + 1
                End If
            End If
        Next i
    End If
    
    ' Сохранение
    Dim cleanTypeName As String
    cleanTypeName = Replace(Replace(Replace(paymentType, " ", "_"), "/", "_"), "\", "_")
    fileName = "Приказ_" & cleanTypeName & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    If config.TypeCode <> "" Then
        fileName = "Приказ_" & config.TypeCode & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    End If
    savePath = ThisWorkbook.Path & "\" & fileName
    
    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
    doc.Activate
    
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
' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Generate order text directly in Word without template
' =============================================
Public Function GeneratePaymentTextDirectly(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod, ByVal index As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Object
    Dim textLine As String
    
    ' Используем общую функцию форматирования
    textLine = FormatEmployeePaymentText(payment, index) & vbCrLf
    
    ' Вставляем текст в документ
    Set rng = doc.Range
    rng.Collapse Direction:=0
    rng.Text = textLine
    rng.Font.Name = "Times New Roman"
    rng.Font.Size = 14
    
    GeneratePaymentTextDirectly = True
    Exit Function
    
ErrorHandler:
    GeneratePaymentTextDirectly = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Формирует готовый текст для одного военнослужащего со всеми проверками
' =============================================
Private Function FormatEmployeePaymentText(ByRef payment As PaymentWithoutPeriod, ByVal index As Long) As String
    Dim cleanFIO As String, cleanRank As String, cleanPos As String, cleanVC As String, cleanFound As String
    
    ' 1. Очищаем все данные от случайных переносов строк (Alt+Enter из Excel)
    cleanFIO = Replace(Replace(payment.fio, vbCr, ""), vbLf, " ")
    cleanRank = Replace(Replace(payment.Rank, vbCr, ""), vbLf, " ")
    cleanPos = Replace(Replace(payment.Position, vbCr, ""), vbLf, " ")
    cleanVC = Replace(Replace(payment.VoinskayaChast, vbCr, ""), vbLf, " ")
    cleanFound = Replace(Replace(payment.foundation, vbCr, ""), vbLf, " ")
    
    ' Убираем двойные пробелы, если они появились после замены
    While InStr(cleanPos, "  ") > 0: cleanPos = Replace(cleanPos, "  ", " "): Wend
    While InStr(cleanFound, "  ") > 0: cleanFound = Replace(cleanFound, "  ", " "): Wend
    
    ' 2. Преобразуем размер (0,3 -> 30) независимым от системы способом
    Dim formattedAmount As String
    Dim numVal As Double
    
    formattedAmount = Trim(payment.amount)
    formattedAmount = Replace(formattedAmount, "%", "") ' Сначала убираем знак процента, если он был
    
    ' Функция Val понимает только точку, поэтому принудительно меняем запятую на точку
    Dim dotAmount As String
    dotAmount = Replace(formattedAmount, ",", ".")
    
    ' Если внутри действительно число
    If IsNumeric(dotAmount) Or IsNumeric(formattedAmount) Then
        numVal = val(dotAmount)
        ' Если число дробное (от 0.01 до 1), умножаем на 100
        If numVal > 0 And numVal <= 1 Then
            formattedAmount = CStr(numVal * 100)
        End If
    End If
    
    ' 3. Формируем текст БЕЗ лишних переносов строк
    Dim textLine As String
    textLine = index & ". " & mdlHelper.SklonitZvanie(cleanRank) & " " & _
               mdlHelper.SklonitFIO(cleanFIO) & ", личный номер " & payment.lichniyNomer & ", " & _
               mdlHelper.SklonitDolzhnost(cleanPos, cleanVC)
               
    ' Добавляем размер через ПРОБЕЛ, продолжая строку
    If formattedAmount <> "" And formattedAmount <> "0" Then
        textLine = textLine & " в размере " & formattedAmount & " процентов оклада по воинской должности."
    End If
    
    ' А вот основание спускаем на новую строку через vbCrLf
    If cleanFound <> "" Then
        textLine = textLine & vbCrLf & "Основание: " & cleanFound
    End If
    
    FormatEmployeePaymentText = textLine
End Function

