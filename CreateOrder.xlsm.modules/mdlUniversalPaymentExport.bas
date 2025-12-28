Attribute VB_Name = "mdlUniversalPaymentExport"
' ===============================================================================
' Модуль mdlUniversalPaymentExport
' Версия: 1.0.0
' Дата: 01.12.2025
' Описание: Универсальный экспорт надбавок без периодов в Word
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' ===============================================================================

Option Explicit

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Преобразование PaymentWithoutPeriod в Dictionary (для хранения в Collection)
' @param payment As PaymentWithoutPeriod - данные о выплате
' @return Object (Dictionary) - словарь с данными выплаты
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Преобразование Dictionary обратно в PaymentWithoutPeriod
' @param dict As Object (Dictionary) - словарь с данными выплаты
' @return PaymentWithoutPeriod - данные о выплате
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Главная функция экспорта надбавок
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
    
    ' Собираем все данные о выплатах
    Set payments = CollectPaymentsData()
    
    If payments.count = 0 Then
        MsgBox "Нет данных для экспорта в листе 'Выплаты_Без_Периодов'.", vbExclamation, "Экспорт"
        GoTo CleanUp
    End If
    
    Application.StatusBar = "Группировка по типам выплат..."
    
    ' Группируем по типу выплаты
    Set groupedPayments = GroupPaymentsByType(payments)
    
    Application.StatusBar = "Генерация приказов..."
    
    ' Генерируем приказы для каждого типа выплаты
    successCount = 0
    For Each paymentType In groupedPayments.Keys
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Сбор всех данных о выплатах из листа "Выплаты_Без_Периодов"
' @return Collection - коллекция объектов PaymentWithoutPeriod
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
    
    ' Ищем лист "Выплаты_Без_Периодов"
    Set wsPayments = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
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
    
    ' Собираем данные из каждой строки
    For i = 2 To lastRow
        payment.fio = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_FIO).value))
        payment.lichniyNomer = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_LICHNIY_NOMER).value))
        payment.paymentType = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_PAYMENT_TYPE).value))
        payment.amount = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_AMOUNT).value))
        payment.foundation = Trim(CStr(wsPayments.Cells(i, mdlPaymentValidation.COL_FOUNDATION).value))
        
        ' Пропускаем пустые строки
        If payment.lichniyNomer = "" Then
            GoTo NextRow
        End If
        
        ' Получаем данные о военнослужащем из листа "Штат"
        Set staffData = mdlHelper.GetStaffData(payment.lichniyNomer, True)
        If staffData.count > 0 Then
            payment.Rank = CStr(staffData("Воинское звание"))
            payment.Position = CStr(staffData("Штатная должность"))
            payment.VoinskayaChast = mdlHelper.ExtractVoinskayaChast(CStr(staffData("Часть")))
        Else
            ' Если данные не найдены, используем значения из таблицы или заглушки
            payment.Rank = "Звание не найдено"
            payment.Position = "Должность не найдена"
            payment.VoinskayaChast = ""
        End If
        
        ' Если ФИО не указано в таблице, берем из штата
        If payment.fio = "" And staffData.count > 0 Then
            payment.fio = CStr(staffData("Лицо"))
        End If
        
        ' Преобразуем UDT в Dictionary для хранения в Collection
        result.Add PaymentToDictionary(payment)
        
NextRow:
    Next i
    
    Set CollectPaymentsData = result
    Exit Function
    
ErrorHandler:
    Set CollectPaymentsData = New Collection
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Группировка выплат по типу
' @param payments As Collection - коллекция выплат
' @return Object (Dictionary) - словарь, где ключ = тип выплаты, значение = коллекция выплат
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
        ' Извлекаем Dictionary из Collection и преобразуем в UDT
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
        
        ' Добавляем Dictionary обратно в paymentList
        paymentList.Add paymentDict
    Next i
    
    Set GroupPaymentsByType = result
    Exit Function
    
ErrorHandler:
    Set GroupPaymentsByType = CreateObject("Scripting.Dictionary")
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Генерация приказа Word для конкретного типа выплаты
' @param paymentType As String - тип выплаты
' @param payments As Collection - коллекция выплат данного типа
' @return Boolean - True если приказ успешно создан
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
    
    ' Получаем конфигурацию типа выплаты
    config = mdlPaymentTypes.GetPaymentTypeConfig(paymentType)
    
    ' Получаем путь к шаблону с учетом приоритета
    templatePath = mdlPaymentTypes.GetTemplatePathWithFallback(config)
    
    ' Создаем Word приложение
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
    
    ' Создаем один документ из шаблона
    If templatePath <> "" Then
        ' Открываем шаблон для копирования содержимого
        Set templateDoc = wordApp.Documents.Open(templatePath)
        ' Создаем новый документ
        Set doc = wordApp.Documents.Add
        ' Копируем содержимое шаблона в новый документ (для первой записи)
        templateDoc.Content.Copy
        doc.Content.Paste
        ' Закрываем шаблон
        templateDoc.Close False
        Set templateDoc = Nothing
    Else
        Set doc = wordApp.Documents.Add
        ' Устанавливаем шрифт по умолчанию
        With doc.Styles(1).Font
            .name = "Times New Roman"
            .Size = 12
        End With
    End If
    
    successCount = 0
    
    ' Добавляем все записи в один документ
    For i = 1 To payments.count
        ' Извлекаем Dictionary из Collection и преобразуем в UDT
        Set paymentDict = payments(i)
        payment = DictionaryToPayment(paymentDict)
        
        If templatePath <> "" Then
            ' Для каждой записи создаем копию шаблона с заменой маркеров
            If i = 1 Then
                ' Первая запись - используем уже созданный документ
                If FillPaymentTemplate(doc, payment) Then
                    successCount = successCount + 1
                End If
            Else
                ' Для остальных записей открываем шаблон, заменяем маркеры и добавляем в документ
                Set templateDoc = wordApp.Documents.Open(templatePath)
                
                ' Заменяем маркеры в шаблоне
                If FillPaymentTemplate(templateDoc, payment) Then
                    ' Копируем содержимое шаблона
                    templateDoc.Content.Copy
                    
                    ' Вставляем в конец основного документа
                    Set endRange = doc.Range
                    endRange.Collapse Direction:=0 ' wdCollapseEnd
                    ' Добавляем разрыв между записями
                    If i > 1 Then
                        endRange.InsertAfter vbCrLf & vbCrLf
                        endRange.Collapse Direction:=0
                    End If
                    endRange.Paste
                    
                    successCount = successCount + 1
                End If
                
                ' Закрываем шаблон без сохранения
                templateDoc.Close False
                Set templateDoc = Nothing
            End If
        Else
            ' Если шаблона нет, добавляем текст напрямую
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
    
    ' Формируем имя файла для приказа
    Dim cleanTypeName As String
    cleanTypeName = Replace(Replace(Replace(paymentType, " ", "_"), "/", "_"), "\", "_")
    fileName = "Приказ_" & cleanTypeName & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    If config.TypeCode <> "" Then
        fileName = "Приказ_" & config.TypeCode & "_" & Format(Date, "dd.mm.yyyy") & ".docx"
    End If
    savePath = ThisWorkbook.Path & "\" & fileName
    
    ' Сохраняем документ
    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
    doc.Activate
    
    ' Закрываем Word только если мы его создали
    If wordWasNotRunning And Not wordApp Is Nothing Then
        ' Оставляем документ открытым, но не закрываем Word
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Заполнение шаблона Word данными о выплате
' @param doc As Object - документ Word
' @param payment As PaymentWithoutPeriod - данные о выплате
' @return Boolean - True если успешно
' =============================================
Public Function FillPaymentTemplate(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Object
    
    ' Замена плейсхолдеров в шаблоне
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        ' [ФИО]
        .text = "[ФИО]"
        .Replacement.text = payment.fio
        .Execute Replace:=2
        
        ' [ФИО_ИМЕНИТЕЛЬНЫЙ]
        .text = "[ФИО_ИМЕНИТЕЛЬНЫЙ]"
        .Replacement.text = mdlHelper.SklonitFIO(payment.fio)
        .Execute Replace:=2
        
        ' [ЗВАНИЕ]
        .text = "[ЗВАНИЕ]"
        .Replacement.text = payment.Rank
        .Execute Replace:=2
        
        ' [ЗВАНИЕ_СКЛОНЕННОЕ]
        .text = "[ЗВАНИЕ_СКЛОНЕННОЕ]"
        .Replacement.text = mdlHelper.SklonitZvanie(payment.Rank)
        .Execute Replace:=2
        
        ' [ЛИЧНЫЙ_НОМЕР]
        .text = "[ЛИЧНЫЙ_НОМЕР]"
        .Replacement.text = payment.lichniyNomer
        .Execute Replace:=2
        
        ' [ДОЛЖНОСТЬ]
        .text = "[ДОЛЖНОСТЬ]"
        .Replacement.text = payment.Position
        .Execute Replace:=2
        
        ' [ДОЛЖНОСТЬ_СКЛОНЕННАЯ]
        .text = "[ДОЛЖНОСТЬ_СКЛОНЕННАЯ]"
        .Replacement.text = mdlHelper.SklonitDolzhnost(payment.Position, payment.VoinskayaChast)
        .Execute Replace:=2
        
        ' [РАЗМЕР]
        .text = "[РАЗМЕР]"
        .Replacement.text = payment.amount
        .Execute Replace:=2
        
        ' [ОСНОВАНИЕ]
        .text = "[ОСНОВАНИЕ]"
        .Replacement.text = payment.foundation
        .Execute Replace:=2
    End With
    
    FillPaymentTemplate = True
    Exit Function
    
ErrorHandler:
    FillPaymentTemplate = False
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Формирование текста приказа напрямую в Word без шаблона
' @param doc As Object - документ Word
' @param payment As PaymentWithoutPeriod - данные о выплате
' @return Boolean - True если успешно
' =============================================
Public Function GeneratePaymentTextDirectly(ByVal doc As Object, ByRef payment As PaymentWithoutPeriod) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rng As Object
    Dim textLine As String
    
    ' Формируем текст приказа
    textLine = mdlHelper.SklonitZvanie(payment.Rank) & " " & _
               mdlHelper.SklonitFIO(payment.fio) & ", личный номер " & payment.lichniyNomer & ", " & _
               mdlHelper.SklonitDolzhnost(payment.Position, payment.VoinskayaChast) & vbCrLf
    textLine = textLine & "Размер: " & payment.amount & vbCrLf
    textLine = textLine & "Основание: " & payment.foundation & vbCrLf & vbCrLf
    
    ' Вставляем текст в документ
    Set rng = doc.Range
    rng.Collapse Direction:=0
    rng.text = textLine
    rng.Font.name = "Times New Roman"
    rng.Font.Size = 12
    
    GeneratePaymentTextDirectly = True
    Exit Function
    
ErrorHandler:
    GeneratePaymentTextDirectly = False
End Function

