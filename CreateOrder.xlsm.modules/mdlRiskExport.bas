Attribute VB_Name = "mdlRiskExport"
'==============================================================
' Модуль формирования приказа о надбавке за риск (2% в день, макс 60% в месяц)
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' Версия: 1.2 от 01.12.2025
' Описание: Полный модуль с интеграцией глобальных функций поиска и корректным форматированием
'==============================================================
Option Explicit

'/** Тип для хранения данных о периоде риска с разбивкой по месяцам (ПУБЛИЧНЫЙ) */
Public Type RiskPeriod
    StartDate As Date
    EndDate As Date
    DaysCount As Long
    PercentValue As Double
    MonthYear As String ' Формат: "Февраль 2025"
    PeriodString As String ' "с 01.02 по 20.02, с 25.02 по 28.02"
    IsExpired As Boolean ' Превышен срок 3 года 6 месяцев
End Type

'/** Тип для хранения данных о сотруднике с периодами риска (ПУБЛИЧНЫЙ) */
Public Type EmployeeRiskData
    fio As String
    lichniyNomer As String
    Rank As String
    Position As String
    VoinskayaChast As String
    periods() As RiskPeriod
    periodCount As Long
End Type

'/**
'* ExportRiskAllowanceOrder — главная процедура формирования приказа за риск
'* Выгружает приказ в Word с разбивкой периодов по месяцам
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Public Sub ExportRiskAllowanceOrder()
    On Error GoTo ErrorHandler
    
    ' Проверка критических ошибок
    If Not ValidateRiskData() Then Exit Sub
    
    ' Сбор данных по всем сотрудникам
    Dim employees() As EmployeeRiskData
    Dim empCount As Long
    empCount = CollectRiskEmployeesData(employees)
    
    If empCount = 0 Then
        MsgBox "Нет данных для формирования приказа за риск.", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Формирование Word-документа
    Call GenerateRiskWordDocument(employees, empCount)
    
    MsgBox "Приказ о надбавке за риск успешно сформирован!", vbInformation, "Готово"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при формировании приказа за риск: " & Err.Description, vbCritical, "Ошибка"
End Sub

'/**
'* ValidateRiskData — проверка критических ошибок перед формированием приказа
'* @return Boolean — True если данные корректны, False если есть критические ошибки
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Function ValidateRiskData() As Boolean
    Dim wsDSO As Worksheet, wsStaff As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    If wsDSO Is Nothing Or wsStaff Is Nothing Then
        MsgBox "Необходимые листы 'ДСО' или 'Штат' не найдены.", vbCritical, "Критическая ошибка"
        ValidateRiskData = False
        Exit Function
    End If
    
    ' Проверка: есть ли данные в ДСО
    Dim lastRow As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Нет данных в листе ДСО для формирования приказа.", vbCritical, "Критическая ошибка"
        ValidateRiskData = False
        Exit Function
    End If
    
    ValidateRiskData = True
End Function

'/**
'* CollectRiskEmployeesData — сбор данных по всем сотрудникам с периодами риска (ПУБЛИЧНАЯ)
'* Разбивает периоды по месяцам и вычисляет процент надбавки (макс 60% в месяц)
'* @param employees() — массив структур EmployeeRiskData для заполнения
'* @return Long — количество собранных сотрудников
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Public Function CollectRiskEmployeesData(ByRef employees() As EmployeeRiskData) As Long
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim lastRow As Long, i As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    Dim uniqueLN As Object
    Set uniqueLN = CreateObject("Scripting.Dictionary")
    
    ' Сбор уникальных личных номеров
    For i = 2 To lastRow
        Dim ln As String
        ln = Trim(wsDSO.Cells(i, 3).value)
        If ln <> "" And Not uniqueLN.exists(ln) Then
            uniqueLN.Add ln, ln
        End If
    Next i
    
    ' Инициализация массива сотрудников
    Dim empCount As Long
    empCount = uniqueLN.count
    
    If empCount = 0 Then
        CollectRiskEmployeesData = 0
        Exit Function
    End If
    
    ReDim employees(1 To empCount)
    
    ' Заполнение данных по каждому сотруднику
    Dim empIndex As Long
    empIndex = 1
    
    Dim lnKey As Variant
    For Each lnKey In uniqueLN.Keys
        Call FillEmployeeRiskData(CStr(lnKey), employees(empIndex))
        empIndex = empIndex + 1
    Next lnKey
    
    CollectRiskEmployeesData = empCount
End Function

'/**
'* FillEmployeeRiskData — заполнение данных о сотруднике и его периодах риска
'* Использует глобальную функцию mdlHelper.GetStaffData для поиска данных
'* @param lichniyNomer — личный номер сотрудника
'* @param emp — структура EmployeeRiskData для заполнения
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub FillEmployeeRiskData(ByVal lichniyNomer As String, ByRef emp As EmployeeRiskData)
    emp.lichniyNomer = lichniyNomer
    
    ' Использование глобальной функции поиска
    Dim staffData As Object
    Set staffData = mdlHelper.GetStaffData(lichniyNomer, True)
    
    If staffData.count > 0 Then
        emp.fio = staffData("Лицо")
        emp.Rank = staffData("Воинское звание")
        emp.Position = staffData("Штатная должность")
        emp.VoinskayaChast = mdlHelper.ExtractVoinskayaChast(staffData("Часть"))
    Else
        emp.fio = "ФИО не найдено"
        emp.Rank = "Звание не найдено"
        emp.Position = "Должность не найдена"
        emp.VoinskayaChast = ""
    End If
    
    ' Сбор всех периодов сотрудника
    Dim rawPeriods() As RiskPeriod
    Dim rawCount As Long
    rawCount = CollectRawRiskPeriods(lichniyNomer, rawPeriods)
    
    ' Разбивка периодов по месяцам и объединение
    Dim monthlyPeriods() As RiskPeriod
    Dim monthlyCount As Long
    monthlyCount = SplitAndMergePeriodsByMonth(rawPeriods, rawCount, monthlyPeriods)
    
    ' Сохранение периодов в структуру сотрудника
    emp.periodCount = monthlyCount
    If monthlyCount > 0 Then
        ReDim emp.periods(1 To monthlyCount)
        Dim i As Long
        For i = 1 To monthlyCount
            emp.periods(i) = monthlyPeriods(i)
        Next i
    End If
End Sub

'/**
'* CollectRawRiskPeriods — сбор всех периодов сотрудника из листа ДСО
'* Безопасное чтение дат с проверкой типов
'* @param lichniyNomer — личный номер сотрудника
'* @param periods() — массив RiskPeriod для заполнения
'* @return Long — количество собранных периодов
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Function CollectRawRiskPeriods(ByVal lichniyNomer As String, ByRef periods() As RiskPeriod) As Long
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim lastRow As Long, i As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    If lastRow < 2 Then
        CollectRawRiskPeriods = 0
        Exit Function
    End If
    
    Dim tempPeriods() As RiskPeriod
    ReDim tempPeriods(1 To lastRow * 10)
    Dim pCount As Long
    pCount = 0
    
    ' Срок актуальности: 3 года 6 месяцев (42 месяца)
    Dim expirationDate As Date
    expirationDate = DateAdd("m", -42, Date)
    
    For i = 2 To lastRow
        If Trim(CStr(wsDSO.Cells(i, 3).value)) = lichniyNomer Then
            Dim lastCol As Long
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            
            Dim j As Long
            ' Периоды идут парами начиная с 5-го столбца
            For j = 5 To lastCol Step 2
                Dim startVal As Variant, endVal As Variant
                startVal = wsDSO.Cells(i, j).value
                endVal = wsDSO.Cells(i, j + 1).value
                
                If IsDate(startVal) And IsDate(endVal) Then
                    Dim StartDate As Date, EndDate As Date
                    StartDate = CDate(startVal)
                    EndDate = CDate(endVal)
                    
                    If StartDate <= EndDate Then
                        pCount = pCount + 1
                        tempPeriods(pCount).StartDate = StartDate
                        tempPeriods(pCount).EndDate = EndDate
                        tempPeriods(pCount).IsExpired = (StartDate < expirationDate)
                    End If
                End If
            Next j
        End If
    Next i
    
    ' Объединение пересекающихся периодов
    If pCount > 0 Then
        Dim mergedPeriods() As RiskPeriod
        Dim mergedCount As Long
        
        Dim validPeriods() As RiskPeriod
        ReDim validPeriods(1 To pCount)
        For i = 1 To pCount
            validPeriods(i) = tempPeriods(i)
        Next i
        
        mergedCount = MergeOverlappingRiskPeriods(validPeriods, pCount, mergedPeriods)
        
        ReDim periods(1 To mergedCount)
        For i = 1 To mergedCount
            periods(i) = mergedPeriods(i)
        Next i
        CollectRawRiskPeriods = mergedCount
    Else
        CollectRawRiskPeriods = 0
    End If
End Function

'/**
'* MergeOverlappingRiskPeriods — объединение пересекающихся периодов
'* @param rawPeriods() — исходные периоды
'* @param rawCount — количество исходных периодов
'* @param merged() — результирующий массив объединённых периодов
'* @return Long — количество объединённых периодов
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Function MergeOverlappingRiskPeriods(ByRef rawPeriods() As RiskPeriod, ByVal rawCount As Long, ByRef merged() As RiskPeriod) As Long
    If rawCount = 0 Then
        MergeOverlappingRiskPeriods = 0
        Exit Function
    End If
    
    ' Сортировка по дате начала
    Call SortRiskPeriods(rawPeriods, 1, rawCount)
    
    ReDim merged(1 To rawCount)
    Dim mergedCount As Long
    mergedCount = 1
    merged(1) = rawPeriods(1)
    
    Dim i As Long
    For i = 2 To rawCount
        If rawPeriods(i).StartDate <= merged(mergedCount).EndDate + 1 Then
            ' Пересечение или смежность — объединяем
            If rawPeriods(i).EndDate > merged(mergedCount).EndDate Then
                merged(mergedCount).EndDate = rawPeriods(i).EndDate
            End If
            ' Если хотя бы один период просрочен — весь объединённый просрочен
            If rawPeriods(i).IsExpired Then
                merged(mergedCount).IsExpired = True
            End If
        Else
            ' Новый период
            mergedCount = mergedCount + 1
            merged(mergedCount) = rawPeriods(i)
        End If
    Next i
    
    MergeOverlappingRiskPeriods = mergedCount
End Function

'/**
'* SortRiskPeriods — сортировка периодов по дате начала (пузырьковая сортировка)
'* @param arr() — массив RiskPeriod
'* @param leftIndex — начальный индекс
'* @param rightIndex — конечный индекс
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub SortRiskPeriods(ByRef arr() As RiskPeriod, ByVal leftIndex As Long, ByVal rightIndex As Long)
    Dim i As Long, j As Long
    Dim temp As RiskPeriod
    
    For i = leftIndex To rightIndex - 1
        For j = i + 1 To rightIndex
            If arr(i).StartDate > arr(j).StartDate Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

'/**
'* SplitAndMergePeriodsByMonth — группировка периодов по месяцам с сохранением разрывов дат
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Function SplitAndMergePeriodsByMonth(ByRef rawPeriods() As RiskPeriod, ByVal rawCount As Long, ByRef monthlyPeriods() As RiskPeriod) As Long
    If rawCount = 0 Then
        SplitAndMergePeriodsByMonth = 0
        Exit Function
    End If
    
    ' Словарь для группировки по месяцам: Key = "YYYYMM", Item = Index in monthlyPeriods
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim result() As RiskPeriod
    Dim resCount As Long
    resCount = 0
    
    Dim i As Long
    For i = 1 To rawCount
        ' Разбиваем исходный период, если он переходит через границу месяца
        Dim curDate As Date
        curDate = rawPeriods(i).StartDate
        
        Do While curDate <= rawPeriods(i).EndDate
            Dim endOfMonth As Date
            endOfMonth = DateSerial(Year(curDate), Month(curDate) + 1, 0)
            
            Dim segEnd As Date
            If rawPeriods(i).EndDate < endOfMonth Then
                segEnd = rawPeriods(i).EndDate
            Else
                segEnd = endOfMonth
            End If
            
            Dim segDays As Long
            segDays = DateDiff("d", curDate, segEnd) + 1
            
            Dim key As String
            key = Format(curDate, "yyyymm")
            
            ' Формируем строку для текущего отрезка
            Dim segStr As String
            segStr = "с " & Format(curDate, "dd.mm.yyyy") & " по " & Format(segEnd, "dd.mm.yyyy")
            
            If Not dict.exists(key) Then
                resCount = resCount + 1
                ReDim Preserve result(1 To resCount)
                result(resCount).StartDate = curDate ' Первая дата месяца
                result(resCount).EndDate = segEnd   ' Последняя дата месяца (пока)
                result(resCount).DaysCount = segDays
                result(resCount).MonthYear = Format(curDate, "mmmm yyyy")
                result(resCount).PeriodString = segStr
                result(resCount).IsExpired = rawPeriods(i).IsExpired
                dict.Add key, resCount
            Else
                Dim idx As Long
                idx = dict(key)
                result(idx).DaysCount = result(idx).DaysCount + segDays
                result(idx).PeriodString = result(idx).PeriodString & ", " & segStr
                
                ' Обновляем границы месяца (для сортировки)
                If curDate < result(idx).StartDate Then result(idx).StartDate = curDate
                If segEnd > result(idx).EndDate Then result(idx).EndDate = segEnd
                If rawPeriods(i).IsExpired Then result(idx).IsExpired = True
            End If
            
            curDate = endOfMonth + 1
        Loop
    Next i
    
    ' Финальный расчет процентов
    For i = 1 To resCount
        Dim pct As Double
        pct = result(i).DaysCount * 2
        If pct > 60 Then pct = 60
        result(i).PercentValue = pct
    Next i
    
    ' Сортировка по дате
    Call SortRiskPeriods(result, 1, resCount)
    
    If resCount > 0 Then
        monthlyPeriods = result
    End If
    
    SplitAndMergePeriodsByMonth = resCount
End Function

'/**
'* GenerateRiskWordDocument — формирование Word-документа с корректным форматированием
'* Использует Times New Roman 12, добавляет основание из ДСО
'* @param employees() — массив данных о сотрудниках
'* @param empCount — количество сотрудников
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub GenerateRiskWordDocument(ByRef employees() As EmployeeRiskData, ByVal empCount As Long)
    Dim wordApp As Object, doc As Object, rng As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add
    
    ' Установка шрифта Times New Roman по умолчанию для всего документа
    With doc.Styles(1).Font
        .name = "Times New Roman"
        .Size = 12
    End With
    
    ' Заголовок приказа
    Set rng = doc.Range
    rng.text = "ПРИКАЗ" & vbCrLf
    rng.Font.Bold = True
    rng.Font.Size = 12
    rng.Font.name = "Times New Roman"
    rng.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    
    ' Нумерация сотрудников
    Dim i As Long, j As Long
    For i = 1 To empCount
        Set rng = doc.Range
        rng.Collapse Direction:=0
        
        ' Формирование заголовка сотрудника
        Dim header As String
        header = i & ". " & mdlHelper.SklonitZvanie(employees(i).Rank) & " " & _
                 mdlHelper.SklonitFIO(employees(i).fio) & ", личный номер " & employees(i).lichniyNomer & ", " & _
                 mdlHelper.SklonitDolzhnost(employees(i).Position, employees(i).VoinskayaChast) & vbCrLf
        
        rng.text = header
        rng.Font.Bold = False
        rng.Font.Size = 12
        rng.Font.name = "Times New Roman"
        rng.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
        
  ' Периоды
        For j = 1 To employees(i).periodCount
            Set rng = doc.Range
            rng.Collapse Direction:=0
            
            ' Используем сформированную строку периодов
            Dim periodText As String
            periodText = "   - " & employees(i).periods(j).PeriodString & _
                         " (" & employees(i).periods(j).DaysCount & " дн.) = " & _
                         employees(i).periods(j).PercentValue & "%" & vbCrLf
            
            rng.text = periodText
            
            rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
            rng.Font.Bold = False
            rng.Font.Size = 12
            rng.Font.name = "Times New Roman"
            
            ' Предупреждение о просроченности
            If employees(i).periods(j).IsExpired Then
                Set rng = doc.Range
                rng.Collapse Direction:=0
                rng.text = "      ВНИМАНИЕ: Период с " & Format(employees(i).periods(j).StartDate, "dd.mm.yyyy") & _
                          " по " & Format(employees(i).periods(j).EndDate, "dd.mm.yyyy") & _
                          " превышает срок в три года на обращение за получением надбавки согласно ПМО 727" & vbCrLf
                rng.Font.Color = RGB(255, 0, 0)
                rng.Font.Bold = True
                rng.Font.Size = 12
                rng.Font.name = "Times New Roman"
                rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
            End If
        Next j
        
        ' Добавление основания из ДСО
        Call AddRiskFoundationFromDSO(doc, employees(i).lichniyNomer)
        
        Set rng = doc.Range
        rng.Collapse Direction:=0
        rng.InsertAfter vbCrLf ' Пустая строка между сотрудниками
    Next i
    
    ' Сохранение
    Dim savePath As String
    savePath = ThisWorkbook.Path & "\ПриказЗаРиск_" & Format(Date, "dd.mm.yyyy") & ".docx"
    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
End Sub


'/**
'* AddRiskFoundationFromDSO — добавление строки основания из столбца 4 листа ДСО
'* @param doc — объект Word.Document
'* @param lichniyNomer — личный номер сотрудника
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub AddRiskFoundationFromDSO(ByVal doc As Object, ByVal lichniyNomer As String)
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim lastRow As Long, i As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    Dim foundationText As String
    foundationText = ""
    
    ' Поиск оснований в столбце 4 для данного личного номера
    For i = 2 To lastRow
        If Trim(CStr(wsDSO.Cells(i, 3).value)) = lichniyNomer Then
            Dim basis As String
            basis = Trim(CStr(wsDSO.Cells(i, 4).value))
            
            If basis <> "" Then
                If foundationText <> "" Then
                    foundationText = foundationText & "; " & basis
                Else
                    foundationText = basis
                End If
            End If
        End If
    Next i
    
    ' Добавление основания в документ
    If foundationText <> "" Then
        Dim rng As Object
        Set rng = doc.Range
        rng.Collapse Direction:=0
        
        Dim foundationLine As String
        foundationLine = "Основание: " & foundationText & vbCrLf
        
        rng.text = foundationLine
        rng.Font.Size = 12
        rng.Font.name = "Times New Roman"
        rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
    End If
End Sub



