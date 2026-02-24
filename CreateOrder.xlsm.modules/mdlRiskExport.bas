Attribute VB_Name = "mdlRiskExport"
'==============================================================
' Module for generating risk allowance order (2% per day, max 60% per month)
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Version: 1.2 from 14.02.2026
' Description: Full module with integration of global search functions and correct formatting
'==============================================================
Option Explicit

'/** Type for storing risk period data broken down by month (PUBLIC) */
Public Type RiskPeriod
    StartDate As Date
    EndDate As Date
    daysCount As Long
    PercentValue As Double
    MonthYear As String ' Format: "February 2025"
    PeriodString As String ' "from 01.02 to 20.02, from 25.02 to 28.02"
    IsExpired As Boolean ' Exceeded 3 years 6 months limit
End Type

'/** Type for storing employee data with risk periods (PUBLIC) */
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
'* ExportRiskAllowanceOrder — main procedure for generating risk order
'* Exports order to Word with period breakdown by month
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Public Sub ExportRiskAllowanceOrder()

    If modActivation.GetLicenseStatus() = 1 Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Check critical errors
    If Not ValidateRiskData() Then Exit Sub
    
    ' Collect data for all employees
    Dim employees() As EmployeeRiskData
    Dim empCount As Long
    empCount = CollectRiskEmployeesData(employees)
    
    If empCount = 0 Then
        MsgBox "Нет данных для формирования приказа за риск.", vbExclamation, "Ошибка"
        Exit Sub
    End If
    
    ' Generate Word document
    Call GenerateRiskWordDocument(employees, empCount)
    
    MsgBox "Приказ о надбавке за риск успешно сформирован!", vbInformation, "Готово"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при формировании приказа за риск: " & Err.Description, vbCritical, "Ошибка"
End Sub

'/**
'* ValidateRiskData — check critical errors before generating order
'* @return Boolean — True if data is valid, False if critical errors exist
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
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
    
    ' Check: is there data in DSO
    Dim lastRow As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Нет данных в листе ДСО для формирования приказа.", vbCritical, "Критическая ошибка"
        ValidateRiskData = False
        Exit Function
    End If
    
    ' --- ДОБАВЛЕН БЛОК ПРОВЕРКИ ОШИБОК ---
    If mdlHelper.hasCriticalErrors() Then
        MsgBox "Экспорт приказа за риск заблокирован из-за критических ошибок в данных!" & vbCrLf & _
               "Исправьте все ошибки (красные ячейки) в листе ДСО.", vbCritical, "Экспорт невозможен"
        ValidateRiskData = False
        Exit Function
    End If
    ' -------------------------------------
    
    ValidateRiskData = True
End Function

'/**
'* CollectRiskEmployeesData — collect data for all employees with risk periods (PUBLIC)
'* Splits periods by month and calculates allowance percentage (max 60% per month)
'* @param employees() — array of EmployeeRiskData structures to fill
'* @return Long — number of collected employees
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Public Function CollectRiskEmployeesData(ByRef employees() As EmployeeRiskData) As Long
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim lastRow As Long, i As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    Dim uniqueLN As Object
    Set uniqueLN = CreateObject("Scripting.Dictionary")
    
    ' Collect unique personal numbers
    For i = 2 To lastRow
        Dim ln As String
        ln = Trim(wsDSO.Cells(i, 3).value)
        If ln <> "" And Not uniqueLN.exists(ln) Then
            uniqueLN.Add ln, ln
        End If
    Next i
    
    ' Initialize employees array
    Dim empCount As Long
    empCount = uniqueLN.count
    
    If empCount = 0 Then
        CollectRiskEmployeesData = 0
        Exit Function
    End If
    
    ReDim employees(1 To empCount)
    
    ' Fill data for each employee
    Dim empIndex As Long
    empIndex = 1
    
    Dim lnKey As Variant
    For Each lnKey In uniqueLN.keys
        Call FillEmployeeRiskData(CStr(lnKey), employees(empIndex))
        empIndex = empIndex + 1
    Next lnKey
    
    CollectRiskEmployeesData = empCount
End Function

'/**
'* FillEmployeeRiskData — fill employee data and risk periods
'* Uses global search function mdlHelper.GetStaffData
'* @param lichniyNomer — employee personal number
'* @param emp — EmployeeRiskData structure to fill
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Private Sub FillEmployeeRiskData(ByVal lichniyNomer As String, ByRef emp As EmployeeRiskData)
    emp.lichniyNomer = lichniyNomer
    
    ' Use global search function
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
    
    ' Collect all employee periods
    Dim rawPeriods() As RiskPeriod
    Dim rawCount As Long
    rawCount = CollectRawRiskPeriods(lichniyNomer, rawPeriods)
    
    ' Split periods by month and merge
    Dim monthlyPeriods() As RiskPeriod
    Dim monthlyCount As Long
    monthlyCount = SplitAndMergePeriodsByMonth(rawPeriods, rawCount, monthlyPeriods)
    
    ' Save periods to employee structure
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
'* CollectRawRiskPeriods — collect all employee periods from DSO sheet
'* Safe date reading with type checking
'* @param lichniyNomer — employee personal number
'* @param periods() — RiskPeriod array to fill
'* @return Long — number of collected periods
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
'/**
'* CollectRawRiskPeriods — collect all employee periods from DSO sheet
'* FIX: Uses mdlHelper.ParseDateSafe to handle dots and 2-digit years correctly
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
    
    ' Срок давности: 3 года 6 месяцев (42 месяца)
    Dim expirationDate As Date
    expirationDate = DateAdd("m", -42, Date)
    
    For i = 2 To lastRow
        If Trim(CStr(wsDSO.Cells(i, 3).value)) = lichniyNomer Then
            Dim lastCol As Long
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            
            Dim j As Long
            ' Периоды идут парами, начиная с 5 колонки
            For j = 5 To lastCol Step 2
                Dim startVal As Variant, endVal As Variant
                ' ВАЖНО: Читаем .Text, чтобы получить именно то, что видит пользователь (с точками)
                startVal = wsDSO.Cells(i, j).Text
                endVal = wsDSO.Cells(i, j + 1).Text
                
                ' Используем наш мощный парсер из mdlHelper
                Dim sDate As Date, eDate As Date
                sDate = mdlHelper.ParseDateSafe(startVal)
                eDate = mdlHelper.ParseDateSafe(endVal)
                
                ' Если даты валидны (функция возвращает > 0 для валидных дат)
                If sDate > 0 And eDate > 0 Then
                    If sDate <= eDate Then
                        pCount = pCount + 1
                        tempPeriods(pCount).StartDate = sDate
                        tempPeriods(pCount).EndDate = eDate
                        tempPeriods(pCount).IsExpired = (sDate < expirationDate)
                    End If
                End If
            Next j
        End If
    Next i
    
    ' Слияние пересекающихся периодов
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
'* MergeOverlappingRiskPeriods — merge overlapping periods
'* @param rawPeriods() — source periods
'* @param rawCount — number of source periods
'* @param merged() — resulting array of merged periods
'* @return Long — number of merged periods
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Private Function MergeOverlappingRiskPeriods(ByRef rawPeriods() As RiskPeriod, ByVal rawCount As Long, ByRef merged() As RiskPeriod) As Long
    If rawCount = 0 Then
        MergeOverlappingRiskPeriods = 0
        Exit Function
    End If
    
    ' Sort by start date
    Call SortRiskPeriods(rawPeriods, 1, rawCount)
    
    ReDim merged(1 To rawCount)
    Dim mergedCount As Long
    mergedCount = 1
    merged(1) = rawPeriods(1)
    
    Dim i As Long
    For i = 2 To rawCount
        If rawPeriods(i).StartDate <= merged(mergedCount).EndDate + 1 Then
            ' Overlap or adjacent — merge
            If rawPeriods(i).EndDate > merged(mergedCount).EndDate Then
                merged(mergedCount).EndDate = rawPeriods(i).EndDate
            End If
            ' If at least one period is expired — whole merged period is expired
            If rawPeriods(i).IsExpired Then
                merged(mergedCount).IsExpired = True
            End If
        Else
            ' New period
            mergedCount = mergedCount + 1
            merged(mergedCount) = rawPeriods(i)
        End If
    Next i
    
    MergeOverlappingRiskPeriods = mergedCount
End Function

'/**
'* SortRiskPeriods — sort periods by start date (bubble sort)
'* @param arr() — RiskPeriod array
'* @param leftIndex — start index
'* @param rightIndex — end index
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
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
'* SplitAndMergePeriodsByMonth — group periods by month preserving date gaps
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Private Function SplitAndMergePeriodsByMonth(ByRef rawPeriods() As RiskPeriod, ByVal rawCount As Long, ByRef monthlyPeriods() As RiskPeriod) As Long
    If rawCount = 0 Then
        SplitAndMergePeriodsByMonth = 0
        Exit Function
    End If
    
    ' Dictionary for monthly grouping: Key = "YYYYMM", Item = Index in monthlyPeriods
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim result() As RiskPeriod
    Dim resCount As Long
    resCount = 0
    
    Dim i As Long
    For i = 1 To rawCount
        ' Split source period if it crosses month boundary
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
            
            ' Form string for current segment
            Dim segStr As String
            segStr = "с " & Format(curDate, "dd.mm.yyyy") & " по " & Format(segEnd, "dd.mm.yyyy")
            
            If Not dict.exists(key) Then
                resCount = resCount + 1
                ReDim Preserve result(1 To resCount)
                result(resCount).StartDate = curDate ' First date of month
                result(resCount).EndDate = segEnd    ' Last date of month (for now)
                result(resCount).daysCount = segDays
                result(resCount).MonthYear = Format(curDate, "mmmm yyyy")
                result(resCount).PeriodString = segStr
                result(resCount).IsExpired = rawPeriods(i).IsExpired
                dict.Add key, resCount
            Else
                Dim idx As Long
                idx = dict(key)
                result(idx).daysCount = result(idx).daysCount + segDays
                result(idx).PeriodString = result(idx).PeriodString & ", " & segStr
                
                ' Update month boundaries (for sorting)
                If curDate < result(idx).StartDate Then result(idx).StartDate = curDate
                If segEnd > result(idx).EndDate Then result(idx).EndDate = segEnd
                If rawPeriods(i).IsExpired Then result(idx).IsExpired = True
            End If
            
            curDate = endOfMonth + 1
        Loop
    Next i
    
    ' Final percentage calculation
    For i = 1 To resCount
        Dim pct As Double
        pct = result(i).daysCount * 2
        If pct > 60 Then pct = 60
        result(i).PercentValue = pct
    Next i
    
    ' Sort by date
    Call SortRiskPeriods(result, 1, resCount)
    
    If resCount > 0 Then
        monthlyPeriods = result
    End If
    
    SplitAndMergePeriodsByMonth = resCount
End Function

'/**
'* GenerateRiskWordDocument — generate Word document with correct formatting
'* Uses Times New Roman 12, adds foundation from DSO
'* @param employees() — employee data array
'* @param empCount — number of employees
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Private Sub GenerateRiskWordDocument(ByRef employees() As EmployeeRiskData, ByVal empCount As Long)
    Dim wordApp As Object, doc As Object, rng As Object
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    Set doc = wordApp.Documents.Add
    
    ' Set default font Times New Roman for the whole document
    With doc.Styles(1).Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    
    ' Order header
    Set rng = doc.Range
    rng.Text = "ПРИКАЗ" & vbCrLf
    rng.Font.Bold = True
    rng.Font.Size = 12
    rng.Font.Name = "Times New Roman"
    rng.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    
    ' Employee numbering
    Dim i As Long, j As Long
    For i = 1 To empCount
        Set rng = doc.Range
        rng.Collapse Direction:=0
        
        ' Generate employee header
        Dim header As String
        header = i & ". " & mdlHelper.SklonitZvanie(employees(i).Rank) & " " & _
                 mdlHelper.SklonitFIO(employees(i).fio) & ", личный номер " & employees(i).lichniyNomer & ", " & _
                 mdlHelper.SklonitDolzhnost(employees(i).Position, employees(i).VoinskayaChast) & vbCrLf
        
        rng.Text = header
        rng.Font.Bold = False
        rng.Font.Size = 12
        rng.Font.Name = "Times New Roman"
        rng.ParagraphFormat.Alignment = 3 ' wdAlignParagraphJustify
        
        ' Periods
        For j = 1 To employees(i).periodCount
            Set rng = doc.Range
            rng.Collapse Direction:=0
            
            ' Use generated period string
            Dim periodText As String
            periodText = "    - " & employees(i).periods(j).PeriodString & _
                         " (" & employees(i).periods(j).daysCount & " дн.) = " & _
                         employees(i).periods(j).PercentValue & "%" & vbCrLf
            
            rng.Text = periodText
            
            rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
            rng.Font.Bold = False
            rng.Font.Size = 12
            rng.Font.Name = "Times New Roman"
            
            ' Expiration warning
            If employees(i).periods(j).IsExpired Then
                Set rng = doc.Range
                rng.Collapse Direction:=0
                rng.Text = "      ВНИМАНИЕ: Период с " & Format(employees(i).periods(j).StartDate, "dd.mm.yyyy") & _
                           " по " & Format(employees(i).periods(j).EndDate, "dd.mm.yyyy") & _
                           " превышает срок в три года на обращение за получением надбавки согласно ПМО 727" & vbCrLf
                rng.Font.Color = RGB(255, 0, 0)
                rng.Font.Bold = True
                rng.Font.Size = 12
                rng.Font.Name = "Times New Roman"
                rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
            End If
        Next j
        
        ' Add foundation from DSO
        Call AddRiskFoundationFromDSO(doc, employees(i).lichniyNomer)
        
        Set rng = doc.Range
        rng.Collapse Direction:=0
        rng.InsertAfter vbCrLf ' Empty line between employees
    Next i
    
    ' Save
    Dim savePath As String
    savePath = ThisWorkbook.Path & "\ПриказЗаРиск_" & Format(Date, "dd.mm.yyyy") & ".docx"
    Call mdlHelper.SaveWordDocumentSafe(doc, savePath)
End Sub


'/**
'* AddRiskFoundationFromDSO — add foundation string from column 4 of DSO sheet
'* @param doc — Word.Document object
'* @param lichniyNomer — employee personal number
'* @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
'*/
Private Sub AddRiskFoundationFromDSO(ByVal doc As Object, ByVal lichniyNomer As String)
    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    
    Dim lastRow As Long, i As Long
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    Dim foundationText As String
    foundationText = ""
    
    ' Search for foundations in column 4 for given personal number
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
    
    ' Add foundation to document
    If foundationText <> "" Then
        Dim rng As Object
        Set rng = doc.Range
        rng.Collapse Direction:=0
        
        Dim foundationLine As String
        foundationLine = "Основание: " & foundationText & vbCrLf
        
        rng.Text = foundationLine
        rng.Font.Size = 12
        rng.Font.Name = "Times New Roman"
        rng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
    End If
End Sub

