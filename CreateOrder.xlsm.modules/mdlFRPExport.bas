Attribute VB_Name = "mdlFRPExport"
' ===================================================================
' Module mdlFRPExport (Universal)
' Version: 3.1.0 (Alushta Update)
' Date: 15.02.2026
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Description: Генерация Excel отчетов:
' 1. ДСО (сутки отдыха)
' 2. ФРП Риск (надбавка 2%)
' Updates: Добавлен столбец "Табельный номер" (берется из цифрового столбца "Лицо" в Штате)
' ===================================================================
Option Explicit

'/**
'* ExportPeriodsToExcel_WithChoice — Точка входа для кнопки "Выгрузка Алушта/ФРП"
'*/
Public Sub ExportPeriodsToExcel_WithChoice()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim choice As VbMsgBoxResult
    choice = MsgBox("Выберите тип отчёта:" & vbCrLf & vbCrLf & _
                    "Да - Отчёт ДСО (Сутки отдыха)" & vbCrLf & _
                    "Нет - Отчёт ФРП Риск (Надбавка 2%)" & vbCrLf & _
                    "Отмена - Выход", _
                    vbYesNoCancel + vbQuestion, "Выбор типа отчёта")
    
    If choice = vbYes Then
        Call CreateExcelReportPeriodsByLichniyNomer ' Старый отчет ДСО (обновленный)
    ElseIf choice = vbNo Then
        Call CreateRiskExcelReport ' Отчет Риск (обновленный)
    End If
End Sub

'====================================================================
' ЧАСТЬ 1: Отчет ДСО
'====================================================================
Sub CreateExcelReportPeriodsByLichniyNomer()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim wbNew As Workbook, wsNew As Worksheet
    Dim wsMain As Worksheet, wsStaff As Worksheet
    Dim i As Long, j As Long, outputRow As Long, lastRowMain As Long
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim colTableNumber As Long ' Переменная для номера столбца с табельным номером
    
    Dim uniquePersons As Collection, personData As Collection, periodList As Collection
    Dim periodArr() As Variant, cutoffDate As Date, filePath As String
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Создание отчёта ДСО..."

    Set wsMain = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")

    If Not mdlHelper.FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast) Then
        MsgBox "Ошибка: Не удалось найти основные столбцы в листе 'Штат'!", vbCritical
        GoTo CleanUp
    End If
    
    ' Находим столбец с табельным номером (цифровое "Лицо")
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)
    
    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastRowMain = wsMain.Cells(wsMain.Rows.count, "C").End(xlUp).Row

    Set uniquePersons = New Collection
    For i = 2 To lastRowMain
        Dim currentLichniyNomer As String
        currentLichniyNomer = Trim(CStr(wsMain.Cells(i, 3).value))
        If currentLichniyNomer <> "" Then
            On Error Resume Next
            Set personData = uniquePersons(currentLichniyNomer)
            If Err.number <> 0 Then
                Set personData = New Collection
                personData.Add currentLichniyNomer, "lichniyNomer"
                personData.Add Trim(CStr(wsMain.Cells(i, 2).value)), "fio"
                
                ' Получаем табельный номер из штата
                Dim tableNum As String
                tableNum = ""
                If colTableNumber > 0 Then
                    Dim staffRow As Long
                    ' Ищем строку сотрудника в штате по личному номеру
                    staffRow = mdlHelper.FindStaffRow(wsStaff, currentLichniyNomer, colLichniyNomer)
                    If staffRow > 0 Then
                        tableNum = Trim(CStr(wsStaff.Cells(staffRow, colTableNumber).value))
                    End If
                End If
                personData.Add tableNum, "tableNumber"
                
                Set periodList = New Collection
                personData.Add periodList, "periods"
                uniquePersons.Add personData, currentLichniyNomer
            End If
            On Error GoTo 0
            mdlHelper.CollectAllPersonPeriods wsMain, i, personData("periods")
        End If
    Next i

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "Отчет ДСО"
    
    ' Заголовки (Обновленная структура)
    wsNew.Cells(1, 1).value = "№ п/п"
    wsNew.Cells(1, 2).value = "ФИО"
    wsNew.Cells(1, 3).value = "Личный номер"
    wsNew.Cells(1, 4).value = "Табельный номер" ' НОВЫЙ СТОЛБЕЦ
    wsNew.Cells(1, 5).value = "Начало периода"
    wsNew.Cells(1, 6).value = "Конец периода"
    wsNew.Cells(1, 7).value = "Длительность, сут."
    wsNew.Cells(1, 8).value = "Сутки отдыха"
    wsNew.Cells(1, 9).value = "Актуален"
    
    wsNew.Range("A1:I1").Font.Bold = True
    outputRow = 2

    Dim infoRow As Long
    For infoRow = 1 To uniquePersons.count
        Set personData = uniquePersons(infoRow)
        Set periodList = personData("periods")

        If periodList.count > 0 Then
            ReDim periodArr(1 To periodList.count, 1 To 3)
            For j = 1 To periodList.count
                periodArr(j, 1) = periodList(j)(1)
                periodArr(j, 2) = periodList(j)(2)
                periodArr(j, 3) = periodList(j)(3)
            Next j
            
            ' Сортировка
            Call SortArray(periodArr)

            ' Расчет суток отдыха
            Dim totalDays As Long, totalRestDays As Long, restDaysArr() As Long
            totalDays = 0
            For j = 1 To UBound(periodArr)
                totalDays = totalDays + periodArr(j, 3)
            Next j
            totalRestDays = Int(totalDays / 3) * 2

            ReDim restDaysArr(1 To UBound(periodArr))
            Dim restBase As Long, restExtra As Long
            If periodList.count > 0 Then
                restBase = totalRestDays \ periodList.count
                restExtra = totalRestDays Mod periodList.count
                For j = 1 To periodList.count
                    restDaysArr(j) = restBase
                    If restExtra > 0 Then
                        restDaysArr(j) = restDaysArr(j) + 1
                        restExtra = restExtra - 1
                    End If
                    If restDaysArr(j) = 0 And totalRestDays > 0 Then restDaysArr(j) = 1
                Next j
            End If

            ' Вывод данных
            For j = 1 To UBound(periodArr)
                wsNew.Cells(outputRow, 1).value = outputRow - 1
                wsNew.Cells(outputRow, 2).value = personData("fio")
                wsNew.Cells(outputRow, 3).value = personData("lichniyNomer")
                wsNew.Cells(outputRow, 4).value = personData("tableNumber") ' Вывод табельного номера
                wsNew.Cells(outputRow, 5).value = periodArr(j, 1)
                wsNew.Cells(outputRow, 6).value = periodArr(j, 2)
                wsNew.Cells(outputRow, 7).value = periodArr(j, 3)
                wsNew.Cells(outputRow, 8).value = restDaysArr(j)
                wsNew.Cells(outputRow, 9).value = IIf(periodArr(j, 2) >= cutoffDate, "Да", "Нет")
                outputRow = outputRow + 1
            Next j
        End If
    Next infoRow

    wsNew.Columns("A:I").AutoFit
    filePath = ThisWorkbook.Path & "\Выгрузка_Алушта_ДСО_" & Format(Date, "dd.mm.yyyy") & ".xlsx"
    Application.DisplayAlerts = False
    If dir(filePath) <> "" Then Kill filePath
    wbNew.SaveAs filePath
    MsgBox "Отчёт ДСО сохранён: " & filePath, vbInformation
    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'====================================================================
' ЧАСТЬ 2: Отчет ФРП Риск
'====================================================================
Sub CreateRiskExcelReport()

    Call mdlHelper.EnsureStaffColumnsInitialized

    Dim wbNew As Workbook, wsNew As Worksheet
    Dim wsDSO As Worksheet, wsStaff As Worksheet
    Dim lastRow As Long, i As Long, outputRow As Long
    Dim lichniyNomer As String, fio As String, tableNumber As String
    Dim rawPeriods() As mdlRiskExport.RiskPeriod
    Dim splitPeriods() As mdlRiskExport.RiskPeriod
    Dim periodCount As Long, k As Long
    
    Dim colLichniyNomer As Long, colTableNumber As Long
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Создание отчёта ФРП Риск..."
    
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    ' Инициализация колонок штата
    colLichniyNomer = mdlHelper.colLichniyNomer_Global
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)
    
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row

    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = "ФРП Риск"
    
    ' Заголовки (Обновленные)
    wsNew.Cells(1, 1).value = "№ п/п"
    wsNew.Cells(1, 2).value = "ФИО"
    wsNew.Cells(1, 3).value = "Личный номер"
    wsNew.Cells(1, 4).value = "Табельный номер" ' НОВЫЙ СТОЛБЕЦ
    wsNew.Cells(1, 5).value = "Начало периода"
    wsNew.Cells(1, 6).value = "Конец периода"
    wsNew.Cells(1, 7).value = "Дней"
    wsNew.Cells(1, 8).value = "Процент"
    wsNew.Cells(1, 9).value = "Актуален"
    
    With wsNew.Range("A1:I1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    outputRow = 2
    
    ' Перебор всех сотрудников в ДСО
    For i = 2 To lastRow
        lichniyNomer = Trim(CStr(wsDSO.Cells(i, 3).value))
        fio = Trim(CStr(wsDSO.Cells(i, 2).value))
        tableNumber = ""
        
        If lichniyNomer <> "" Then
            ' Получаем табельный номер из штата
            If colTableNumber > 0 Then
                Dim staffRow As Long
                staffRow = mdlHelper.FindStaffRow(wsStaff, lichniyNomer, colLichniyNomer)
                If staffRow > 0 Then
                    tableNumber = Trim(CStr(wsStaff.Cells(staffRow, colTableNumber).value))
                End If
            End If
            
            ' 1. Сбор периодов
            periodCount = CollectRawRiskPeriods_Local(wsDSO, i, rawPeriods)
            
            If periodCount > 0 Then
                ' 2. Разбивка по месяцам
                Dim splitCount As Long
                splitCount = SplitPeriodsByMonth_SeparateRows(rawPeriods, periodCount, splitPeriods)
                
                ' 3. Вывод в Excel
                For k = 1 To splitCount
                    wsNew.Cells(outputRow, 1).value = outputRow - 1
                    wsNew.Cells(outputRow, 2).value = fio
                    wsNew.Cells(outputRow, 3).value = lichniyNomer
                    wsNew.Cells(outputRow, 4).value = tableNumber ' Вывод табельного номера
                    wsNew.Cells(outputRow, 5).value = splitPeriods(k).StartDate
                    wsNew.Cells(outputRow, 6).value = splitPeriods(k).EndDate
                    wsNew.Cells(outputRow, 7).value = splitPeriods(k).daysCount
                    wsNew.Cells(outputRow, 8).value = splitPeriods(k).PercentValue & "%"
                    wsNew.Cells(outputRow, 9).value = IIf(splitPeriods(k).IsExpired, "Нет", "Да")
                    
                    If splitPeriods(k).IsExpired Then
                        wsNew.Range("A" & outputRow & ":I" & outputRow).Interior.Color = RGB(255, 200, 200)
                    End If
                    
                    outputRow = outputRow + 1
                Next k
            End If
        End If
    Next i
    
    wsNew.Columns("A:I").AutoFit
    Dim filePathRisk As String
    filePathRisk = ThisWorkbook.Path & "\Выгрузка_Алушта_Риск_" & Format(Date, "dd.mm.yyyy") & ".xlsx"
    wbNew.SaveAs filePathRisk
    
    MsgBox "Отчёт ФРП Риск (Алушта) сохранён: " & filePathRisk, vbInformation
    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка при создании отчёта Риск: " & Err.Description, vbCritical
    If Not wbNew Is Nothing Then wbNew.Close False
    Resume CleanUp
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

'====================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (Локальные для отчета)
'====================================================================

' Сбор сырых периодов из одной строки ДСО
Private Function CollectRawRiskPeriods_Local(ws As Worksheet, rowNum As Long, ByRef periods() As mdlRiskExport.RiskPeriod) As Long
    Dim lastCol As Long, j As Long, pCount As Long
    pCount = 0
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    
    ReDim periods(1 To 50) ' Резерв
    Dim expirationDate As Date
    expirationDate = DateAdd("m", -42, Date) ' 3 года 6 месяцев
    
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

' Разбивка периодов по месяцам
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
        If monthlyAccumulator.exists(key) Then
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
        
        If monthlyAccumulator.exists(key) Then
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

