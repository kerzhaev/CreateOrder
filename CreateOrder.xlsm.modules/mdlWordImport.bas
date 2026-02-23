Attribute VB_Name = "mdlWordImport"
' ===============================================================================
' Module: mdlWordImport
' Version: 1.6.0 (InputBox Reason Update + Memory Leak Fix)
' Date: 23.02.2026
' Author: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' Description: Полный цикл ETL: извлечение рапортов из Word, конвертация в HTML,
'              чтение в Dictionary, выгрузка в ДСО с сортировкой и группировкой.
'              Включает интерактивный запрос основания (номера приказа) для импортируемых периодов.
' ===============================================================================

Option Explicit

Private Const wdFormatFilteredHTML As Long = 10

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description [T008] Главная процедура импорта (точка входа)
' =============================================
Public Sub ExecuteWordImport()
    On Error GoTo ErrorHandler
    
    Dim wordFilePath As String
    Dim htmlFilePath As String
    Dim parsedData As Object ' Dictionary
    Dim finalReport As String
    Dim baseReasonVal As Variant
    Dim baseReason As String
    
    ' 1. Выбор файла
    wordFilePath = SelectWordFile()
    If wordFilePath = "" Then
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' 2. Запрос основания для импортируемых периодов
    baseReasonVal = Application.InputBox("Введите основание (номер приказа/распоряжения) для импортируемых периодов:" & vbCrLf & _
                                         "Если основание не требуется, оставьте поле пустым и нажмите ОК." & vbCrLf & _
                                         "Для полной отмены импорта нажмите 'Отмена'.", _
                                         "Основание (Колонка D)", "", Type:=2)
    
    ' Если пользователь нажал "Отмена"
    If VarType(baseReasonVal) = vbBoolean And baseReasonVal = False Then
        MsgBox "Импорт отменен пользователем.", vbInformation, "Отмена"
        Application.StatusBar = False
        Exit Sub
    End If
    
    baseReason = Trim(CStr(baseReasonVal))
    
    Application.StatusBar = "Конвертация Word документа... Пожалуйста, подождите."
    
    ' 3. [T004] Конвертация в HTML
    htmlFilePath = ConvertWordToTempHTML(wordFilePath)
    If htmlFilePath = "" Then Err.Raise vbObjectError + 1, , "Не удалось конвертировать файл Word."
    
    Application.StatusBar = "Сбор данных из таблицы..."
    
    ' 4. [T005] Чтение в Dictionary
    Set parsedData = ParseHTMLToDict(htmlFilePath)
    
    ' Удаляем временный HTML-файл
    On Error Resume Next
    Kill htmlFilePath
    On Error GoTo ErrorHandler
    
    If parsedData.count = 0 Then
        MsgBox "Не удалось найти корректные данные (личные номера) в таблице. Проверьте формат рапорта.", vbExclamation, "Результат импорта"
        Application.StatusBar = False
        Exit Sub
    End If
    
    Application.StatusBar = "Запись данных в лист ДСО..."
    
    ' 5. [T006, T007] Запись в лист ДСО и сортировка
    finalReport = ApplyDictToDSOSheet(parsedData, baseReason)
    
    ' 6. Финализация
    Application.StatusBar = False
    
    finalReport = finalReport & vbCrLf & vbCrLf & _
                  "ВНИМАНИЕ: Для визуальной проверки новых периодов на ошибки " & _
                  "нажмите кнопку 'Проверить данные' на ленте (вкладка Валидация)."
                  
    MsgBox "Импорт завершен успешно!" & vbCrLf & vbCrLf & finalReport, vbInformation, "Итоги импорта рапорта"
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.EnableEvents = True ' Возвращаем события при ошибке
    MsgBox "Критическая ошибка в процессе импорта: " & Err.Description, vbCritical, "Ошибка импорта"
End Sub

' =============================================
' @description Диалоговое окно выбора файла Word
' =============================================
Private Function SelectWordFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Выберите утвержденный рапорт (Word)"
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc; *.docx"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectWordFile = .SelectedItems(1)
        Else
            SelectWordFile = ""
        End If
    End With
End Function

' =============================================
' [T004] Экстракция: Открытие Word и сохранение в HTML
' =============================================
Private Function ConvertWordToTempHTML(ByVal filePath As String) As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim tempPath As String
    Dim wordWasNotRunning As Boolean
    
    tempPath = Environ("TEMP") & "\raport_temp_" & Format(Now, "hhmmss") & ".htm"
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        wordWasNotRunning = True
    Else
        wordWasNotRunning = False
    End If
    On Error GoTo ErrorHandler
    
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileName:=filePath, ReadOnly:=True, Visible:=False)
    
    wdDoc.SaveAs2 fileName:=tempPath, fileFormat:=wdFormatFilteredHTML
    wdDoc.Close False
    
    If wordWasNotRunning Then wdApp.Quit False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    
    ConvertWordToTempHTML = tempPath
    Exit Function
    
ErrorHandler:
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If wordWasNotRunning And Not wdApp Is Nothing Then wdApp.Quit False
    Set wdDoc = Nothing
    Set wdApp = Nothing
    ConvertWordToTempHTML = ""
End Function

' =============================================
' [T005] Трансформация: Чтение HTML таблицы Excel-ем и упаковка в Dictionary
' =============================================
Private Function ParseHTMLToDict(ByVal htmlPath As String) As Object
    Dim wbTemp As Workbook
    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long, i As Long
    Dim colLichniy As Long, colStart As Long, colEnd As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler ' ВАЖНО: Добавили перехват ошибок
    
    Set wbTemp = Workbooks.Open(fileName:=htmlPath, ReadOnly:=True)
    Set ws = wbTemp.Sheets(1)
    
    colLichniy = FindColBySubstring(ws, "личный")
    colStart = FindColBySubstring(ws, "начал")
    colEnd = FindColBySubstring(ws, "окончан")
    
    If colLichniy = 0 Then colLichniy = 4
    If colStart = 0 Then colStart = 5
    If colEnd = 0 Then colEnd = 6
    
    lastRow = ws.Cells(ws.Rows.count, colLichniy).End(xlUp).Row
    
    Dim lnVal As String, strStart As String, strEnd As String
    Dim dStart As Date, dEnd As Date
    Dim periodDict As Object
    Dim personPeriods As Collection
    
    For i = 1 To lastRow
        lnVal = Trim(CStr(ws.Cells(i, colLichniy).value))
        strStart = Trim(CStr(ws.Cells(i, colStart).value))
        strEnd = Trim(CStr(ws.Cells(i, colEnd).value))
        
        If lnVal <> "" And InStr(1, LCase(lnVal), "личный") = 0 Then
            dStart = mdlHelper.ParseDateSafe(strStart)
            dEnd = mdlHelper.ParseDateSafe(strEnd)
            
            Set periodDict = CreateObject("Scripting.Dictionary")
            periodDict.Add "StartText", strStart
            periodDict.Add "EndText", strEnd
            periodDict.Add "StartDate", dStart
            periodDict.Add "EndDate", dEnd
            
            If Not dict.exists(lnVal) Then
                Set personPeriods = New Collection
                dict.Add lnVal, personPeriods
            End If
            
            dict(lnVal).Add periodDict
        End If
    Next i
    
    wbTemp.Close False
    Set wbTemp = Nothing
    Application.DisplayAlerts = True
    
    Set ParseHTMLToDict = dict
    Exit Function
    
ErrorHandler:
    ' ГАРАНТИРОВАННО ЗАКРЫВАЕМ ФАЙЛ ПРИ СБОЕ
    If Not wbTemp Is Nothing Then
        On Error Resume Next
        wbTemp.Close False
        Set wbTemp = Nothing
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    Set ParseHTMLToDict = dict
End Function

Private Function FindColBySubstring(ws As Worksheet, subStr As String) As Long
    Dim i As Long, j As Long
    Dim cellText As String
    For i = 1 To 5
        For j = 1 To 10
            cellText = LCase(Trim(CStr(ws.Cells(i, j).value)))
            If InStr(1, cellText, LCase(subStr)) > 0 Then
                FindColBySubstring = j
                Exit Function
            End If
        Next j
    Next i
    FindColBySubstring = 0
End Function

' =============================================
' [T006] Загрузка (Load): Запись данных в лист ДСО с учетом Основания
' =============================================
Private Function ApplyDictToDSOSheet(dict As Object, ByVal baseReason As String) As String
    Dim wsDSO As Worksheet
    Dim lnKey As Variant
    Dim i As Long, lastRowDSO As Long, rowNum As Long
    Dim newEmpCount As Long, updEmpCount As Long, addedPeriodsCount As Long
    Dim pCol As Long
    Dim personPeriods As Collection
    Dim period As Object
    Dim staffData As Object
    Dim currentReason As String
    
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    If lastRowDSO < 2 Then lastRowDSO = 1
    
    newEmpCount = 0
    updEmpCount = 0
    addedPeriodsCount = 0
    
    ' Отключаем события, чтобы Worksheet_Change (раскраска) не тормозила вставку
    Application.EnableEvents = False
    
    For Each lnKey In dict.keys()
        rowNum = 0
        
        ' Ищем сотрудника в ДСО
        For i = 2 To lastRowDSO
            If Trim(CStr(wsDSO.Cells(i, 3).value)) = Trim(CStr(lnKey)) Then
                rowNum = i
                Exit For
            End If
        Next i
        
        ' Если не найден - создаем новую строку
        If rowNum = 0 Then
            lastRowDSO = lastRowDSO + 1
            rowNum = lastRowDSO
            
            wsDSO.Cells(rowNum, 1).value = rowNum - 1
            wsDSO.Cells(rowNum, 3).value = lnKey
            
            ' Пытаемся вытянуть ФИО из Штата
            Set staffData = mdlHelper.GetStaffData(CStr(lnKey), True)
            If staffData.count > 0 Then
                wsDSO.Cells(rowNum, 2).value = staffData("Лицо")
            Else
                wsDSO.Cells(rowNum, 2).value = "НОВЫЙ: ФИО не найдено"
            End If
            
            newEmpCount = newEmpCount + 1
        Else
            updEmpCount = updEmpCount + 1
        End If
        
        ' Логика добавления текста Основания (если пользователь его ввел)
        If baseReason <> "" Then
            currentReason = Trim(CStr(wsDSO.Cells(rowNum, 4).value))
            
            If currentReason = "" Then
                wsDSO.Cells(rowNum, 4).value = baseReason
            Else
                ' Проверяем, нет ли уже такого текста в ячейке, чтобы избежать дублирования
                If InStr(1, currentReason, baseReason, vbTextCompare) = 0 Then
                    If Right(currentReason, 1) <> "," And Right(currentReason, 1) <> ";" Then
                        currentReason = currentReason & ","
                    End If
                    wsDSO.Cells(rowNum, 4).value = currentReason & " " & baseReason
                End If
            End If
        End If
        
        Set personPeriods = dict(lnKey)
        
        ' Находим первую пустую пару колонок для записи (начиная с 5-й)
        pCol = 5
        Do While Trim(CStr(wsDSO.Cells(rowNum, pCol).value)) <> "" Or Trim(CStr(wsDSO.Cells(rowNum, pCol + 1).value)) <> ""
            pCol = pCol + 2
        Loop
        
        ' Вписываем периоды "как есть" (пользовательский текст)
        Dim pItem As Variant
        For Each pItem In personPeriods
            Set period = pItem
            wsDSO.Cells(rowNum, pCol).value = period("StartText")
            wsDSO.Cells(rowNum, pCol + 1).value = period("EndText")
            pCol = pCol + 2
            addedPeriodsCount = addedPeriodsCount + 1
        Next pItem
        
        ' [T007] Сортируем строку хронологически
        Call SortPeriodsInRow(wsDSO, rowNum)
        
    Next lnKey
    
    ' Включаем события обратно
    Application.EnableEvents = True
    
    ApplyDictToDSOSheet = "Добавлено периодов: " & addedPeriodsCount & vbCrLf & _
                          "Обновлено сотрудников: " & updEmpCount & vbCrLf & _
                          "Добавлено новых строк (сотрудников): " & newEmpCount
End Function

' =============================================
' [T007] Хронологическая сортировка периодов в строке
' =============================================
Private Sub SortPeriodsInRow(ws As Worksheet, rowNum As Long)
    Dim lastCol As Long
    Dim j As Long, pCount As Long
    Dim periods() As Variant
    
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 6 Then Exit Sub
    
    ReDim periods(1 To (lastCol - 4) / 2 + 1, 1 To 3)
    pCount = 0
    
    ' Читаем периоды в массив
    For j = 5 To lastCol Step 2
        If Trim(CStr(ws.Cells(rowNum, j).value)) <> "" Or Trim(CStr(ws.Cells(rowNum, j + 1).value)) <> "" Then
            pCount = pCount + 1
            periods(pCount, 1) = ws.Cells(rowNum, j).value
            periods(pCount, 2) = ws.Cells(rowNum, j + 1).value
            ' Парсим дату для сортировки. Если дата кривая, она получит значение 0 и улетит в начало
            periods(pCount, 3) = mdlHelper.ParseDateSafe(periods(pCount, 1))
        End If
    Next j
    
    If pCount <= 1 Then Exit Sub
    
    ' Сортировка Пузырьком (Bubble sort) по возрастанию даты начала
    Dim i As Long, k As Long
    Dim t1 As Variant, t2 As Variant, t3 As Date
    
    For i = 1 To pCount - 1
        For k = i + 1 To pCount
            If periods(i, 3) > periods(k, 3) Then
                t1 = periods(i, 1): t2 = periods(i, 2): t3 = periods(i, 3)
                periods(i, 1) = periods(k, 1): periods(i, 2) = periods(k, 2): periods(i, 3) = periods(k, 3)
                periods(k, 1) = t1: periods(k, 2) = t2: periods(k, 3) = t3
            End If
        Next k
    Next i
    
    ' Перезаписываем отсортированные данные обратно
    Dim colIdx As Long
    colIdx = 5
    For i = 1 To pCount
        ws.Cells(rowNum, colIdx).value = periods(i, 1)
        ws.Cells(rowNum, colIdx + 1).value = periods(i, 2)
        colIdx = colIdx + 2
    Next i
    
    ' Зачищаем хвосты (если были пустые "дырки" между периодами, они сместились)
    If colIdx <= lastCol Then
        ws.Range(ws.Cells(rowNum, colIdx), ws.Cells(rowNum, lastCol)).ClearContents
        ws.Range(ws.Cells(rowNum, colIdx), ws.Cells(rowNum, lastCol)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(rowNum, colIdx), ws.Cells(rowNum, lastCol)).ClearComments
    End If
End Sub

