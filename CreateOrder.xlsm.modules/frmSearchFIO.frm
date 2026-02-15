VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchFIO 
   Caption         =   "UserForm1"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15240
   OleObjectBlob   =   "frmSearchFIO.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSearchFIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






' =====================================================================
' Автоматизированная форма поиска и внесения периодов для военнослужащего
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' =====================================================================

Option Explicit

' === Объявление модульных переменных для хранения состояния ===
Public selectedLichniyNomer As String     ' Хранит личный номер выбранного ФИО
Dim periodsArray As Variant            ' Хранит массив периодов выбранного ФИО (для упрощения работы)





' ================== Поиск по листу штата =====================
' === Обработчик поиска военнослужащего ===
'/**
'* Реализация поиска военнослужащего на форме с использованием глобальных индексов.
'* Теперь поиск и отображение работают независимо от порядка столбцов.
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub txtSearch_Change()
    EnsureStaffColumnsInitialized
    Dim wsStaff As Worksheet, lastRowFIO As Long, lastRowLN As Long, lastRow As Long, i As Long, query As String, foundCount As Long
    Set wsStaff = ThisWorkbook.Sheets("Штат")

    If colFIO_Global <= 0 Or colFIO_Global > wsStaff.Columns.count Or _
       colLichniyNomer_Global <= 0 Or colLichniyNomer_Global > wsStaff.Columns.count Then
        MsgBox "Ошибка: Индексы столбцов данных определены некорректно. Проверьте структуру листа 'Штат' и повторите инициализацию!", vbCritical
        Exit Sub
    End If

    lastRowFIO = wsStaff.Cells(wsStaff.Rows.count, colFIO_Global).End(xlUp).Row
    lastRowLN = wsStaff.Cells(wsStaff.Rows.count, colLichniyNomer_Global).End(xlUp).Row
    lastRow = Application.WorksheetFunction.Max(lastRowFIO, lastRowLN)

    query = LCase(Trim(txtSearch.text))
    lstResults.Clear
    If Len(query) < 2 Then
        lstResults.AddItem "Введите минимум 2 символа для поиска..."
        lblStatus.Caption = ""
        Exit Sub
    End If

    foundCount = 0
    For i = 2 To lastRow
        If InStr(LCase(wsStaff.Cells(i, colFIO_Global).value), query) > 0 Or _
           InStr(LCase(wsStaff.Cells(i, colLichniyNomer_Global).value), query) > 0 Then
            Dim infoLine As String
            infoLine = wsStaff.Cells(i, colLichniyNomer_Global).value & vbTab & _
                       wsStaff.Cells(i, colFIO_Global).value & vbTab & _
                       wsStaff.Cells(i, colZvanie_Global).value & vbTab & _
                       wsStaff.Cells(i, colDolzhnost_Global).value & vbTab & _
                       wsStaff.Cells(i, colVoinskayaChast_Global).value
            lstResults.AddItem infoLine
            foundCount = foundCount + 1
        End If
    Next i
    If foundCount = 0 Then
        lstResults.AddItem "Совпадений не найдено."
        lblStatus.Caption = ""
    Else
        lblStatus.Caption = "Найдено: " & foundCount
    End If
End Sub




Private Sub UserForm_Initialize()
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
        InitStaffColumnIndexes
    End If
End Sub


Public Sub FillByLichniyNomer()
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
    End If
End Sub

' ============ Реакция на выбор ФИО из результатов поиска =============
' === Обработка клика по результату поиска (выбор ФИО) ===
'/**
'* Обработка выбора строки в списке результатов поиска военнослужащих.
'* Использует глобальные индексы: корректно работает при любом порядке столбцов на листе "Штат".
'* При выборе — осуществляет переход на строку в "ДСО", выводит паспортные данные и периоды.
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub lstResults_Click()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub
    Dim fields As Variant, lichniyNomer As String, wsDSO As Worksheet, i As Long, lastRowDSO As Long
    fields = Split(lstResults.List(lstResults.ListIndex), vbTab)
    lichniyNomer = fields(0)

    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, 3).value) = Trim(lichniyNomer) Then
            Application.GoTo wsDSO.Cells(i, 3), True
            Exit For
        End If
    Next i

    selectedLichniyNomer = lichniyNomer
    ShowPassportData selectedLichniyNomer
    LoadPeriodsForLichniy selectedLichniyNomer
    lblStatus.Caption = "Готово. Можно добавлять или редактировать периоды."
End Sub




' ================ Вывод "паспорта" военнослужащего ================
' === Вывод "паспорта" военнослужащего на форму ===
'/**
'* Процедура вывода "паспорта" военнослужащего на форму frmSearchFIO.
'* Использует только глобальные индексы столбцов из листа "Штат" — независимо от порядка.
'* @param lichniyNomer String — личный номер военнослужащего
'* @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'*/
Private Sub ShowPassportData(lichniyNomer As String)

    EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet, lastRow As Long, i As Long
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    lastRow = wsStaff.Cells(wsStaff.Rows.count, colLichniyNomer_Global).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(wsStaff.Cells(i, colLichniyNomer_Global).value) = Trim(lichniyNomer) Then
            lblFIO.Caption = wsStaff.Cells(i, colFIO_Global).value
            lblZvanie.Caption = "Звание: " & wsStaff.Cells(i, colZvanie_Global).value
            lblDolzhnost.Caption = "Должность: " & wsStaff.Cells(i, colDolzhnost_Global).value
            lblChast.Caption = "Часть: " & Trim(wsStaff.Cells(i, colVoinskayaChast_Global).value)

            Exit Sub
        End If
    Next i
    lblFIO.Caption = "ФИО:"
    lblZvanie.Caption = "Звание:"
    lblDolzhnost.Caption = "Должность:"
    lblChast.Caption = "Часть:"
    
    Debug.Print "selectedLichniyNomer: " & selectedLichniyNomer

End Sub


' ============== Загрузка периодов из "ДСО" по личному номеру ============
' === Вывод таблицы периодов для выбранного ФИО ===
' --- Показывает красиво периоды с заголовком для выбранного военнослужащего ---
' ===== КРАСИВЫЙ ВЫВОД ПЕРИОДОВ В 4 КОЛОНКИ (ListBox с ColumnCount=4) =====
Private Sub LoadPeriodsForLichniy(lichniyNomer As String)
    Dim wsDSO As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long, periodCounter As Integer
    Dim baseReasonRaw As String, baseReasonArr() As String
    Dim rowIdx As Long
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    lstPeriods.Clear
    lstPeriods.ColumnCount = 4
    lstPeriods.ColumnWidths = "35 pt;75 pt;75 pt;200 pt"
    
    periodCounter = 0

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(lichniyNomer) Then
            ' Получаем массив оснований, разделённых запятыми
            baseReasonRaw = wsDSO.Cells(i, 4).value
            baseReasonArr = Split(baseReasonRaw, ",")
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            j = 5
            Do While j + 1 <= lastCol
                If wsDSO.Cells(i, j).value <> "" And wsDSO.Cells(i, j + 1).value <> "" Then
                    periodCounter = periodCounter + 1
                    rowIdx = lstPeriods.ListCount
                    lstPeriods.AddItem periodCounter
                    lstPeriods.List(rowIdx, 1) = Format(wsDSO.Cells(i, j).value, "dd.mm.yyyy")
                    lstPeriods.List(rowIdx, 2) = Format(wsDSO.Cells(i, j + 1).value, "dd.mm.yyyy")
                    If UBound(baseReasonArr) >= periodCounter - 1 Then
                        lstPeriods.List(rowIdx, 3) = Trim(baseReasonArr(periodCounter - 1))
                    Else
                        lstPeriods.List(rowIdx, 3) = ""
                    End If
                End If
                j = j + 2
            Loop
            Exit For
        End If
    Next i
    If lstPeriods.ListCount = 0 Then
        lstPeriods.AddItem "Нет действующих периодов"
    End If
End Sub



' ============ Автозаполнение ComboBox оснований ===============
' --- Заполнение ComboBox "Основание" только приказами из столбца 4 (через запятую) ---
Private Sub cmbReason_Enter()
    Dim wsDSO As Worksheet
    Dim lastRow As Long, i As Long
    Dim reasonsDict As Object
    Dim arrBase As Variant, part As Variant
    Set reasonsDict = CreateObject("Scripting.Dictionary")
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 4).End(xlUp).Row

    cmbReason.Clear

    ' Проходим по строкам и собираем все приказы из столбца 4
    For i = 2 To lastRow
        If wsDSO.Cells(i, 4).value <> "" Then
            arrBase = Split(wsDSO.Cells(i, 4).value, ",")
            For Each part In arrBase
                part = Trim(part)
                If part <> "" Then
                    If Not reasonsDict.exists(part) Then
                        reasonsDict.Add part, 1
                    End If
                End If
            Next part
        End If
    Next i

    ' Заносим в ComboBox только уникальные значения
    Dim key As Variant
    For Each key In reasonsDict.keys
        cmbReason.AddItem key
    Next key
End Sub



' ===== Добавить новый период во вкладку ДСО для выбранного ФИО =====
'============================================================
' Добавление периода: если ФИО не существует в ДСО,
' автоматически добавляет строку, присваивает порядковый номер,
' переносит ФИО, личный номер, а далее — периоды и основания.
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'============================================================
Private Sub btnAddPeriod_Click()
    ' Проверка заполненности полей периода
    If txtPeriodStart.text = "" Or txtPeriodEnd.text = "" Or cmbReason.text = "" Then
        MsgBox "Укажите даты и основание для периода.", vbExclamation
        Exit Sub
    End If
    If selectedLichniyNomer = "" Then
        MsgBox "ФИО не выбрана!", vbExclamation
        Exit Sub
    End If

    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    Dim lastRow As Long, rowNum As Long, found As Boolean, i As Long

    ' Поиск ФИО по личному номеру
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    found = False
    rowNum = 0
    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            found = True
            rowNum = i
            Exit For
        End If
    Next i

    ' Если ФИО нет — создаём новую запись
    If Not found Then
        rowNum = lastRow + 1
        wsDSO.Cells(rowNum, 1).value = rowNum - 1 ' порядковый номер (столбец 1)
        wsDSO.Cells(rowNum, 2).value = lblFIO.Caption ' ФИО (столбец 2)
        wsDSO.Cells(rowNum, 3).value = selectedLichniyNomer ' Личный номер (столбец 3)
        wsDSO.Cells(rowNum, 4).value = "" ' столбец основание — заполнится ниже
    End If

    ' Ищем последний занятый столбец в строке и добавляем период далее
    Dim lastCol As Long
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    If lastCol < 4 Then lastCol = 4 ' если только что создана строка

    ' Даты периода попадают в новые столбцы (пары: начало, конец)
    wsDSO.Cells(rowNum, lastCol + 1).value = txtPeriodStart.text
    wsDSO.Cells(rowNum, lastCol + 2).value = txtPeriodEnd.text

    ' --- Корректно добавлять основание только в столбец 4 (через запятую) ---
    Dim oldReasonRaw As String, newReasonRaw As String
    oldReasonRaw = wsDSO.Cells(rowNum, 4).value
    If Trim(oldReasonRaw) = "" Then
        newReasonRaw = cmbReason.text
    Else
        If Right(Trim(oldReasonRaw), 1) <> "," Then
            oldReasonRaw = oldReasonRaw & ","
        End If
        newReasonRaw = oldReasonRaw & " " & cmbReason.text
    End If
    wsDSO.Cells(rowNum, 4).value = newReasonRaw

    Call LoadPeriodsForLichniy(selectedLichniyNomer)
    txtPeriodStart.text = ""
    txtPeriodEnd.text = ""
    cmbReason.text = ""
    lblStatus.Caption = "Период добавлен."
End Sub



'============================================================
' Удаляет выбранный период (дату начала/конца и основание)
' Данные после удаления сдвигаются влево, последние ячейки
' очищаются, чтобы не возникало "дыр" и ошибок в паре дат.
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
'============================================================
Private Sub btnDeletePeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim periodNum As Integer, reasonsArr() As String, maxPairs As Integer, k As Integer
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    periodNum = lstPeriods.List(lstPeriods.ListIndex, 0)
    If IsNumeric(periodNum) = False Or periodNum < 1 Then Exit Sub

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            ' Считаем максимальное число пар дат (для этой строки)
            maxPairs = (lastCol - 5 + 1) \ 2
            ' Сдвигаем все правее стоящие пары дат влево
            For k = periodNum To maxPairs - 1
                wsDSO.Cells(i, 5 + (k - 1) * 2).value = wsDSO.Cells(i, 5 + k * 2).value
                wsDSO.Cells(i, 5 + (k - 1) * 2 + 1).value = wsDSO.Cells(i, 5 + k * 2 + 1).value
            Next k
            ' Очищаем самую последнюю пару ячеек (теперь лишняя)
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2).value = vbNullString
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2 + 1).value = vbNullString

            ' ---- Корректируем основания (массив) ----
            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            If UBound(reasonsArr) >= periodNum - 1 Then
                For k = periodNum - 1 To UBound(reasonsArr) - 1
                    reasonsArr(k) = reasonsArr(k + 1)
                Next k
                ReDim Preserve reasonsArr(UBound(reasonsArr) - 1)
                wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")
            End If

            ' ---- Дополнительно: Очищаем все "мусорные" ячейки по датам ----
            For k = 5 To lastCol
                If Trim(wsDSO.Cells(i, k).value) = "" Or IsEmpty(wsDSO.Cells(i, k).value) Then
                    wsDSO.Cells(i, k).value = vbNullString
                End If
            Next k
            
            Call CleanPeriodsForLichniyRow(wsDSO, i)

            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            txtPeriodStart.text = ""
            txtPeriodEnd.text = ""
            cmbReason.text = ""
            lblStatus.Caption = "Период удалён корректно."
            Exit Sub
        End If
    Next i
End Sub






' ============ Закрыть форму ============
Private Sub btnClose_Click()
    Unload Me
End Sub


' --- При выборе периода заполняются поля ввода дат и основания ---
Private Sub lstPeriods_Click()
    If lstPeriods.ListIndex = -1 Then Exit Sub
    ' Проверка на заголовок или спец. строку
    If lstPeriods.List(lstPeriods.ListIndex, 0) = "Нет действующих периодов" Then Exit Sub
    
    txtPeriodStart.text = lstPeriods.List(lstPeriods.ListIndex, 1) ' Начало
    txtPeriodEnd.text = lstPeriods.List(lstPeriods.ListIndex, 2)   ' Конец
    cmbReason.text = lstPeriods.List(lstPeriods.ListIndex, 3)      ' Основание
End Sub


' --- Редактирование выбранного периода ---
' --- Обновление выбранного периода, основания инлайн ---
Private Sub btnEditPeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub
    Dim wsDSO As Worksheet, lastRow As Long, i As Long, lastCol As Long, j As Long, periodNum As Integer
    Dim reasonsArr() As String, k As Integer, maxPairs As Integer
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    periodNum = lstPeriods.List(lstPeriods.ListIndex, 0)
    If IsNumeric(periodNum) = False Then Exit Sub

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            maxPairs = (lastCol - 4 + 1) \ 2
            ' Индексирование строго по совпадающей паре
            j = 5 + (periodNum - 1) * 2
            wsDSO.Cells(i, j).value = txtPeriodStart.text
            wsDSO.Cells(i, j + 1).value = txtPeriodEnd.text

            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            ' Если оснований меньше, расширить массив до нужного размера
            If UBound(reasonsArr) < periodNum - 1 Then
                ReDim Preserve reasonsArr(periodNum - 1)
                For k = 0 To UBound(reasonsArr)
                    If Trim(reasonsArr(k)) = "" Then
                        reasonsArr(k) = "-"
                    End If
                Next k
            End If
            reasonsArr(periodNum - 1) = cmbReason.text
            wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")

            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            lblStatus.Caption = "Период успешно обновлён."
            Exit Sub
        End If
    Next i
End Sub


Private Sub CleanPeriodsForLichniyRow(wsDSO As Worksheet, rowNum As Long)
    ' Очищает весь "хвост" после последней заполненной пары дат для корректной работы валидатора
    Dim lastCol As Long, lastPeriodCol As Long, j As Long, isEmptyPair As Boolean
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    lastPeriodCol = 4 ' Последний "реально заполненный" столбец в паре

    ' Определяем реальный хвост (последний заполненный период)
    For j = 5 To lastCol Step 2
        If wsDSO.Cells(rowNum, j).value <> "" And wsDSO.Cells(rowNum, j + 1).value <> "" Then
            lastPeriodCol = j + 1
        End If
    Next j

    ' Очищаем всё дальше этого столбца (если вдруг что-то осталось в хвосте)
    For j = lastPeriodCol + 1 To lastCol
        wsDSO.Cells(rowNum, j).value = vbNullString
    Next j
End Sub

Private Sub btnSelect_Click()
    Call AddSearchResultToDSO_LastRow
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call AddSearchResultToDSO_LastRow
End Sub

' === ДОБАВЛЕНИЕ ФИО и личного номера после последней строки ДСО ===
'===============================================================
' Добавляет ФИО и личный номер после последней записи,
' используя имена столбцов из строки заголовка таблицы "Штат".
'===============================================================
'===============================================================
' Добавление новой записи ФИО (без дубликатов) в "ДСО" и перенос атрибутов
Private Sub AddSearchResultToDSO_LastRow()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")

    ' Получаем личный номер и ФИО из списка (или другие ключи)
    Dim fields As Variant, fioVal As String, lnVal As String
    fields = Split(lstResults.List(lstResults.ListIndex), vbTab)
    fioVal = fields(1)
    lnVal = fields(0)

    ' Проверяем на дубли в ДСО
    Dim exists As Boolean, lastRowDSO As Long, i As Long
    exists = False
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, 3).value) = Trim(lnVal) Or Trim(wsDSO.Cells(i, 2).value) = Trim(fioVal) Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        MsgBox "ФИО или личный номер уже есть на листе ДСО! Добавление невозможно.", vbInformation
        Exit Sub
    End If

    ' Получаем все данные сотрудника централизованно
    Dim staffInfo As Object
    Set staffInfo = GetStaffData(lnVal, True) ' по личному номеру

    lastRowDSO = lastRowDSO + 1
    wsDSO.Cells(lastRowDSO, 1).value = lastRowDSO - 1
    wsDSO.Cells(lastRowDSO, 2).value = staffInfo("Лицо")
    wsDSO.Cells(lastRowDSO, 3).value = staffInfo("Личный номер")
'    wsDSO.Cells(lastRowDSO, 5).Value = staffInfo("Воинское звание")
'    wsDSO.Cells(lastRowDSO, 6).Value = staffInfo("Должность")
'    wsDSO.Cells(lastRowDSO, 7).Value = staffInfo("Часть")
    ' ...дополнительно любые другие поля из staffInfo

    MsgBox "Данные успешно добавлены в новую строку " & lastRowDSO & " листа ДСО.", vbInformation

    wsDSO.Activate
    wsDSO.Cells(lastRowDSO, 2).Select
    Unload Me
End Sub



