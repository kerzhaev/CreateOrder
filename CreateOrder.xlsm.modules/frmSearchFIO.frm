VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchFIO 
   Caption         =   "UserForm1"
   ClientHeight    =   8925.001
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
' Обновление: v.1.5.1 (Multi-column Search & Keyboard Navigation)
' =====================================================================

Option Explicit

' === Объявление модульных переменных ===
Public selectedLichniyNomer As String     ' Хранит личный номер выбранного ФИО

' =====================================================================
' ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКА
' =====================================================================
Private Sub UserForm_Initialize()
    ' Настройка списка результатов (5 колонок)
    ' 1:Л.Номер, 2:ФИО, 3:Звание, 4:Должность, 5:Часть
    With lstResults
        .ColumnCount = 5
        .ColumnWidths = "35 pt;120 pt;70 pt;100 pt;30 pt"
    End With

    ' Если форма открыта с уже переданным номером (например, из другой части программы)
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
        ' Убедимся, что индексы столбцов инициализированы
        mdlHelper.EnsureStaffColumnsInitialized
    End If
End Sub

Public Sub FillByLichniyNomer()
    If selectedLichniyNomer <> "" Then
        Call ShowPassportData(selectedLichniyNomer)
        Call LoadPeriodsForLichniy(selectedLichniyNomer)
    End If
End Sub

' =====================================================================
' ПОИСК (LIVE SEARCH) И НАВИГАЦИЯ
' =====================================================================

'/**
'* Живой поиск с заполнением 5 колонок ListBox.
'*/
Private Sub txtSearch_Change()
    mdlHelper.EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet
    Dim lastRow As Long, i As Long, foundCount As Long
    Dim query As String
    Dim fLichniy As String, fFIO As String, fZvanie As String, fDolzhnost As String, fChast As String
    
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    ' Проверка корректности индексов столбцов (глобальные переменные из mdlHelper)
    If mdlHelper.colFIO_Global <= 0 Then Exit Sub

    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colFIO_Global).End(xlUp).Row
    
    query = LCase(Trim(txtSearch.Text))
    lstResults.Clear
    
    If Len(query) < 2 Then
        lblStatus.Caption = "Введите минимум 2 символа..."
        Exit Sub
    End If

    foundCount = 0
    Application.ScreenUpdating = False ' Отключаем перерисовку для скорости
    
    For i = 2 To lastRow
        fLichniy = CStr(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value)
        fFIO = CStr(wsStaff.Cells(i, mdlHelper.colFIO_Global).value)
        
        ' Поиск по ФИО или Личному номеру
        If InStr(LCase(fFIO), query) > 0 Or InStr(LCase(fLichniy), query) > 0 Then
            fZvanie = CStr(wsStaff.Cells(i, mdlHelper.colZvanie_Global).value)
            fDolzhnost = CStr(wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value)
            fChast = CStr(wsStaff.Cells(i, mdlHelper.colVoinskayaChast_Global).value)
            
            ' Заполнение колонок
            lstResults.AddItem fLichniy             ' Col 0 (Hidden ID/Key)
            lstResults.List(foundCount, 1) = fFIO        ' Col 1
            lstResults.List(foundCount, 2) = fZvanie     ' Col 2
            lstResults.List(foundCount, 3) = fDolzhnost  ' Col 3
            lstResults.List(foundCount, 4) = fChast      ' Col 4
            
            foundCount = foundCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If foundCount = 0 Then
        lblStatus.Caption = "Совпадений не найдено."
    Else
        lblStatus.Caption = "Найдено: " & foundCount
    End If
End Sub

' Обработка клика по результату поиска
Private Sub lstResults_Click()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub
    
    Dim lichniyNomer As String
    Dim wsDSO As Worksheet, i As Long, lastRowDSO As Long
    
    ' Получаем личный номер из 0-й колонки
    lichniyNomer = lstResults.List(lstResults.ListIndex, 0)

    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    ' Пытаемся найти сотрудника в ДСО и перейти к нему
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, 3).value) = Trim(lichniyNomer) Then
            Application.GoTo wsDSO.Cells(i, 3), True
            Exit For
        End If
    Next i

    selectedLichniyNomer = lichniyNomer
    ShowPassportData selectedLichniyNomer
    LoadPeriodsForLichniy selectedLichniyNomer
    lblStatus.Caption = "Выбран: " & lstResults.List(lstResults.ListIndex, 1)
End Sub

' Клавиатура: Поле поиска
Private Sub txtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Стрелка ВНИЗ - переход в список
    If KeyCode = vbKeyDown Then
        If lstResults.ListCount > 0 Then
            lstResults.SetFocus
            If lstResults.ListIndex = -1 Then lstResults.ListIndex = 0
        End If
        KeyCode = 0
    End If
    
    ' ENTER - Если 1 результат, выбираем его сразу
    If KeyCode = vbKeyReturn Then
        If lstResults.ListCount = 1 Then
            lstResults.ListIndex = 0
            Call lstResults_Click
            txtPeriodStart.SetFocus ' Сразу переходим к вводу даты
        ElseIf lstResults.ListCount > 1 Then
            lstResults.SetFocus
            lstResults.ListIndex = 0
        End If
        KeyCode = 0
    End If
End Sub

' Клавиатура: Список результатов
Private Sub lstResults_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' ENTER - Подтверждение выбора
    If KeyCode = vbKeyReturn Then
        If lstResults.ListIndex > -1 Then
            Call lstResults_Click
            txtPeriodStart.SetFocus
        End If
        KeyCode = 0
    End If
End Sub

' Двойной клик - Добавить в ДСО (если нет) или выбрать
Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call AddSearchResultToDSO_LastRow
End Sub

Private Sub btnSelect_Click()
    Call AddSearchResultToDSO_LastRow
End Sub

' =====================================================================
' ЛОГИКА ДОБАВЛЕНИЯ В ДСО (Add New Logic)
' =====================================================================
Private Sub AddSearchResultToDSO_LastRow()
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")

    Dim fioVal As String, lnVal As String
    ' Читаем данные из колонок ListBox
    lnVal = lstResults.List(lstResults.ListIndex, 0)
    fioVal = lstResults.List(lstResults.ListIndex, 1)

    ' Проверка на дубликаты
    Dim exists As Boolean, lastRowDSO As Long, i As Long
    exists = False
    lastRowDSO = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    For i = 2 To lastRowDSO
        If Trim(wsDSO.Cells(i, 3).value) = Trim(lnVal) Then
            exists = True
            Exit For
        End If
    Next i

    If exists Then
        MsgBox "Сотрудник с таким личным номером уже есть в ДСО (строка " & i & ").", vbInformation
        ' Просто переходим к нему
        selectedLichniyNomer = lnVal
        Call lstResults_Click
        Exit Sub
    End If

    ' Добавляем новую строку
    lastRowDSO = lastRowDSO + 1
    wsDSO.Cells(lastRowDSO, 1).value = lastRowDSO - 1      ' № п/п
    wsDSO.Cells(lastRowDSO, 2).value = fioVal              ' ФИО
    wsDSO.Cells(lastRowDSO, 3).value = lnVal               ' Личный номер
    wsDSO.Cells(lastRowDSO, 4).value = ""                  ' Основания (пусто)

    MsgBox "Сотрудник добавлен в строку " & lastRowDSO & ".", vbInformation

    ' Активируем выбор
    selectedLichniyNomer = lnVal
    Call lstResults_Click
End Sub

' =====================================================================
' ОТОБРАЖЕНИЕ ДАННЫХ ("ПАСПОРТ")
' =====================================================================
Private Sub ShowPassportData(lichniyNomer As String)
    mdlHelper.EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet, lastRow As Long, i As Long
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colLichniyNomer_Global).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value) = Trim(lichniyNomer) Then
            lblFIO.Caption = wsStaff.Cells(i, mdlHelper.colFIO_Global).value
            lblZvanie.Caption = "Звание: " & wsStaff.Cells(i, mdlHelper.colZvanie_Global).value
            lblDolzhnost.Caption = "Должность: " & wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value
            lblChast.Caption = "Часть: " & Trim(wsStaff.Cells(i, mdlHelper.colVoinskayaChast_Global).value)
            Exit Sub
        End If
    Next i
    
    ' Если не нашли (такое бывает, если удалили из Штата, но остался в ДСО)
    lblFIO.Caption = "ФИО: Не найдено в штате"
    lblZvanie.Caption = "Звание: -"
    lblDolzhnost.Caption = "Должность: -"
    lblChast.Caption = "Часть: -"
End Sub

' =====================================================================
' РАБОТА С ПЕРИОДАМИ
' =====================================================================
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
            ' Получаем массив оснований
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

' Выбор периода в списке (заполнение полей)
Private Sub lstPeriods_Click()
    If lstPeriods.ListIndex = -1 Then Exit Sub
    If lstPeriods.List(lstPeriods.ListIndex, 0) = "Нет действующих периодов" Then Exit Sub
    
    txtPeriodStart.Text = lstPeriods.List(lstPeriods.ListIndex, 1) ' Начало
    txtPeriodEnd.Text = lstPeriods.List(lstPeriods.ListIndex, 2)   ' Конец
    cmbReason.Text = lstPeriods.List(lstPeriods.ListIndex, 3)      ' Основание
End Sub

' Автозаполнение ComboBox оснований
Private Sub cmbReason_Enter()
    Dim wsDSO As Worksheet
    Dim lastRow As Long, i As Long
    Dim reasonsDict As Object
    Dim arrBase As Variant, part As Variant
    
    Set reasonsDict = CreateObject("Scripting.Dictionary")
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 4).End(xlUp).Row

    cmbReason.Clear

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

    Dim key As Variant
    For Each key In reasonsDict.keys
        cmbReason.AddItem key
    Next key
End Sub

' =====================================================================
' КНОПКИ УПРАВЛЕНИЯ ПЕРИОДАМИ (ADD, EDIT, DELETE)
' =====================================================================

' Добавить новый период
Private Sub btnAddPeriod_Click()
    If txtPeriodStart.Text = "" Or txtPeriodEnd.Text = "" Or cmbReason.Text = "" Then
        MsgBox "Укажите даты и основание для периода.", vbExclamation
        Exit Sub
    End If
    If selectedLichniyNomer = "" Then
        MsgBox "Сотрудник не выбран!", vbExclamation
        Exit Sub
    End If

    Dim wsDSO As Worksheet
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    Dim lastRow As Long, rowNum As Long, found As Boolean, i As Long

    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    found = False
    
    ' Ищем строку сотрудника
    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            found = True
            rowNum = i
            Exit For
        End If
    Next i

    If Not found Then
        MsgBox "Ошибка: Сотрудник не найден в ДСО. Сначала добавьте его через поиск.", vbCritical
        Exit Sub
    End If

    ' Ищем куда вставить
    Dim lastCol As Long
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    If lastCol < 4 Then lastCol = 4

    ' Записываем даты
    wsDSO.Cells(rowNum, lastCol + 1).value = txtPeriodStart.Text
    wsDSO.Cells(rowNum, lastCol + 2).value = txtPeriodEnd.Text

    ' Добавляем основание через запятую
    Dim oldReasonRaw As String, newReasonRaw As String
    oldReasonRaw = CStr(wsDSO.Cells(rowNum, 4).value)
    
    If Trim(oldReasonRaw) = "" Then
        newReasonRaw = cmbReason.Text
    Else
        If Right(Trim(oldReasonRaw), 1) <> "," Then oldReasonRaw = oldReasonRaw & ","
        newReasonRaw = oldReasonRaw & " " & cmbReason.Text
    End If
    wsDSO.Cells(rowNum, 4).value = newReasonRaw

    Call LoadPeriodsForLichniy(selectedLichniyNomer)
    txtPeriodStart.Text = ""
    txtPeriodEnd.Text = ""
    cmbReason.Text = ""
    lblStatus.Caption = "Период добавлен."
End Sub

' Редактировать выбранный период
Private Sub btnEditPeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub
    
    Dim wsDSO As Worksheet, lastRow As Long, i As Long
    Dim periodNum As Integer, reasonsArr() As String
    
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    ' Получаем номер периода из первого столбца ListBox
    If Not IsNumeric(lstPeriods.List(lstPeriods.ListIndex, 0)) Then Exit Sub
    periodNum = CInt(lstPeriods.List(lstPeriods.ListIndex, 0))

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            ' Обновляем даты
            Dim colIndex As Long
            colIndex = 5 + (periodNum - 1) * 2
            wsDSO.Cells(i, colIndex).value = txtPeriodStart.Text
            wsDSO.Cells(i, colIndex + 1).value = txtPeriodEnd.Text

            ' Обновляем основание в массиве
            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            
            ' Если массив меньше чем нужно (на всякий случай)
            If UBound(reasonsArr) < periodNum - 1 Then
                ReDim Preserve reasonsArr(periodNum - 1)
            End If
            
            reasonsArr(periodNum - 1) = cmbReason.Text
            wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")

            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            lblStatus.Caption = "Период обновлён."
            Exit Sub
        End If
    Next i
End Sub

' Удалить выбранный период
Private Sub btnDeletePeriod_Click()
    If selectedLichniyNomer = "" Or lstPeriods.ListIndex = -1 Then Exit Sub

    Dim wsDSO As Worksheet
    Dim lastRow As Long, lastCol As Long, i As Long
    Dim periodNum As Integer, reasonsArr() As String, maxPairs As Integer, k As Integer
    
    Set wsDSO = ThisWorkbook.Sheets("ДСО")
    lastRow = wsDSO.Cells(wsDSO.Rows.count, 3).End(xlUp).Row
    
    If Not IsNumeric(lstPeriods.List(lstPeriods.ListIndex, 0)) Then Exit Sub
    periodNum = CInt(lstPeriods.List(lstPeriods.ListIndex, 0))

    For i = 2 To lastRow
        If Trim(wsDSO.Cells(i, 3).value) = Trim(selectedLichniyNomer) Then
            lastCol = wsDSO.Cells(i, wsDSO.Columns.count).End(xlToLeft).Column
            maxPairs = (lastCol - 5 + 1) \ 2
            
            ' Сдвигаем даты влево
            For k = periodNum To maxPairs - 1
                wsDSO.Cells(i, 5 + (k - 1) * 2).value = wsDSO.Cells(i, 5 + k * 2).value
                wsDSO.Cells(i, 5 + (k - 1) * 2 + 1).value = wsDSO.Cells(i, 5 + k * 2 + 1).value
            Next k
            
            ' Очищаем последние
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2).value = vbNullString
            wsDSO.Cells(i, 5 + (maxPairs - 1) * 2 + 1).value = vbNullString

            ' Удаляем основание из массива
            reasonsArr = Split(wsDSO.Cells(i, 4).value, ",")
            If UBound(reasonsArr) >= periodNum - 1 Then
                For k = periodNum - 1 To UBound(reasonsArr) - 1
                    reasonsArr(k) = reasonsArr(k + 1)
                Next k
                
                If UBound(reasonsArr) > 0 Then
                    ReDim Preserve reasonsArr(UBound(reasonsArr) - 1)
                    wsDSO.Cells(i, 4).value = Join(reasonsArr, ",")
                Else
                    wsDSO.Cells(i, 4).value = "" ' Если это был последний
                End If
            End If

            Call CleanPeriodsForLichniyRow(wsDSO, i)
            Call LoadPeriodsForLichniy(selectedLichniyNomer)
            
            txtPeriodStart.Text = ""
            txtPeriodEnd.Text = ""
            cmbReason.Text = ""
            lblStatus.Caption = "Период удалён."
            Exit Sub
        End If
    Next i
End Sub

' Очистка хвостов (мусора) в строке
Private Sub CleanPeriodsForLichniyRow(wsDSO As Worksheet, rowNum As Long)
    Dim lastCol As Long, lastPeriodCol As Long, j As Long
    lastCol = wsDSO.Cells(rowNum, wsDSO.Columns.count).End(xlToLeft).Column
    lastPeriodCol = 4

    For j = 5 To lastCol Step 2
        If wsDSO.Cells(rowNum, j).value <> "" And wsDSO.Cells(rowNum, j + 1).value <> "" Then
            lastPeriodCol = j + 1
        End If
    Next j

    For j = lastPeriodCol + 1 To lastCol
        wsDSO.Cells(rowNum, j).value = vbNullString
    Next j
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
