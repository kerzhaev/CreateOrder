VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectEmployee 
   Caption         =   "UserForm1"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16335
   OleObjectBlob   =   "frmSelectEmployee.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'version 5#
'Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectEmployee
'   Caption = "Выбор сотрудника"
'   ClientHeight = 6000
'   ClientLeft = 120
'   ClientTop = 465
'   ClientWidth = 10000
'   OleObjectBlob   =   "frmSelectEmployee.frx":0000
'   StartUpPosition = 2    'CenterScreen
'End
'Attribute VB_Name = "frmSelectEmployee"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False

' =====================================================================
' Форма выбора сотрудника для заполнения листа "Выплаты_Без_Периодов"
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' =====================================================================

Option Explicit

' === Публичные переменные для передачи результата ===
Public selectedLichniyNomer As String     ' Выбранный личный номер
Public selectedFIO As String              ' Выбранное ФИО
Public isCancelled As Boolean             ' Флаг отмены

' Примечание: Форма требует создания элементов управления в дизайнере Excel:
' - txtSearch (TextBox) - поле поиска
' - lstResults (ListBox) - список результатов, ColumnCount = 4
' - lblStatus (Label) - статус (количество найденных)
' - lblFIO (Label) - ФИО выбранного сотрудника
' - lblZvanie (Label) - звание
' - lblDolzhnost (Label) - должность
' - btnSelect (CommandButton) - кнопка "Выбрать"
' - btnCancel (CommandButton) - кнопка "Отмена"

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Инициализация формы
' =============================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Call mdlHelper.EnsureStaffColumnsInitialized
    
    isCancelled = True
    selectedLichniyNomer = ""
    selectedFIO = ""
    
    ' Настройка ListBox
    lstResults.ColumnCount = 4
    lstResults.ColumnWidths = "80 pt;150 pt;100 pt;200 pt"
    lstResults.Clear
    
    lblStatus.Caption = ""
    lblFIO.Caption = "ФИО:"
    lblZvanie.Caption = "Звание:"
    lblDolzhnost.Caption = "Должность:"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при инициализации формы: " & Err.Description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Обработка изменения текста поиска
' =============================================
Private Sub txtSearch_Change()
    On Error GoTo ErrorHandler
    
    Call mdlHelper.EnsureStaffColumnsInitialized
    
    Dim wsStaff As Worksheet
    Dim lastRow As Long, i As Long
    Dim query As String
    Dim foundCount As Long
    Dim infoLine As String
    Dim colTableNumber As Long
    Dim testValue As Variant
    
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    
    If mdlHelper.colFIO_Global <= 0 Or mdlHelper.colFIO_Global > wsStaff.Columns.count Or _
       mdlHelper.colLichniyNomer_Global <= 0 Or mdlHelper.colLichniyNomer_Global > wsStaff.Columns.count Then
        MsgBox "Ошибка: Столбцы листа не инициализированы. Проверьте структуру листа 'Штат' и перезапустите!", vbCritical
        Exit Sub
    End If
    
    query = LCase(Trim(txtSearch.Text))
    lstResults.Clear
    
    If Len(query) < 2 Then
        lstResults.AddItem "Введите минимум 2 символа для поиска..."
        lblStatus.Caption = ""
        Exit Sub
    End If
    
    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colFIO_Global).End(xlUp).Row
    foundCount = 0
    
    ' Получаем номер колонки с табельными номерами (если есть)
    colTableNumber = mdlHelper.FindTableNumberColumn(wsStaff)
    
    For i = 2 To lastRow
        Dim matchFound As Boolean
        matchFound = False
        
        ' Поиск по ФИО (частичное совпадение)
        If InStr(LCase(wsStaff.Cells(i, mdlHelper.colFIO_Global).value), query) > 0 Then
            matchFound = True
        End If
        
        ' Поиск по личному номеру (частичное совпадение)
        If Not matchFound Then
            If InStr(LCase(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value), query) > 0 Then
                matchFound = True
            End If
        End If
        
        ' Поиск по табельному номеру (точное совпадение, если колонка найдена)
        If Not matchFound And colTableNumber > 0 Then
            testValue = wsStaff.Cells(i, colTableNumber).value
            If Not IsEmpty(testValue) And IsNumeric(testValue) Then
                If Trim(CStr(testValue)) = Trim(query) Then
                    matchFound = True
                End If
            End If
        End If
        
        If matchFound Then
            infoLine = wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value & vbTab & _
                      wsStaff.Cells(i, mdlHelper.colFIO_Global).value & vbTab & _
                      wsStaff.Cells(i, mdlHelper.colZvanie_Global).value & vbTab & _
                      wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value
            lstResults.AddItem infoLine
            foundCount = foundCount + 1
        End If
    Next i
    
    If foundCount = 0 Then
        lstResults.AddItem "Ничего не найдено."
        lblStatus.Caption = ""
    Else
        lblStatus.Caption = "Найдено: " & foundCount
    End If
    
    Exit Sub
    
ErrorHandler:
    lstResults.AddItem "Ошибка поиска: " & Err.Description
    lblStatus.Caption = ""
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Обработка клика по списку результатов
' =============================================
Private Sub lstResults_Click()
    On Error GoTo ErrorHandler
    
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then
        Exit Sub
    End If
    
    Dim fields As Variant
    Dim lichniyNomer As String
    Dim wsStaff As Worksheet
    Dim lastRow As Long, i As Long
    
    fields = Split(lstResults.List(lstResults.ListIndex), vbTab)
    lichniyNomer = fields(0)
    
    Call mdlHelper.EnsureStaffColumnsInitialized
    Set wsStaff = ThisWorkbook.Sheets("Штат")
    lastRow = wsStaff.Cells(wsStaff.Rows.count, mdlHelper.colLichniyNomer_Global).End(xlUp).Row
    
    ' Находим данные сотрудника
    For i = 2 To lastRow
        If Trim(wsStaff.Cells(i, mdlHelper.colLichniyNomer_Global).value) = Trim(lichniyNomer) Then
            lblFIO.Caption = "ФИО: " & wsStaff.Cells(i, mdlHelper.colFIO_Global).value
            lblZvanie.Caption = "Звание: " & wsStaff.Cells(i, mdlHelper.colZvanie_Global).value
            lblDolzhnost.Caption = "Должность: " & wsStaff.Cells(i, mdlHelper.colDolzhnost_Global).value
            Exit For
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    ' Игнорируем ошибки при клике
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Обработка двойного клика по списку (выбор сотрудника)
' =============================================
Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then
        Exit Sub
    End If
    
    Call btnSelect_Click
    
    Exit Sub
    
ErrorHandler:
    ' Игнорируем ошибки
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Обработка кнопки "Выбрать"
' =============================================
Private Sub btnSelect_Click()
    On Error GoTo ErrorHandler
    
    If lstResults.ListCount = 0 Or lstResults.ListIndex = -1 Then
        MsgBox "Выберите сотрудника из списка.", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    Dim fields As Variant
    fields = Split(lstResults.List(lstResults.ListIndex), vbTab)
    
    selectedLichniyNomer = fields(0)
    selectedFIO = fields(1)
    isCancelled = False
    
    Me.Hide
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при выборе сотрудника: " & Err.Description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Обработка кнопки "Отмена"
' =============================================
Private Sub btnCancel_Click()
    On Error GoTo ErrorHandler
    
    isCancelled = True
    selectedLichniyNomer = ""
    selectedFIO = ""
    
    Me.Hide
    
    Exit Sub
    
ErrorHandler:
    ' Игнорируем ошибки
End Sub


