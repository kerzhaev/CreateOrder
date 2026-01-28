Attribute VB_Name = "mdlDataImport"
' === СНАПШОТ ВЕРСИИ === 12.07.2025 20:45 ===
' Рабочая версия сохранена: 12.07.2025 20:45

Option Explicit

' Модуль mdlDataImport для импорта данных на лист "Штат"
' Версия: 1.0.0
' Дата: 09.07.2025
' Описание: Загружает данные из внешнего Excel файла на лист "Штат" с полной заменой

Sub ImportDataToStaff()
    Dim sourceFile As String
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim targetWorksheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim selectedSheet As String
    
    On Error GoTo ErrorHandler
    
    ' Отключаем обновление экрана для ускорения
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Диалог выбора файла
    sourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Выберите файл для импорта данных на лист 'Штат'")
    
    ' Проверяем, выбрал ли пользователь файл
    If sourceFile = "False" Then
        MsgBox "Импорт отменен пользователем.", vbInformation, "Импорт данных"
        GoTo CleanUp
    End If
    
    ' Открываем исходный файл
    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=True)
    
    ' Если в файле несколько листов, даем пользователю выбрать
    If sourceWorkbook.Worksheets.count > 1 Then
        selectedSheet = SelectWorksheetFromFile(sourceWorkbook)
        If selectedSheet = "" Then
            sourceWorkbook.Close False
            MsgBox "Импорт отменен - лист не выбран.", vbInformation, "Импорт данных"
            GoTo CleanUp
        End If
        Set sourceWorksheet = sourceWorkbook.Worksheets(selectedSheet)
    Else
        Set sourceWorksheet = sourceWorkbook.Worksheets(1)
    End If
    
    ' Проверяем наличие листа "Штат" в текущей книге
    Set targetWorksheet = GetOrCreateStaffWorksheet()
    
    ' Очищаем лист "Штат"
    targetWorksheet.Cells.Clear
    
    ' Определяем диапазон данных в исходном файле
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.count, 1).End(xlUp).Row
    lastCol = sourceWorksheet.Cells(1, sourceWorksheet.Columns.count).End(xlToLeft).Column
    
    ' Проверяем, есть ли данные для копирования
    If lastRow = 1 And sourceWorksheet.Cells(1, 1).value = "" Then
        sourceWorkbook.Close False
        MsgBox "Выбранный лист пуст - нет данных для импорта.", vbExclamation, "Импорт данных"
        GoTo CleanUp
    End If
    
    ' Копируем данные
    sourceWorksheet.Range(sourceWorksheet.Cells(1, 1), sourceWorksheet.Cells(lastRow, lastCol)).Copy
    targetWorksheet.Cells(1, 1).PasteSpecial xlPasteAll
    
    ' Автоподбор ширины столбцов
    targetWorksheet.Columns.AutoFit
    
    ' Закрываем исходный файл
    sourceWorkbook.Close False
    
    ' Очищаем буфер обмена
    Application.CutCopyMode = False
    
    ' Сообщаем об успешном импорте
    MsgBox "Данные успешно импортированы на лист 'Штат'!" & vbCrLf & _
           "Импортировано строк: " & lastRow & vbCrLf & _
           "Импортировано столбцов: " & lastCol, vbInformation, "Импорт завершен"
    
    GoTo CleanUp
    
ErrorHandler:
    ' Обработка ошибок
    If Not sourceWorkbook Is Nothing Then
        sourceWorkbook.Close False
    End If
    MsgBox "Ошибка при импорте данных: " & Err.Description, vbCritical, "Ошибка импорта"
    
CleanUp:
    ' Восстанавливаем настройки
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    
    ' Очищаем переменные
    Set sourceWorkbook = Nothing
    Set sourceWorksheet = Nothing
    Set targetWorksheet = Nothing
    
    Call mdlHelper.InitStaffColumnIndexes
End Sub

' Функция для выбора листа из файла с несколькими листами
' Версия: 1.0.0
' Дата: 09.07.2025
' Описание: Показывает диалог выбора листа, если в файле несколько листов
Function SelectWorksheetFromFile(wb As Workbook) As String
    Dim sheetNames As String
    Dim selectedSheet As String
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Формируем список листов
    sheetNames = "Выберите лист для импорта:" & vbCrLf & vbCrLf
    i = 1
    For Each ws In wb.Worksheets
        sheetNames = sheetNames & i & ". " & ws.Name & vbCrLf
        i = i + 1
    Next ws
    
    ' Запрашиваем номер листа у пользователя
    Dim userInput As String
    userInput = InputBox(sheetNames & vbCrLf & "Введите номер листа (1-" & wb.Worksheets.count & "):", _
                        "Выбор листа для импорта", "1")
    
    ' Проверяем ввод пользователя
    If userInput = "" Then
        SelectWorksheetFromFile = ""
        Exit Function
    End If
    
    Dim sheetNumber As Integer
    If IsNumeric(userInput) Then
        sheetNumber = CInt(userInput)
        If sheetNumber >= 1 And sheetNumber <= wb.Worksheets.count Then
            SelectWorksheetFromFile = wb.Worksheets(sheetNumber).Name
        Else
            MsgBox "Неверный номер листа. Выберите число от 1 до " & wb.Worksheets.count, vbExclamation
            SelectWorksheetFromFile = ""
        End If
    Else
        MsgBox "Введите корректный номер листа.", vbExclamation
        SelectWorksheetFromFile = ""
    End If
End Function

' Функция для получения или создания листа "Штат"
' Версия: 1.0.0
' Дата: 09.07.2025
' Описание: Возвращает лист "Штат", создает его если не существует
' Функция для получения или создания листа "Штат"
Function GetOrCreateStaffWorksheet() As Worksheet
    Dim ws As Worksheet
    Dim staffExists As Boolean
    
    staffExists = False
    
    ' Проверяем существование листа "Штат"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Штат" Then
            Set GetOrCreateStaffWorksheet = ws
            staffExists = True
            Exit For
        End If
    Next ws
    
    ' Создаем лист "Штат", если он не существует
    If Not staffExists Then
        Set GetOrCreateStaffWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateStaffWorksheet.Name = "Штат"
        MsgBox "Лист 'Штат' не существовал и был создан.", vbInformation, "Создание листа"
    End If
End Function


' Функция для предварительного просмотра данных перед импортом
' Версия: 1.0.0
' Дата: 09.07.2025
' Описание: Показывает превью первых 5 строк данных для подтверждения импорта
Sub PreviewImportData()
    Dim sourceFile As String
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim previewText As String
    Dim i As Long, j As Long
    Dim maxRows As Long, maxCols As Long
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Диалог выбора файла
    sourceFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls), *.xlsx;*.xlsm;*.xls", _
        Title:="Выберите файл для предварительного просмотра")
    
    If sourceFile = "False" Then Exit Sub
    
    ' Открываем файл для просмотра
    Set sourceWorkbook = Workbooks.Open(sourceFile, ReadOnly:=True)
    Set sourceWorksheet = sourceWorkbook.Worksheets(1)
    
    ' Формируем превью (первые 5 строк и 5 столбцов)
    maxRows = IIf(sourceWorksheet.Cells(sourceWorksheet.Rows.count, 1).End(xlUp).Row > 5, 5, sourceWorksheet.Cells(sourceWorksheet.Rows.count, 1).End(xlUp).Row)
    maxCols = IIf(sourceWorksheet.Cells(1, sourceWorksheet.Columns.count).End(xlToLeft).Column > 5, 5, sourceWorksheet.Cells(1, sourceWorksheet.Columns.count).End(xlToLeft).Column)
    
    previewText = "=== ПРЕДВАРИТЕЛЬНЫЙ ПРОСМОТР ДАННЫХ ===" & vbCrLf & vbCrLf
    previewText = previewText & "Файл: " & sourceWorkbook.Name & vbCrLf
    previewText = previewText & "Лист: " & sourceWorksheet.Name & vbCrLf & vbCrLf
    
    For i = 1 To maxRows
        For j = 1 To maxCols
            previewText = previewText & sourceWorksheet.Cells(i, j).value & vbTab
        Next j
        previewText = previewText & vbCrLf
    Next i
    
    If maxRows = 5 Then previewText = previewText & "... (показаны первые 5 строк)"
    
    sourceWorkbook.Close False
    
    MsgBox previewText, vbInformation, "Предварительный просмотр"
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close False
    Application.ScreenUpdating = True
    MsgBox "Ошибка при просмотре файла: " & Err.Description, vbCritical
End Sub


