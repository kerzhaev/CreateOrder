Attribute VB_Name = "mdlRibbonHandlers"
' ===============================================================================
' Модуль mdlRibbonHandlers для обработки событий пользовательской ленты
' Версия: 2.1.0
' Дата: 12.07.2025
' Автор: Система управления военным персоналом
' Описание: Полный набор обработчиков событий для кнопок пользовательской ленты
' Функциональность:
' - Обработка всех кнопок экспорта документов
' - Управление данными (импорт, предпросмотр)
' - Валидация и диагностика данных
' - Системные функции (справка, настройки, проверка готовности)
' Изменения в v2.1.0: Добавлены недостающие обработчики для валидации и управления данными
' ===============================================================================

Option Explicit

' Обработчик для основного приказа
Sub RunMainExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call ExportToWordFromStaffByLichniyNomer
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании основного приказа: " & Err.description, vbCritical, "Ошибка"
End Sub

' Обработчик для справки ДСО
Sub RunSpravkaExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call ExportToWordSpravkaFromTemplate
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании справки: " & Err.description, vbCritical, "Ошибка"
End Sub

' Обработчик для рапорта
Sub RunRaportExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call ExportToWordRaportFromTemplateByLichniyNomer
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании рапорта: " & Err.description, vbCritical, "Ошибка"
End Sub

' Обработчик для Excel отчета
Sub RunExcelReport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call CreateExcelReportPeriodsByLichniyNomer
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании Excel отчета: " & Err.description, vbCritical, "Ошибка"
End Sub

' *** НЕДОСТАЮЩИЙ ОБРАБОТЧИК *** Обработчик для валидации данных
Sub RunDataValidation(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Выполняется валидация данных..."
    Call ValidateMainSheetData
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Ошибка при валидации данных: " & Err.description, vbCritical, "Ошибка валидации"
End Sub

' *** НЕДОСТАЮЩИЙ ОБРАБОТЧИК *** Обработчик для диагностики структуры
Sub RunDiagnoseStructure(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call DiagnoseWorkbookStructure
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при диагностике структуры: " & Err.description, vbCritical, "Ошибка диагностики"
End Sub

' *** НЕДОСТАЮЩИЙ ОБРАБОТЧИК *** Обработчик для импорта данных
Sub RunImportData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call ImportDataToStaff
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при импорте данных: " & Err.description, vbCritical, "Ошибка импорта"
End Sub

' *** НЕДОСТАЮЩИЙ ОБРАБОТЧИК *** Обработчик для предпросмотра данных
Sub RunPreviewData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call PreviewImportData
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при предпросмотре данных: " & Err.description, vbCritical, "Ошибка предпросмотра"
End Sub

' Обработчик для справки
Sub ShowHelp(control As IRibbonControl)
    Dim helpText As String
    helpText = "=== МАКРОСЫ ДЛЯ РАБОТЫ С ДАННЫМИ СВО ===" & vbCrLf & vbCrLf
    helpText = helpText & "[ЭКСПОРТ] ЭКСПОРТ ДОКУМЕНТОВ:" & vbCrLf
    helpText = helpText & "• Основной приказ - создает Word документ с приказом в дательном падеже" & vbCrLf
    helpText = helpText & "• Справка ДСО - создает справки на основе шаблона Word" & vbCrLf
    helpText = helpText & "• Рапорт - создает рапорты о выплате компенсации" & vbCrLf & vbCrLf
    helpText = helpText & "[ОТЧЕТЫ] ОТЧЕТЫ:" & vbCrLf
    helpText = helpText & "• Отчет по периодам - создает сводный Excel отчет" & vbCrLf & vbCrLf
    helpText = helpText & "[ДАННЫЕ] УПРАВЛЕНИЕ ДАННЫМИ:" & vbCrLf
    helpText = helpText & "• Импорт данных - загружает данные из Excel в лист 'Штат'" & vbCrLf
    helpText = helpText & "• Предпросмотр - показывает предварительный просмотр файла" & vbCrLf & vbCrLf
    helpText = helpText & "[ВАЛИДАЦИЯ] ПРОВЕРКА ДАННЫХ:" & vbCrLf
    helpText = helpText & "• Проверить данные - выполняет полную валидацию листа ДСО" & vbCrLf
    helpText = helpText & "• Диагностика структуры - анализирует структуру листов" & vbCrLf & vbCrLf
    helpText = helpText & "[ТРЕБОВАНИЯ] ТРЕБОВАНИЯ:" & vbCrLf
    helpText = helpText & "• Шаблоны Word должны находиться в папке с Excel файлом" & vbCrLf
    helpText = helpText & "• Лист 'Штат' должен содержать данные о сотрудниках" & vbCrLf
    helpText = helpText & "• Основной лист должен содержать периоды службы" & vbCrLf
    helpText = helpText & "• Столбец 'Личный номер' обязателен для уникальной идентификации" & vbCrLf & vbCrLf
    helpText = helpText & "[ШАБЛОНЫ] ФАЙЛЫ ШАБЛОНОВ:" & vbCrLf
    helpText = helpText & "• Шаблон_Справка.docx" & vbCrLf
    helpText = helpText & "• Шаблон_Рапорт.docx"
    
    MsgBox helpText, vbInformation, "Справка по макросам СВО"
End Sub

' Обработчик для настроек (обновленная версия)
Sub ShowSettings(control As IRibbonControl)
    Dim settingsText As String
    settingsText = "=== НАСТРОЙКИ МАКРОСОВ ===" & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПАПКА] Текущая папка: " & ThisWorkbook.Path & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПРОВЕРКА] Проверка шаблонов:" & vbCrLf
    
    ' Проверяем наличие шаблонов
    If dir(ThisWorkbook.Path & "\Шаблон_Справка.docx") <> "" Then
        settingsText = settingsText & "[+] Шаблон_Справка.docx - найден" & vbCrLf
    Else
        settingsText = settingsText & "[-] Шаблон_Справка.docx - НЕ НАЙДЕН" & vbCrLf
    End If
    
    If dir(ThisWorkbook.Path & "\Шаблон_Рапорт.docx") <> "" Then
        settingsText = settingsText & "[+] Шаблон_Рапорт.docx - найден" & vbCrLf
    Else
        settingsText = settingsText & "[-] Шаблон_Рапорт.docx - НЕ НАЙДЕН" & vbCrLf
    End If
    
    settingsText = settingsText & vbCrLf & "[ЛИСТЫ] Проверка листов:" & vbCrLf
    
    ' Проверяем наличие листа "ДСО"
    Dim dsoExists As Boolean
    dsoExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "ДСО" Then
            dsoExists = True
            Exit For
        End If
    Next ws
    
    If dsoExists Then
        settingsText = settingsText & "[+] Лист 'ДСО' - найден" & vbCrLf
    Else
        settingsText = settingsText & "[-] Лист 'ДСО' - НЕ НАЙДЕН" & vbCrLf
    End If
    
    ' Проверяем наличие листа "Штат"
    Dim wsExists As Boolean
    wsExists = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Штат" Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    If wsExists Then
        settingsText = settingsText & "[+] Лист 'Штат' - найден" & vbCrLf
    Else
        settingsText = settingsText & "[-] Лист 'Штат' - НЕ НАЙДЕН" & vbCrLf
    End If
    
    settingsText = settingsText & vbCrLf & "[СТАТИСТИКА] Информация о данных:" & vbCrLf
    
    ' Подсчитываем количество записей в листе ДСО
    If dsoExists Then
        Dim dsoSheet As Worksheet
        Set dsoSheet = ThisWorkbook.Sheets("ДСО")
        Dim lastRowDSO As Long
        lastRowDSO = dsoSheet.Cells(dsoSheet.Rows.count, "C").End(xlUp).Row
        
        If lastRowDSO > 1 Then
            settingsText = settingsText & "[ДАННЫЕ] Записей в листе ДСО: " & (lastRowDSO - 1) & vbCrLf
        Else
            settingsText = settingsText & "[ДАННЫЕ] Лист ДСО пуст" & vbCrLf
        End If
    End If
    
    ' Подсчитываем количество записей в листе "Штат"
    If wsExists Then
        Dim staffSheet As Worksheet
        Set staffSheet = ThisWorkbook.Sheets("Штат")
        Dim lastRowStaff As Long
        lastRowStaff = staffSheet.Cells(staffSheet.Rows.count, "A").End(xlUp).Row
        
        If lastRowStaff > 1 Then
            settingsText = settingsText & "[ШТАТ] Записей в листе 'Штат': " & (lastRowStaff - 1) & vbCrLf
        Else
            settingsText = settingsText & "[ШТАТ] Лист 'Штат' пуст" & vbCrLf
        End If
    End If
    
    settingsText = settingsText & vbCrLf & "[ВЕРСИЯ] Версия макросов: 2.1.0" & vbCrLf
    settingsText = settingsText & "[ДАТА] Дата обновления: 12.07.2025" & vbCrLf
    settingsText = settingsText & "[НОВОЕ] Поддержка личных номеров: ДА"
    
    MsgBox settingsText, vbInformation, "Настройки и проверка"
End Sub

' Функция для проверки готовности системы
Sub CheckSystemReadiness(control As IRibbonControl)
    Dim readinessText As String
    Dim isReady As Boolean
    isReady = True
    
    readinessText = "=== ПРОВЕРКА ГОТОВНОСТИ СИСТЕМЫ ===" & vbCrLf & vbCrLf
    
    ' Проверка шаблонов
    readinessText = readinessText & "[ШАБЛОНЫ]" & vbCrLf
    If dir(ThisWorkbook.Path & "\Шаблон_Справка.docx") <> "" Then
        readinessText = readinessText & "[OK] Шаблон справки найден" & vbCrLf
    Else
        readinessText = readinessText & "[ОШИБКА] Шаблон справки отсутствует" & vbCrLf
        isReady = False
    End If
    
    If dir(ThisWorkbook.Path & "\Шаблон_Рапорт.docx") <> "" Then
        readinessText = readinessText & "[OK] Шаблон рапорта найден" & vbCrLf
    Else
        readinessText = readinessText & "[ОШИБКА] Шаблон рапорта отсутствует" & vbCrLf
        isReady = False
    End If
    
    ' Проверка листов
    readinessText = readinessText & vbCrLf & "[СТРУКТУРА ДАННЫХ]" & vbCrLf
    Dim wsExists As Boolean
    wsExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Штат" Then
            wsExists = True
            Exit For
        End If
    Next ws
    
    If wsExists Then
        readinessText = readinessText & "[OK] Лист 'Штат' найден" & vbCrLf
    Else
        readinessText = readinessText & "[ОШИБКА] Лист 'Штат' отсутствует" & vbCrLf
        isReady = False
    End If
    
    ' Проверка наличия данных
    Dim mainSheet As Worksheet
    Set mainSheet = ThisWorkbook.Sheets("ДСО")
    Dim lastRowMain As Long
    lastRowMain = mainSheet.Cells(mainSheet.Rows.count, "C").End(xlUp).Row
    
    If lastRowMain > 1 Then
        readinessText = readinessText & "[OK] Данные в основном листе найдены" & vbCrLf
    Else
        readinessText = readinessText & "[ПРЕДУПРЕЖДЕНИЕ] Основной лист пуст" & vbCrLf
    End If
    
    ' Проверка структуры листа ДСО
    readinessText = readinessText & vbCrLf & "[СТРУКТУРА ЛИСТА ДСО]" & vbCrLf
    If mainSheet.Cells(1, 2).value = "ФИО" And mainSheet.Cells(1, 3).value = "Личный номер" Then
        readinessText = readinessText & "[OK] Структура листа ДСО корректна" & vbCrLf
    Else
        readinessText = readinessText & "[ПРЕДУПРЕЖДЕНИЕ] Проверьте структуру листа ДСО (B=ФИО, C=Личный номер)" & vbCrLf
    End If
    
    ' Итоговый статус
    readinessText = readinessText & vbCrLf & "[СТАТУС] "
    If isReady Then
        readinessText = readinessText & "СИСТЕМА ГОТОВА К РАБОТЕ"
        MsgBox readinessText, vbInformation, "Проверка готовности"
    Else
        readinessText = readinessText & "СИСТЕМА НЕ ГОТОВА - УСТРАНИТЕ ОШИБКИ"
        MsgBox readinessText, vbCritical, "Проверка готовности"
    End If
End Sub

'/** Обработчик кнопки "ПриказЗаРиск" */
Public Sub OnRiskOrderClick(control As IRibbonControl)
    Call mdlRiskExport.ExportRiskAllowanceOrder
End Sub

' Было: Call mdlPeriodsExport.ExportPeriodsToExcel_WithChoice
' Стало:
Public Sub OnPeriodsReportClick(control As IRibbonControl)
    Call mdlFRPExport.ExportPeriodsToExcel_WithChoice
End Sub



