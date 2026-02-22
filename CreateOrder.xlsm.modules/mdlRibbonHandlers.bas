Attribute VB_Name = "mdlRibbonHandlers"
' ===============================================================================
' Module mdlRibbonHandlers for handling custom ribbon events
' Version: 2.2.0 (License Enforced)
' Date: 17.02.2026
' Author: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' Description: Full set of event handlers for custom ribbon buttons
' Functionality:
' - PREMIUM: Export to Word (Main, Spravka, Raport, Risk, Allowances)
' - PREMIUM: Export to Excel (Alushta/FRP)
' - FREE: Data management (import, preview, validation, settings)
' ===============================================================================

Option Explicit

' ===============================================================================
' PREMIUM FUNCTIONS (Require active license)
' ===============================================================================

' Handler for main order
Sub RunMainExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Application.ScreenUpdating = False
    Call mdlMainExport.ExportToWordFromStaffByLichniyNomer
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании основного приказа: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for DSO certificate (spravka)
Sub RunSpravkaExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Application.ScreenUpdating = False
    Call mdlSpravkaExport.ExportToWordSpravkaFromTemplate
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании справки: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for report (raport) with CHOICE
Sub RunRaportExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    ' Спрашиваем пользователя
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Какой рапорт необходимо сформировать?" & vbCrLf & vbCrLf & _
                    "Да - Рапорт на ДСО (Сутки отдыха)" & vbCrLf & _
                    "Нет - Рапорт на РИСК (Денежная выплата)" & vbCrLf & _
                    "Отмена - Выход", vbYesNoCancel + vbQuestion, "Выбор типа рапорта")
    
    If choice = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    
    If choice = vbYes Then
        ' === РАПОРТ ДСО ===
        Call mdlRaportExport.ExportToWordRaportFromTemplateByLichniyNomer
    Else
        ' === РАПОРТ РИСК ===
        MsgBox "Функционал отдельного рапорта на Риск пока в разработке. Используется стандартный шаблон.", vbInformation
        Call mdlRaportExport.ExportToWordRaportFromTemplateByLichniyNomer
    End If
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при создании рапорта: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "OrderForRisk" button
Public Sub OnRiskOrderClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Call mdlRiskExport.ExportRiskAllowanceOrder
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при вызове приказа за риск: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "Export Allowances" button (allowances without periods)
Public Sub OnExportAllowancesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Application.ScreenUpdating = False
    Call mdlUniversalPaymentExport.ExportPaymentsWithoutPeriods
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при экспорте надбавок: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for Excel reports (Alushta / FRP)
Public Sub OnPeriodsReportClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Call mdlFRPExport.ExportPeriodsToExcel_WithChoice
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при создании Excel отчета: " & Err.Description, vbCritical, "Ошибка"
End Sub

' ===============================================================================
' FREE FUNCTIONS (No license required)
' ===============================================================================

' Handler for data validation

' =============================================
' Умная валидация: сама понимает, какой лист проверять
' =============================================
Sub RunSmartValidation(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    If ActiveSheet Is Nothing Then Exit Sub
    
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    
    If sheetName = "ДСО" Then
        ' 1. Если мы в ДСО -> проверяем даты
        Application.ScreenUpdating = False
        Application.StatusBar = "Проверка периодов ДСО..."
        Call mdlDataValidation.ValidateMainSheetData
        Application.ScreenUpdating = True
        
    ElseIf sheetName = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        ' 2. Если мы в надбавках -> проверяем документы
        Application.ScreenUpdating = False
        Call mdlPaymentValidation.ValidatePaymentsWithoutPeriods
        Application.ScreenUpdating = True
        
    Else
        ' 3. Если мы на любом другом листе
        MsgBox "Для проверки данных перейдите на лист 'ДСО' или '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbInformation, "Умная проверка"
    End If
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Ошибка при проверке данных: " & Err.Description, vbCritical, "Ошибка"
End Sub





' Handler for structure diagnostics
Sub RunDiagnoseStructure(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call mdlDataValidation.DiagnoseWorkbookStructure
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при диагностике структуры: " & Err.Description, vbCritical, "Ошибка диагностики"
End Sub

' Handler for data import
Sub RunImportData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call mdlDataImport.ImportDataToStaff
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при импорте данных: " & Err.Description, vbCritical, "Ошибка импорта"
End Sub

' Handler for data preview
Sub RunPreviewData(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call mdlDataImport.PreviewImportData
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при предпросмотре данных: " & Err.Description, vbCritical, "Ошибка предпросмотра"
End Sub


' Handler for Word Raport Import
Sub RunWordRaportImport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' === ПРОВЕРКА ЛИЦЕНЗИИ ===
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Инициализация импорта рапорта..."
    
    ' Вызов главной функции импорта из нового модуля
    Call mdlWordImport.ExecuteWordImport
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Ошибка при вызове импорта: " & Err.Description, vbCritical, "Ошибка"
End Sub


' Handler for "Mass Add Employees" button
Public Sub OnMassImportEmployeesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' Check active sheet
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox "Для массового добавления перейдите на лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    Call mdlUniversalPaymentExport.ImportEmployeesByNumbers
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при массовом добавлении сотрудников: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "Select Employee" button (opening selection form)
Public Sub OnSelectEmployeeClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim activeCell As Range
    Dim targetRow As Long
    
    ' Check active sheet
    If ActiveSheet Is Nothing Then
        MsgBox "Нет активного листа.", vbExclamation, "Внимание"
        Exit Sub
    End If

    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox "Для выбора сотрудника перейдите на лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    Set wsPayments = ActiveSheet

    On Error Resume Next
    Set activeCell = Application.activeCell
    On Error GoTo ErrorHandler
    
    If activeCell Is Nothing Then
        MsgBox "Нет активной ячейки.", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    ' Check if active cell is in column C or D
    If activeCell.Column <> mdlPaymentValidation.COL_FIO And activeCell.Column <> mdlPaymentValidation.COL_LICHNIY_NOMER Then
        MsgBox "Активная ячейка должна находиться в колонке C (ФИО) или D (личный номер).", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    ' Determine target row
    targetRow = activeCell.Row
    
    ' Open employee selection form
    frmSelectEmployee.selectedLichniyNomer = ""
    frmSelectEmployee.selectedFIO = ""
    frmSelectEmployee.isCancelled = True
    
    frmSelectEmployee.Show
    
    ' If selection made (not cancelled), fill cells
    If Not frmSelectEmployee.isCancelled Then
        wsPayments.Cells(targetRow, mdlPaymentValidation.COL_FIO).value = frmSelectEmployee.selectedFIO
        wsPayments.Cells(targetRow, mdlPaymentValidation.COL_LICHNIY_NOMER).value = frmSelectEmployee.selectedLichniyNomer
    End If

    ' Cleanup objects
    Set activeCell = Nothing
    Set wsPayments = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при выборе сотрудника: " & Err.Description, vbCritical, "Ошибка"
    Set activeCell = Nothing
    Set wsPayments = Nothing
End Sub

' Handler for "References" button
Public Sub OnManageReferencesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Dim wsRef As Worksheet
    On Error Resume Next
    Set wsRef = ThisWorkbook.Sheets(mdlReferenceData.SHEET_REF_PAYMENT_TYPES)
    If wsRef Is Nothing Then
        MsgBox "Лист справочников не найден. Убедитесь, что лист '" & mdlReferenceData.SHEET_REF_PAYMENT_TYPES & "' существует.", vbInformation, "Справочники"
    Else
        wsRef.Activate
        wsRef.Cells(1, 1).Select
    End If
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при открытии справочников: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for help
Sub ShowHelp(control As IRibbonControl)
    Dim helpText As String
    helpText = "=== МАКРОСЫ ДЛЯ РАБОТЫ С ДАННЫМИ СВО ===" & vbCrLf & vbCrLf
    helpText = helpText & "[ЭКСПОРТ] ЭКСПОРТ ДОКУМЕНТОВ (Требуется активация):" & vbCrLf
    helpText = helpText & "• Основной приказ - создает Word документ с приказом в дательном падеже" & vbCrLf
    helpText = helpText & "• Справка ДСО - создает справки на основе шаблона Word" & vbCrLf
    helpText = helpText & "• Рапорт - создает рапорты о выплате компенсации" & vbCrLf & vbCrLf
    helpText = helpText & "[ОТЧЕТЫ] ОТЧЕТЫ (Требуется активация):" & vbCrLf
    helpText = helpText & "• Отчет по периодам - создает сводный Excel отчет" & vbCrLf & vbCrLf
    helpText = helpText & "[ДАННЫЕ] УПРАВЛЕНИЕ ДАННЫМИ (Свободный доступ):" & vbCrLf
    helpText = helpText & "• Импорт данных - загружает данные из Excel в лист 'Штат'" & vbCrLf
    helpText = helpText & "• Предпросмотр - показывает предварительный просмотр файла" & vbCrLf & vbCrLf
    helpText = helpText & "[ВАЛИДАЦИЯ] ПРОВЕРКА ДАННЫХ (Свободный доступ):" & vbCrLf
    helpText = helpText & "• Проверить данные - выполняет полную валидацию листа ДСО" & vbCrLf
    helpText = helpText & "• Диагностика структуры - анализирует структуру листов" & vbCrLf & vbCrLf
    helpText = helpText & "[ТРЕБОВАНИЯ] ТРЕБОВАНИЯ:" & vbCrLf
    helpText = helpText & "• Шаблоны Word должны находиться в папке с Excel файлом" & vbCrLf
    helpText = helpText & "• Лист 'Штат' должен содержать данные о сотрудниках" & vbCrLf
    helpText = helpText & "• Столбец 'Личный номер' обязателен для уникальной идентификации"
    
    MsgBox helpText, vbInformation, "Справка по макросам СВО"
End Sub

' Handler for settings
Sub ShowSettings(control As IRibbonControl)
    Dim settingsText As String
    settingsText = "=== НАСТРОЙКИ МАКРОСОВ ===" & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПАПКА] Текущая папка: " & ThisWorkbook.Path & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПРОВЕРКА] Проверка шаблонов:" & vbCrLf
    
    ' Check templates existence
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
    
    settingsText = settingsText & vbCrLf & "[СТАТУС АКТИВАЦИЯ]: "
    Select Case modActivation.GetLicenseStatus()
        Case 0: settingsText = settingsText & "АКТИВНА (до " & modActivation.GetLicenseExpiryDateStr() & ")" & vbCrLf
        Case 1: settingsText = settingsText & "НЕ АКТИВИРОВАНО" & vbCrLf
        Case 2: settingsText = settingsText & "БЛОКИРОВКА (Сбой системного времени)" & vbCrLf
    End Select
    
    settingsText = settingsText & vbCrLf & "[ВЕРСИЯ] Версия макросов: 2.2.0"
    
    MsgBox settingsText, vbInformation, "Настройки и проверка"
    
    ' Если пользователь нажал настройки, и лицензии нет — предложим активировать.
    If modActivation.GetLicenseStatus() <> 0 Then
        frmAbout.Show
    End If
End Sub

' Function to check system readiness
Sub CheckSystemReadiness(control As IRibbonControl)
    Dim readinessText As String
    Dim isReady As Boolean
    isReady = True
    
    readinessText = "=== ПРОВЕРКА ГОТОВНОСТИ СИСТЕМЫ ===" & vbCrLf & vbCrLf
    
    ' Check templates
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
    
    ' Check sheets
    readinessText = readinessText & vbCrLf & "[СТРУКТУРА ДАННЫХ]" & vbCrLf
    Dim wsExists As Boolean
    wsExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Штат" Then
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
    
    ' Check License
    readinessText = readinessText & vbCrLf & "[ЛИЦЕНЗИЯ]" & vbCrLf
    If modActivation.GetLicenseStatus() = 0 Then
        readinessText = readinessText & "[OK] Лицензия активна" & vbCrLf
    Else
        readinessText = readinessText & "[ПРЕДУПРЕЖДЕНИЕ] Требуется активация для работы модулей экспорта" & vbCrLf
        isReady = False
    End If
    
    ' Final status
    readinessText = readinessText & vbCrLf & "[СТАТУС] "
    If isReady Then
        readinessText = readinessText & "СИСТЕМА ГОТОВА К РАБОТЕ"
        MsgBox readinessText, vbInformation, "Проверка готовности"
    Else
        readinessText = readinessText & "СИСТЕМА ТРЕБУЕТ ВНИМАНИЯ"
        MsgBox readinessText, vbCritical, "Проверка готовности"
    End If
End Sub

' Handler for "Remove Duplicate Modules" button
Public Sub OnRemoveDuplicateModulesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call MdlBackup.RemoveDuplicateModules
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Ошибка при удалении дубликатов модулей: " & Err.Description, vbCritical, "Ошибка"
End Sub

