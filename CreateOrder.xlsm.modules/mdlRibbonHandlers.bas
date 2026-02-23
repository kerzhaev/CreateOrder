Attribute VB_Name = "mdlRibbonHandlers"
' ===============================================================================
' Module mdlRibbonHandlers for handling custom ribbon events
' Version: 2.3.0 (Updated UI & Licensing integration)
' Date: 23.02.2026
' Author: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' ===============================================================================
Option Explicit

' ===============================================================================
' PREMIUM FUNCTIONS (Require active license / free period)
' ===============================================================================

' Handler for main order
Sub RunMainExport(control As IRibbonControl)
    On Error GoTo ErrorHandler
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
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Какой рапорт необходимо сформировать?" & vbCrLf & vbCrLf & _
                    "Да - Рапорт на ДСО (Сутки отдыха)" & vbCrLf & _
                    "Нет - Рапорт на РИСК (Денежная выплата)" & vbCrLf & _
                    "Отмена - Выход", vbYesNoCancel + vbQuestion, "Выбор типа рапорта")
    
    If choice = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    If choice = vbYes Then
        Call mdlRaportExport.ExportToWordRaportFromTemplateByLichniyNomer
    Else
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
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Call mdlRiskExport.ExportRiskAllowanceOrder
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка при вызове приказа за риск: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "Export Allowances" button
Public Sub OnExportAllowancesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
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
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Call mdlFRPExport.ExportPeriodsToExcel_WithChoice
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка при создании Excel отчета: " & Err.Description, vbCritical, "Ошибка"
End Sub

' ===============================================================================
' FREE FUNCTIONS (No license required)
' ===============================================================================

' Умная валидация: сама понимает, какой лист проверять
Sub RunSmartValidation(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If ActiveSheet Is Nothing Then Exit Sub
    Dim sheetName As String
    sheetName = ActiveSheet.Name
    
    If sheetName = "ДСО" Then
        Application.ScreenUpdating = False
        Application.StatusBar = "Проверка периодов ДСО..."
        Call mdlDataValidation.ValidateMainSheetData
        Application.ScreenUpdating = True
    ElseIf sheetName = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        Application.ScreenUpdating = False
        Call mdlPaymentValidation.ValidatePaymentsWithoutPeriods
        Application.ScreenUpdating = True
    Else
        MsgBox "Для проверки данных перейдите на лист 'ДСО' или '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbInformation, "Умная проверка"
    End If
    Application.StatusBar = False
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Ошибка при проверке данных: " & Err.Description, vbCritical, "Ошибка"
End Sub

' НОВОЕ: Обработчик кнопки "О программе" (Заменил Диагностику)
Sub RunShowAbout(control As IRibbonControl)
    On Error GoTo ErrorHandler
    frmAbout.Show
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка при открытии окна программы: " & Err.Description, vbCritical, "Ошибка"
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

' Handler for Word Raport Import
Sub RunWordRaportImport(control As IRibbonControl)
    On Error GoTo ErrorHandler
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub
    Application.ScreenUpdating = False
    Application.StatusBar = "Инициализация импорта рапорта..."
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
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox "Для массового добавления перейдите на лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, "Внимание"
        Exit Sub
    End If
    Call mdlUniversalPaymentExport.ImportEmployeesByNumbers
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка при массовом добавлении сотрудников: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "Select Employee" button
Public Sub OnSelectEmployeeClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Dim wsPayments As Worksheet, activeCell As Range, targetRow As Long
    
    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
        MsgBox "Для выбора сотрудника перейдите на лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "'.", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    Set wsPayments = ActiveSheet
    On Error Resume Next
    Set activeCell = Application.activeCell
    On Error GoTo ErrorHandler
    
    If activeCell Is Nothing Then Exit Sub
    
    If activeCell.Column <> mdlPaymentValidation.COL_FIO And activeCell.Column <> mdlPaymentValidation.COL_LICHNIY_NOMER Then
        MsgBox "Активная ячейка должна находиться в колонке C (ФИО) или D (личный номер).", vbExclamation, "Внимание"
        Exit Sub
    End If
    
    targetRow = activeCell.Row
    frmSelectEmployee.selectedLichniyNomer = ""
    frmSelectEmployee.selectedFIO = ""
    frmSelectEmployee.isCancelled = True
    frmSelectEmployee.Show
    
    If Not frmSelectEmployee.isCancelled Then
        wsPayments.Cells(targetRow, mdlPaymentValidation.COL_FIO).value = frmSelectEmployee.selectedFIO
        wsPayments.Cells(targetRow, mdlPaymentValidation.COL_LICHNIY_NOMER).value = frmSelectEmployee.selectedLichniyNomer
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка при выборе сотрудника: " & Err.Description, vbCritical, "Ошибка"
End Sub

' Handler for "References" button
Public Sub OnManageReferencesClick(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Dim wsRef As Worksheet
    On Error Resume Next
    Set wsRef = ThisWorkbook.Sheets(mdlReferenceData.SHEET_REF_PAYMENT_TYPES)
    If wsRef Is Nothing Then
        MsgBox "Лист справочников не найден.", vbInformation, "Справочники"
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

' Handler for settings (Обновлено под новую систему лицензирования)
Sub ShowSettings(control As IRibbonControl)
    Dim settingsText As String
    settingsText = "=== НАСТРОЙКИ МАКРОСОВ ===" & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПАПКА] Текущая папка: " & ThisWorkbook.Path & vbCrLf & vbCrLf
    settingsText = settingsText & "[ПРОВЕРКА] Проверка шаблонов:" & vbCrLf
    
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
    
    settingsText = settingsText & vbCrLf & "[СТАТУС АКТИВАЦИИ]: "
    
    ' Новая расширенная проверка статусов лицензии
    Select Case modActivation.GetLicenseStatus()
        Case 0: settingsText = settingsText & "ПЕРСОНАЛЬНАЯ ЛИЦЕНЗИЯ (до " & modActivation.GetLicenseExpiryDateStr() & ")" & vbCrLf
        Case 3: settingsText = settingsText & "КОРПОРАТИВНАЯ ЛИЦЕНЗИЯ (до " & modActivation.GetLicenseExpiryDateStr() & ")" & vbCrLf
        Case 4: settingsText = settingsText & "ОЗНАКОМИТЕЛЬНЫЙ ПЕРИОД (до " & modActivation.GetLicenseExpiryDateStr() & ")" & vbCrLf
        Case 1: settingsText = settingsText & "ОГРАНИЧЕННАЯ ВЕРСИЯ (Срок истек)" & vbCrLf
        Case 2: settingsText = settingsText & "БЛОКИРОВКА (Сбой системного времени)" & vbCrLf
    End Select
    
    settingsText = settingsText & vbCrLf & "[ВЕРСИЯ] Версия макросов: " & modActivation.PRODUCT_VERSION
    
    MsgBox settingsText, vbInformation, "Настройки и проверка"
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
    MsgBox "Ошибка при удалении дубликатов: " & Err.Description, vbCritical, "Ошибка"
End Sub

