Attribute VB_Name = "mdlPaymentValidation"
' ===============================================================================
' Module mdlPaymentValidation
' Version: 1.0.0
' Date: 14.02.2026
' Description: Validation of allowances without periods
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' Column index constants for sheet "Выплаты_Без_Периодов"
Public Const COL_NUMBER As Long = 1          ' A
Public Const COL_PAYMENT_TYPE As Long = 2    ' B
Public Const COL_FIO As Long = 3             ' C
Public Const COL_LICHNIY_NOMER As Long = 4   ' D
Public Const COL_AMOUNT As Long = 5          ' E
Public Const COL_FOUNDATION As Long = 6      ' F

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Main function for validating all allowances
' =============================================
Public Sub ValidatePaymentsWithoutPeriods()
    On Error GoTo ErrorHandler
    
    Dim wsPayments As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorCount As Long
    Dim warningCount As Long
    Dim reportText As String
    Dim isValid As Boolean
    Dim paymentType As String
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Выполняется валидация надбавок..."
    
    ' Search for sheet "Выплаты_Без_Периодов"
    Set wsPayments = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS Then
            Set wsPayments = ws
            Exit For
        End If
    Next ws
    
    If wsPayments Is Nothing Then
        MsgBox "Лист '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "' не найден.", vbCritical, "Ошибка"
        GoTo CleanUp
    End If
    
    lastRow = wsPayments.Cells(wsPayments.Rows.count, COL_LICHNIY_NOMER).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "В листе '" & mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS & "' нет данных для валидации.", vbInformation, "Информация"
        GoTo CleanUp
    End If
    
    errorCount = 0
    warningCount = 0
    reportText = "====== ОТЧЕТ О ВАЛИДАЦИИ НАДБАВОК ======" & vbCrLf & vbCrLf
    reportText = reportText & "Дата проверки: " & Format(Now, "dd.mm.yyyy hh:mm:ss") & vbCrLf
    reportText = reportText & "Проверено записей: " & (lastRow - 1) & vbCrLf & vbCrLf
    
    ' Check every row
    For i = 2 To lastRow
        Application.StatusBar = "Проверка строки " & i & " из " & lastRow
        
        paymentType = Trim(LCase(CStr(wsPayments.Cells(i, COL_PAYMENT_TYPE).value)))
        
        ' Call corresponding validation function
        Select Case paymentType
            Case "водители сдэ", "водители сде"
                isValid = ValidateDriverSDE(wsPayments, i)
            Case "экипаж"
                isValid = ValidateCrew(wsPayments, i)
            Case "физо"
                isValid = ValidateFIZO(wsPayments, i)
            Case "секретность"
                isValid = ValidateSecrecy(wsPayments, i)
            Case Else
                ' For unknown types - basic check
                isValid = ValidateBasic(wsPayments, i)
                If Not isValid Then
                    warningCount = warningCount + 1
                    reportText = reportText & "Строка " & i & ": Неизвестный тип выплаты '" & paymentType & "'" & vbCrLf
                End If
        End Select
        
        If Not isValid Then
            errorCount = errorCount + 1
            reportText = reportText & "Строка " & i & ": Ошибка валидации для типа '" & paymentType & "'" & vbCrLf
        End If
    Next i
    
    ' Final report
    reportText = reportText & vbCrLf & "Итого:" & vbCrLf
    reportText = reportText & "Ошибок: " & errorCount & vbCrLf
    reportText = reportText & "Предупреждений: " & warningCount & vbCrLf
    
    If errorCount = 0 And warningCount = 0 Then
        reportText = reportText & vbCrLf & "Все данные корректны!" & vbCrLf
        MsgBox reportText, vbInformation, "Валидация завершена"
    Else
        MsgBox reportText, vbExclamation, "Валидация завершена"
    End If
    
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Ошибка при валидации надбавок: " & Err.Description, vbCritical, "Ошибка"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Basic validation (check mandatory fields)
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Private Function ValidateBasic(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fio As String
    Dim lichniyNomer As String
    Dim amount As String
    Dim foundation As String
    
    fio = Trim(CStr(ws.Cells(rowNum, COL_FIO).value))
    lichniyNomer = Trim(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).value))
    amount = Trim(CStr(ws.Cells(rowNum, COL_AMOUNT).value))
    foundation = Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value))
    
    ' Check mandatory fields
    If fio = "" Or lichniyNomer = "" Or amount = "" Or foundation = "" Then
        ValidateBasic = False
        Exit Function
    End If
    
    ValidateBasic = True
    Exit Function
    
ErrorHandler:
    ValidateBasic = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Drivers CDE allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateDriverSDE(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lichniyNomer As String
    Dim foundation As String
    Dim staffData As Object
    Dim Position As String
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateDriverSDE = False
        Exit Function
    End If
    
    ' Get personal number
    lichniyNomer = Trim(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).value))
    
    ' Get data from staff
    Set staffData = mdlHelper.GetStaffData(lichniyNomer, True)
    If staffData.count = 0 Then
        ValidateDriverSDE = False
        Exit Function
    End If
    
    ' Check position (must be "driver" or "senior driver", NOT "mechanic-driver")
    Position = LCase(Trim(CStr(staffData("Штатная должность"))))
    If InStr(Position, "механик-водитель") > 0 Then
        ValidateDriverSDE = False
        Exit Function
    End If
    If InStr(Position, "водитель") = 0 And InStr(Position, "старший водитель") = 0 Then
        ValidateDriverSDE = False
        Exit Function
    End If
    
    ' Check foundation (must contain: copy of license, VAI certificate, brand, license plate)
    foundation = LCase(Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value)))
    
    Dim hasVU As Boolean, hasVAI As Boolean, hasMark As Boolean, hasGRZ As Boolean
    hasVU = (InStr(foundation, "удостоверен") > 0 Or InStr(foundation, "ву") > 0 Or InStr(foundation, "водительск") > 0)
    hasVAI = (InStr(foundation, "ваи") > 0)
    hasMark = (InStr(foundation, "марка") > 0)
    hasGRZ = (InStr(foundation, "грз") > 0 Or InStr(foundation, "регистрационный") > 0 Or InStr(foundation, "гос. номер") > 0)
    
    ValidateDriverSDE = (hasVU And hasVAI And hasMark And hasGRZ)
    Exit Function
    
ErrorHandler:
    ValidateDriverSDE = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Crew allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateCrew(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim lichniyNomer As String
    Dim staffData As Object
    Dim vus As String
    Dim Position As String
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateCrew = False
        Exit Function
    End If
    
    ' Get personal number
    lichniyNomer = Trim(CStr(ws.Cells(rowNum, COL_LICHNIY_NOMER).value))
    
    ' Get data from staff
    Set staffData = mdlHelper.GetStaffData(lichniyNomer, True)
    If staffData.count = 0 Then
        ValidateCrew = False
        Exit Function
    End If
    
    ' Get VUS and position (VUS might be in a separate column or needs a helper function)
    ' Using position for check for now
    Position = LCase(Trim(CStr(staffData("Штатная должность"))))
    
    ' TODO: Need to get VUS from "Staff" sheet - might need to add function to mdlHelper
    ' Checking only position for now
    ' Check VUS-Position pair in reference
    ' ValidateCrew = mdlReferenceData.CheckVUSPositionPair(vus, position)
    
    ' Temporary check: if position contains crew keywords
    Dim crewKeywords As Variant
    crewKeywords = Array("командир", "механик", "наводчик", "оператор", "экипаж")
    
    Dim i As Long
    Dim hasCrewKeyword As Boolean
    hasCrewKeyword = False
    For i = LBound(crewKeywords) To UBound(crewKeywords)
        If InStr(Position, CStr(crewKeywords(i))) > 0 Then
            hasCrewKeyword = True
            Exit For
        End If
    Next i
    
    ' TODO: Replace with reference check when VUS retrieval is implemented
    ValidateCrew = hasCrewKeyword
    Exit Function
    
ErrorHandler:
    ValidateCrew = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for FIZO allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateFIZO(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim foundation As String
    Dim vedomostCount As Long
    Dim i As Long
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateFIZO = False
        Exit Function
    End If
    
    ' Get foundation
    foundation = LCase(Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value)))
    
    ' Count occurrences of "vedomost"
    vedomostCount = 0
    i = 1
    Do While i <= Len(foundation)
        If Mid(foundation, i, 8) = "ведомость" Then
            vedomostCount = vedomostCount + 1
            i = i + 8
        Else
            i = i + 1
        End If
    Loop
    
    ' Must be at least 2 vedomosts
    ValidateFIZO = (vedomostCount >= 2)
    Exit Function
    
ErrorHandler:
    ValidateFIZO = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Validation for Secrecy allowance
' @param ws As Worksheet - sheet "Выплаты_Без_Периодов"
' @param rowNum As Long - row number to check
' @return Boolean - True if data is valid
' =============================================
Public Function ValidateSecrecy(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim foundation As String
    Dim hasForm As Boolean, hasNumber As Boolean, hasDate As Boolean, hasNomenclature As Boolean
    
    ' Basic check
    If Not ValidateBasic(ws, rowNum) Then
        ValidateSecrecy = False
        Exit Function
    End If
    
    ' Get foundation
    foundation = LCase(Trim(CStr(ws.Cells(rowNum, COL_FOUNDATION).value)))
    
    ' Check for mandatory elements
    hasForm = (InStr(foundation, "форма") > 0)
    hasNumber = (InStr(foundation, "номер") > 0)
    hasDate = (InStr(foundation, "дата") > 0)
    hasNomenclature = (InStr(foundation, "номенклатур") > 0 Or InStr(foundation, "пункт") > 0)
    
    ValidateSecrecy = (hasForm And hasNumber And hasDate And hasNomenclature)
    Exit Function
    
ErrorHandler:
    ValidateSecrecy = False
End Function
