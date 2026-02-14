Attribute VB_Name = "mdlReferenceData"
' ===============================================================================
' Module mdlReferenceData
' Version: 1.0.0
' Date: 14.02.2026
' Description: Handling reference data for the allowance system without periods
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' Sheet name constants
Public Const SHEET_PAYMENTS_NO_PERIODS As String = "Выплаты_Без_Периодов"
Public Const SHEET_REF_VUS_CREW As String = "Справочник_ВУС_Экипаж"
Public Const SHEET_REF_PAYMENT_TYPES As String = "Справочник_Типы_Выплат"
Public Const SHEET_STAFF As String = "Штат"

' Type for VUS-Position pair
Public Type VUSPositionPair
    vus As String
    Position As String
End Type

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Initialization of reference sheet (if created programmatically)
' @param ws Worksheet - sheet to initialize
' =============================================
Public Sub InitializeReferencesSheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Creation of header structure for references
    ' This function can be used for programmatic sheet creation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при инициализации листа справочников: " & Err.Description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Load all VUS-Position pairs from reference into Dictionary
' @return Object (Dictionary) - dictionary where key = "VUS|Position", value = True
' =============================================
Public Function LoadVUSPositionPairs() As Object
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As Object
    Dim vus As String
    Dim Position As String
    Dim key As String
    
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Search for sheet "Справочник_ВУС_Экипаж"
    Set wsRef = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_REF_VUS_CREW Then
            Set wsRef = ws
            Exit For
        End If
    Next ws
    
    If wsRef Is Nothing Then
        Set LoadVUSPositionPairs = result
        Exit Function
    End If
    
    ' Load data (assume: column A = VUS, column B = Position)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Skip header
        vus = Trim(LCase(CStr(wsRef.Cells(i, 1).value)))
        Position = Trim(LCase(CStr(wsRef.Cells(i, 2).value)))
        
        If vus <> "" And Position <> "" Then
            ' Create key from VUS-Position pair
            key = vus & "|" & Position
            If Not result.exists(key) Then
                result.Add key, True
            End If
        End If
    Next i
    
    Set LoadVUSPositionPairs = result
    Exit Function
    
ErrorHandler:
    Set LoadVUSPositionPairs = CreateObject("Scripting.Dictionary")
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Check existence of (VUS, Position) pair in reference
' @param vus As String - soldier's VUS
' @param position As String - soldier's position
' @return Boolean - True if pair found in reference
' =============================================
Public Function CheckVUSPositionPair(ByVal vus As String, ByVal Position As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim refVUS As String
    Dim refPosition As String
    
    ' Search for sheet "Справочник_ВУС_Экипаж"
    Set wsRef = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_REF_VUS_CREW Then
            Set wsRef = ws
            Exit For
        End If
    Next ws
    
    If wsRef Is Nothing Then
        CheckVUSPositionPair = False
        Exit Function
    End If
    
    ' Normalize input data for comparison
    vus = Trim(LCase(vus))
    Position = Trim(LCase(Position))
    
    ' Check every row of the reference
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Skip header
        refVUS = Trim(LCase(CStr(wsRef.Cells(i, 1).value)))
        refPosition = Trim(LCase(CStr(wsRef.Cells(i, 2).value)))
        
        ' Check VUS-Position pair match
        If refVUS = vus And refPosition = Position Then
            CheckVUSPositionPair = True
            Exit Function
        End If
    Next i
    
    ' Pair not found
    CheckVUSPositionPair = False
    Exit Function
    
ErrorHandler:
    CheckVUSPositionPair = False
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Get payment type configuration from reference
' @param paymentType As String - payment type name
' @return Object (Dictionary) - payment type configuration or Nothing
' =============================================
Public Function GetPaymentTypeConfig(ByVal paymentType As String) As Object
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim resultDict As Object
    Dim refTypeName As String
    
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    ' Search for sheet "Справочник_Типы_Выплат"
    Set wsRef = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_REF_PAYMENT_TYPES Then
            Set wsRef = ws
            Exit For
        End If
    Next ws
    
    If wsRef Is Nothing Then
        Set GetPaymentTypeConfig = resultDict
        Exit Function
    End If
    
    ' Search for payment type (col A = Type, B = Code, C = Template, D = Description)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    paymentType = Trim(LCase(paymentType))
    
    For i = 2 To lastRow ' Skip header
        refTypeName = Trim(LCase(CStr(wsRef.Cells(i, 1).value)))
        
        If refTypeName = paymentType Then
            resultDict("TypeName") = Trim(CStr(wsRef.Cells(i, 1).value))
            resultDict("TypeCode") = Trim(CStr(wsRef.Cells(i, 2).value))
            resultDict("WordTemplate") = Trim(CStr(wsRef.Cells(i, 3).value))
            resultDict("Description") = Trim(CStr(wsRef.Cells(i, 4).value))
            Set GetPaymentTypeConfig = resultDict
            Exit Function
        End If
    Next i
    
    ' Type not found
    Set GetPaymentTypeConfig = resultDict
    Exit Function
    
ErrorHandler:
    Set GetPaymentTypeConfig = CreateObject("Scripting.Dictionary")
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Get list of all payment types from reference
' @return Collection - collection of payment type name strings
' =============================================
Public Function GetAllPaymentTypes() As Collection
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As Collection
    Dim typeName As String
    
    Set result = New Collection
    
    ' Search for sheet "Справочник_Типы_Выплат"
    Set wsRef = Nothing
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_REF_PAYMENT_TYPES Then
            Set wsRef = ws
            Exit For
        End If
    Next ws
    
    If wsRef Is Nothing Then
        Set GetAllPaymentTypes = result
        Exit Function
    End If
    
    ' Load all payment types (column A = Payment Type)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Skip header
        typeName = Trim(CStr(wsRef.Cells(i, 1).value))
        If typeName <> "" Then
            result.Add typeName
        End If
    Next i
    
    Set GetAllPaymentTypes = result
    Exit Function
    
ErrorHandler:
    Set GetAllPaymentTypes = New Collection
End Function
