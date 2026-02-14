Attribute VB_Name = "mdlPaymentTypes"
' ===============================================================================
' Module mdlPaymentTypes
' Version: 1.0.0
' Date: 14.02.2026
' Description: Configuration of payment types for the system of allowances without periods
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

' Constants
Public Const DEFAULT_TEMPLATE As String = "Шаблон_Универсальный.docx"

' Type for payment type configuration
Public Type PaymentTypeConfig
    typeName As String              ' "Drivers CDE"
    TypeCode As String              ' "DRIVER_SDE"
    WordTemplate As String          ' "Шаблон_Водители.docx"
    Description As String           ' Description
End Type

' Type for payment data without periods
Public Type PaymentWithoutPeriod
    fio As String
    lichniyNomer As String
    Rank As String                  ' From sheet "Staff" (Shtat)
    Position As String              ' From sheet "Staff" (Shtat)
    VoinskayaChast As String        ' From sheet "Staff" (Shtat)
    paymentType As String
    amount As String
    foundation As String
End Type

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Get payment type configuration from reference
' @param paymentType As String - payment type name
' @return PaymentTypeConfig - payment type configuration
' =============================================
Public Function GetPaymentTypeConfig(ByVal paymentType As String) As PaymentTypeConfig
    On Error GoTo ErrorHandler
    
    Dim configDict As Object
    Dim config As PaymentTypeConfig
    
    ' Get configuration from reference
    Set configDict = mdlReferenceData.GetPaymentTypeConfig(paymentType)
    
    ' If configuration is found
    If configDict.count > 0 Then
        config.typeName = CStr(configDict("TypeName"))
        config.TypeCode = CStr(configDict("TypeCode"))
        config.WordTemplate = CStr(configDict("WordTemplate"))
        config.Description = CStr(configDict("Description"))
    Else
        ' Default configuration
        config.typeName = paymentType
        config.TypeCode = ""
        config.WordTemplate = DEFAULT_TEMPLATE
        config.Description = "Тип выплаты: " & paymentType
    End If
    
    GetPaymentTypeConfig = config
    Exit Function
    
ErrorHandler:
    ' Return default configuration on error
    config.typeName = paymentType
    config.TypeCode = ""
    config.WordTemplate = DEFAULT_TEMPLATE
    config.Description = "Тип выплаты: " & paymentType
    GetPaymentTypeConfig = config
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Get full path to Word template
' @param templateName As String - template file name
' @return String - full path to template or empty string if file not found
' =============================================
Public Function GetTemplatePath(ByVal templateName As String) As String
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim basePath As String
    
    basePath = ThisWorkbook.Path
    If Right(basePath, 1) <> "\" Then
        basePath = basePath & "\"
    End If
    
    filePath = basePath & templateName
    
    ' Check file existence
    If dir(filePath) <> "" Then
        GetTemplatePath = filePath
    Else
        GetTemplatePath = ""
    End If
    
    Exit Function
    
ErrorHandler:
    GetTemplatePath = ""
End Function

' =============================================
' @author Kerzhaev Evgeniy, FKU "95 FES" MO RF
' @description Get template path with priority (type -> universal -> missing)
' @param config As PaymentTypeConfig - payment type configuration
' @return String - full path to template or empty string if templates are missing
' =============================================
Public Function GetTemplatePathWithFallback(ByRef config As PaymentTypeConfig) As String
    On Error GoTo ErrorHandler
    
    Dim templatePath As String
    
    ' 1. Try template for payment type
    If config.WordTemplate <> "" Then
        templatePath = GetTemplatePath(config.WordTemplate)
        If templatePath <> "" Then
            GetTemplatePathWithFallback = templatePath
            Exit Function
        End If
    End If
    
    ' 2. Try single universal template
    templatePath = GetTemplatePath(DEFAULT_TEMPLATE)
    If templatePath <> "" Then
        GetTemplatePathWithFallback = templatePath
        Exit Function
    End If
    
    ' 3. Templates are missing
    GetTemplatePathWithFallback = ""
    Exit Function
    
ErrorHandler:
    GetTemplatePathWithFallback = ""
End Function
