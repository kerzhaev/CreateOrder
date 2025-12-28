Attribute VB_Name = "mdlPaymentTypes"
' ===============================================================================
' Модуль mdlPaymentTypes
' Версия: 1.0.0
' Дата: 01.12.2025
' Описание: Конфигурация типов выплат для системы надбавок без периодов
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' ===============================================================================

Option Explicit

' Константы
Public Const DEFAULT_TEMPLATE As String = "Шаблон_Универсальный.docx"

' Тип для конфигурации типа выплаты
Public Type PaymentTypeConfig
    typeName As String              ' "Водители СдЕ"
    TypeCode As String              ' "DRIVER_SDE"
    WordTemplate As String          ' "Шаблон_Водители.docx"
    Description As String           ' Описание
End Type

' Тип для данных о выплате без периодов
Public Type PaymentWithoutPeriod
    fio As String
    lichniyNomer As String
    Rank As String                  ' Из листа "Штат"
    Position As String              ' Из листа "Штат"
    VoinskayaChast As String        ' Из листа "Штат"
    paymentType As String
    amount As String
    foundation As String
End Type

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получение конфигурации типа выплаты из справочника
' @param paymentType As String - название типа выплаты
' @return PaymentTypeConfig - конфигурация типа выплаты
' =============================================
Public Function GetPaymentTypeConfig(ByVal paymentType As String) As PaymentTypeConfig
    On Error GoTo ErrorHandler
    
    Dim configDict As Object
    Dim config As PaymentTypeConfig
    
    ' Получаем конфигурацию из справочника
    Set configDict = mdlReferenceData.GetPaymentTypeConfig(paymentType)
    
    ' Если конфигурация найдена
    If configDict.count > 0 Then
        config.typeName = CStr(configDict("TypeName"))
        config.TypeCode = CStr(configDict("TypeCode"))
        config.WordTemplate = CStr(configDict("WordTemplate"))
        config.Description = CStr(configDict("Description"))
    Else
        ' Конфигурация по умолчанию
        config.typeName = paymentType
        config.TypeCode = ""
        config.WordTemplate = DEFAULT_TEMPLATE
        config.Description = "Тип выплаты: " & paymentType
    End If
    
    GetPaymentTypeConfig = config
    Exit Function
    
ErrorHandler:
    ' Возвращаем конфигурацию по умолчанию при ошибке
    config.typeName = paymentType
    config.TypeCode = ""
    config.WordTemplate = DEFAULT_TEMPLATE
    config.Description = "Тип выплаты: " & paymentType
    GetPaymentTypeConfig = config
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получение полного пути к шаблону Word
' @param templateName As String - имя файла шаблона
' @return String - полный путь к шаблону или пустая строка если файл не найден
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
    
    ' Проверяем существование файла
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получение пути к шаблону с учетом приоритета (тип -> универсальный -> отсутствует)
' @param config As PaymentTypeConfig - конфигурация типа выплаты
' @return String - полный путь к шаблону или пустая строка если шаблоны отсутствуют
' =============================================
Public Function GetTemplatePathWithFallback(ByRef config As PaymentTypeConfig) As String
    On Error GoTo ErrorHandler
    
    Dim templatePath As String
    
    ' 1. Пробуем шаблон для типа выплаты
    If config.WordTemplate <> "" Then
        templatePath = GetTemplatePath(config.WordTemplate)
        If templatePath <> "" Then
            GetTemplatePathWithFallback = templatePath
            Exit Function
        End If
    End If
    
    ' 2. Пробуем единый универсальный шаблон
    templatePath = GetTemplatePath(DEFAULT_TEMPLATE)
    If templatePath <> "" Then
        GetTemplatePathWithFallback = templatePath
        Exit Function
    End If
    
    ' 3. Шаблоны отсутствуют
    GetTemplatePathWithFallback = ""
    Exit Function
    
ErrorHandler:
    GetTemplatePathWithFallback = ""
End Function

