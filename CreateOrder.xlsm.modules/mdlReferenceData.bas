Attribute VB_Name = "mdlReferenceData"
' ===============================================================================
' Модуль mdlReferenceData
' Версия: 1.0.0
' Дата: 01.12.2025
' Описание: Работа со справочниками для системы надбавок без периодов
' Автор: Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' ===============================================================================

Option Explicit

' Константы имен листов
Public Const SHEET_PAYMENTS_NO_PERIODS As String = "Выплаты_Без_Периодов"
Public Const SHEET_REF_VUS_CREW As String = "Справочник_ВУС_Экипаж"
Public Const SHEET_REF_PAYMENT_TYPES As String = "Справочник_Типы_Выплат"
Public Const SHEET_STAFF As String = "Штат"

' Тип для пары ВУС-Должность
Public Type VUSPositionPair
    vus As String
    Position As String
End Type

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Инициализация листа справочников (если создается программно)
' @param ws Worksheet - лист для инициализации
' =============================================
Public Sub InitializeReferencesSheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    ' Создание структуры заголовков для справочников
    ' Эта функция может быть использована для программного создания листов
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при инициализации листа справочников: " & Err.Description, vbCritical, "Ошибка"
End Sub

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Загрузка всех пар ВУС-должность из справочника в Dictionary
' @return Object (Dictionary) - словарь, где ключ = "ВУС|Должность", значение = True
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
    
    ' Ищем лист "Справочник_ВУС_Экипаж"
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
    
    ' Загружаем данные (предполагаем: колонка A = ВУС, колонка B = Должность)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Пропускаем заголовок
        vus = Trim(LCase(CStr(wsRef.Cells(i, 1).value)))
        Position = Trim(LCase(CStr(wsRef.Cells(i, 2).value)))
        
        If vus <> "" And Position <> "" Then
            ' Создаем ключ из пары ВУС-Должность
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
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Проверка наличия пары (ВУС, должность) в справочнике
' @param vus As String - ВУС военнослужащего
' @param position As String - должность военнослужащего
' @return Boolean - True если пара найдена в справочнике
' =============================================
Public Function CheckVUSPositionPair(ByVal vus As String, ByVal Position As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim refVUS As String
    Dim refPosition As String
    
    ' Ищем лист "Справочник_ВУС_Экипаж"
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
    
    ' Нормализуем входные данные для сравнения
    vus = Trim(LCase(vus))
    Position = Trim(LCase(Position))
    
    ' Проверяем каждую строку справочника
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Пропускаем заголовок
        refVUS = Trim(LCase(CStr(wsRef.Cells(i, 1).value)))
        refPosition = Trim(LCase(CStr(wsRef.Cells(i, 2).value)))
        
        ' Проверяем совпадение пары ВУС-должность
        If refVUS = vus And refPosition = Position Then
            CheckVUSPositionPair = True
            Exit Function
        End If
    Next i
    
    ' Пара не найдена
    CheckVUSPositionPair = False
    Exit Function
    
ErrorHandler:
    CheckVUSPositionPair = False
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получение конфигурации типа выплаты из справочника
' @param paymentType As String - название типа выплаты
' @return Object (Dictionary) - конфигурация типа выплаты или Nothing
' =============================================
Public Function GetPaymentTypeConfig(ByVal paymentType As String) As Object
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim resultDict As Object
    Dim refTypeName As String
    
    Set resultDict = CreateObject("Scripting.Dictionary")
    
    ' Ищем лист "Справочник_Типы_Выплат"
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
    
    ' Ищем тип выплаты (колонка A = Тип выплаты, B = Код, C = Шаблон Word, D = Описание)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    paymentType = Trim(LCase(paymentType))
    
    For i = 2 To lastRow ' Пропускаем заголовок
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
    
    ' Тип не найден
    Set GetPaymentTypeConfig = resultDict
    Exit Function
    
ErrorHandler:
    Set GetPaymentTypeConfig = CreateObject("Scripting.Dictionary")
End Function

' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Получение списка всех типов выплат из справочника
' @return Collection - коллекция строк с названиями типов выплат
' =============================================
Public Function GetAllPaymentTypes() As Collection
    On Error GoTo ErrorHandler
    
    Dim wsRef As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As Collection
    Dim typeName As String
    
    Set result = New Collection
    
    ' Ищем лист "Справочник_Типы_Выплат"
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
    
    ' Загружаем все типы выплат (колонка A = Тип выплаты)
    lastRow = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row
    
    For i = 2 To lastRow ' Пропускаем заголовок
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

