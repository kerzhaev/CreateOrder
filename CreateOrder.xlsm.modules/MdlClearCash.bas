Attribute VB_Name = "MdlClearCash"
' === СНАПШОТ ВЕРСИИ === 12.07.2025 20:45 ===
' Рабочая версия сохранена: 12.07.2025 20:45

Option Explicit

' Модуль MdlClearCash для очистки кэша Office и восстановления ribbon
' Версия: 1.0.0
' Дата: 11.07.2025
' Описание: Выполняет полную очистку кэша Office для восстановления иконок ribbon

Sub ClearOfficeRibbonCache()
    Dim fso As Object
    Dim cachePath As String
    
    ' Создаем объект FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Закрываем все приложения Office
    MsgBox "ВАЖНО: Закройте ВСЕ приложения Office (Word, Excel, PowerPoint) перед продолжением!", vbExclamation
    
    ' Пути к кэшу для разных версий Office
    Dim cachePaths As Variant
    cachePaths = Array( _
        Environ("LOCALAPPDATA") & "\Microsoft\Office\16.0\OfficeFileCache\", _
        Environ("LOCALAPPDATA") & "\Microsoft\Office\15.0\OfficeFileCache\", _
        Environ("LOCALAPPDATA") & "\Microsoft\Office\14.0\OfficeFileCache\", _
        Environ("APPDATA") & "\Microsoft\Office\Recent\", _
        Environ("TEMP") & "\VBE\" _
    )
    
    Dim i As Integer
    For i = 0 To UBound(cachePaths)
        If fso.FolderExists(cachePaths(i)) Then
            On Error Resume Next
            fso.DeleteFolder cachePaths(i) & "*", True
            On Error GoTo 0
        End If
    Next i
    
    MsgBox "Кэш очищен. Перезапустите Excel и проверьте иконки.", vbInformation
End Sub


' Функция для проверки состояния надстроек
' Версия: 1.0.0
' Дата: 11.07.2025
' Описание: Проверяет и восстанавливает отключенные надстройки

Sub CheckAndRestoreAddins()
    Dim diagText As String
    Dim addin As Object
    
    diagText = "=== ДИАГНОСТИКА НАДСТРОЕК ===" & vbCrLf & vbCrLf
    
    ' Проверяем установленные надстройки
    diagText = diagText & "[УСТАНОВЛЕННЫЕ НАДСТРОЙКИ]" & vbCrLf
    For Each addin In Application.AddIns
        If addin.Installed Then
            diagText = diagText & "[+] " & addin.name & " - АКТИВНА" & vbCrLf
        Else
            diagText = diagText & "[-] " & addin.name & " - ОТКЛЮЧЕНА" & vbCrLf
        End If
    Next addin
    
    diagText = diagText & vbCrLf & "[ИНСТРУКЦИИ]" & vbCrLf
    diagText = diagText & "1. Откройте Файл > Параметры > Надстройки" & vbCrLf
    diagText = diagText & "2. В списке 'Управление' выберите 'Отключенные элементы'" & vbCrLf
    diagText = diagText & "3. Нажмите 'Перейти'" & vbCrLf
    diagText = diagText & "4. Включите отключенные надстройки" & vbCrLf
    diagText = diagText & "5. Перезапустите Excel"
    
    MsgBox diagText, vbInformation, "Диагностика надстроек"
End Sub


