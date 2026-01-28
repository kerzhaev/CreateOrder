Attribute VB_Name = "MdlBackup"
' === СНАПШОТ ВЕРСИИ === 12.07.2025 20:45 ===
' Рабочая версия сохранена: 12.07.2025 20:45

Option Explicit

' Модуль MdlBackup для создания резервных копий VBA проекта
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Экспортирует все модули VBA в отдельные файлы для создания снапшота проекта
' Функциональность: Создает резервную копию всех модулей, форм и классов в указанную папку

Sub CreateVBASnapshot()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    Dim timeStamp As String
    Dim fso As Object
    
    ' Создаем временную метку для уникальности
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    
    ' Создаем папку для снапшота
    exportPath = ThisWorkbook.Path & "\VBA_Snapshots\Snapshot_" & timeStamp & "\"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder Left(exportPath, Len(exportPath) - 1)
    End If
    
    ' Экспортируем все компоненты VBA проекта
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1 ' vbext_ct_StdModule - стандартные модули
                fileName = vbComp.Name & ".bas"
            Case 2 ' vbext_ct_ClassModule - модули классов
                fileName = vbComp.Name & ".cls"
            Case 3 ' vbext_ct_MSForm - пользовательские формы
                fileName = vbComp.Name & ".frm"
            Case 100 ' vbext_ct_Document - модули листов/книги
                fileName = vbComp.Name & ".cls"
        End Select
        
        ' Экспортируем компонент
        vbComp.Export exportPath & fileName
        Debug.Print "Экспортирован: " & fileName
    Next vbComp
    
    ' Создаем файл с информацией о снапшоте
    Dim infoFile As String
    infoFile = exportPath & "SnapshotInfo.txt"
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open infoFile For Output As fileNum
    Print #fileNum, "VBA Project Snapshot"
    Print #fileNum, "Дата создания: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
    Print #fileNum, "Файл проекта: " & ThisWorkbook.Name
    Print #fileNum, "Путь: " & ThisWorkbook.FullName
    Print #fileNum, "Количество модулей: " & ThisWorkbook.VBProject.VBComponents.count
    Print #fileNum, ""
    Print #fileNum, "Список модулей:"
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Print #fileNum, "- " & vbComp.Name & " (Тип: " & GetComponentTypeName(vbComp.Type) & ")"
    Next vbComp
    Close fileNum
    
    MsgBox "Снапшот VBA проекта создан!" & vbCrLf & _
           "Путь: " & exportPath & vbCrLf & _
           "Экспортировано модулей: " & ThisWorkbook.VBProject.VBComponents.count, _
           vbInformation, "Снапшот создан"
    
    ' Открываем папку со снапшотом
    Shell "explorer.exe " & exportPath, vbNormalFocus
End Sub

' Вспомогательная функция для получения названия типа компонента
Function GetComponentTypeName(componentType As Integer) As String
    Select Case componentType
        Case 1: GetComponentTypeName = "Стандартный модуль"
        Case 2: GetComponentTypeName = "Модуль класса"
        Case 3: GetComponentTypeName = "Пользовательская форма"
        Case 100: GetComponentTypeName = "Модуль документа"
        Case Else: GetComponentTypeName = "Неизвестный тип"
    End Select
End Function


' Модуль для восстановления VBA проекта из снапшота
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Импортирует модули VBA из папки снапшота для восстановления предыдущего состояния
' Функциональность: Удаляет текущие модули и загружает модули из выбранной резервной копии

Sub RestoreFromSnapshot()
    Dim importPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbComp As Object
    Dim componentName As String
    
    ' Выбираем папку со снапшотом
    importPath = SelectSnapshotFolder()
    If importPath = "" Then Exit Sub
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(importPath)
    
    ' Предупреждение пользователя
    Dim response As VbMsgBoxResult
    response = MsgBox("ВНИМАНИЕ!" & vbCrLf & _
                     "Это действие удалит все текущие модули VBA и заменит их модулями из снапшота." & vbCrLf & _
                     "Вы уверены, что хотите продолжить?", _
                     vbYesNo + vbExclamation, "Подтверждение восстановления")
    
    If response = vbNo Then Exit Sub
    
    ' Удаляем все пользовательские модули (кроме модулей листов)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then ' Стандартные модули, классы, формы
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
        End If
    Next vbComp
    
    ' Импортируем модули из снапшота
    For Each file In folder.Files
        If Right(LCase(file.Name), 4) = ".bas" Or _
           Right(LCase(file.Name), 4) = ".cls" Or _
           Right(LCase(file.Name), 4) = ".frm" Then
            
            ThisWorkbook.VBProject.VBComponents.Import file.Path
            Debug.Print "Импортирован: " & file.Name
        End If
    Next file
    
    MsgBox "Снапшот успешно восстановлен!" & vbCrLf & _
           "Папка снапшота: " & importPath, _
           vbInformation, "Восстановление завершено"
End Sub

' Функция для выбора папки со снапшотом
Function SelectSnapshotFolder() As String
    Dim snapshotsPath As String
    Dim selectedPath As String
    
    snapshotsPath = ThisWorkbook.Path & "\VBA_Snapshots\"
    
    ' Проверяем существование папки со снапшотами
    If dir(snapshotsPath, vbDirectory) = "" Then
        MsgBox "Папка со снапшотами не найдена: " & snapshotsPath, vbExclamation
        SelectSnapshotFolder = ""
        Exit Function
    End If
    
    ' Здесь можно добавить диалог выбора папки
    ' Для простоты используем InputBox
    selectedPath = InputBox("Введите имя папки снапшота:" & vbCrLf & _
                           "Доступные снапшоты находятся в: " & snapshotsPath, _
                           "Выбор снапшота", "Snapshot_")
    
    If selectedPath <> "" Then
        SelectSnapshotFolder = snapshotsPath & selectedPath & "\"
    Else
        SelectSnapshotFolder = ""
    End If
End Function


' Модуль для управления версиями кода через комментарии
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Добавляет метки версий в код для отслеживания изменений
' Функциональность: Вставляет комментарии с версиями и датами в начало каждого модуля

Sub AddVersionTagsToAllModules()
    Dim vbComp As Object
    Dim codeModule As Object
    Dim versionTag As String
    Dim currentDate As String
    
    currentDate = Format(Now, "dd.mm.yyyy hh:mm")
    versionTag = "' === СНАПШОТ ВЕРСИИ === " & currentDate & " ==="
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 1 Then ' Только стандартные модули
            Set codeModule = vbComp.codeModule
            
            ' Добавляем метку версии в начало модуля
            codeModule.InsertLines 1, versionTag
            codeModule.InsertLines 2, "' Рабочая версия сохранена: " & currentDate
            codeModule.InsertLines 3, ""
            
            Debug.Print "Метка версии добавлена в модуль: " & vbComp.Name
        End If
    Next vbComp
    
    MsgBox "Метки версий добавлены во все модули!", vbInformation
End Sub


' Модуль для создания именованных копий Excel файла
' Версия: 1.0.0
' Дата: 12.07.2025
' Описание: Создает копии текущего Excel файла с временными метками для быстрого отката
' Функциональность: Сохраняет копию файла с описательным именем и временной меткой

Sub CreateWorkbookSnapshot()
    Dim originalPath As String
    Dim snapshotPath As String
    Dim timeStamp As String
    Dim Description As String
    Dim fileName As String
    
    ' Получаем описание от пользователя
    Description = InputBox("Введите краткое описание этого снапшота:", _
                          "Описание снапшота", "Рабочая_версия")
    
    If Description = "" Then Exit Sub
    
    ' Создаем временную метку
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm")
    
    ' Формируем имя файла
    fileName = Replace(ThisWorkbook.Name, ".xlsm", "") & "_" & Description & "_" & timeStamp & ".xlsm"
    
    ' Определяем путь для сохранения
    originalPath = ThisWorkbook.Path
    snapshotPath = originalPath & "\Snapshots\"
    
    ' Создаем папку, если её нет
    If dir(snapshotPath, vbDirectory) = "" Then
        MkDir snapshotPath
    End If
    
    ' Сохраняем копию
    ThisWorkbook.SaveCopyAs snapshotPath & fileName
    
    MsgBox "Снапшот создан!" & vbCrLf & _
           "Файл: " & fileName & vbCrLf & _
           "Путь: " & snapshotPath, _
           vbInformation, "Снапшот сохранен"
    
    ' Открываем папку со снапшотами
    Shell "explorer.exe " & snapshotPath, vbNormalFocus
End Sub


' Быстрая команда для создания снапшота
Sub QuickSnapshot()
    Call AddVersionTagsToAllModules
    Call CreateVBASnapshot
    Call CreateWorkbookSnapshot
End Sub


' =============================================
' @author Кержаев Евгений, ФКУ "95 ФЭС" МО РФ
' @description Удаление дубликатов модулей с именами, заканчивающимися на цифру (например, mdlHelper1)
' @description Используется перед импортом модулей через VbaModuleManager для предотвращения конфликтов имен
' =============================================
Sub RemoveDuplicateModules()
    Dim vbComp As Object
    Dim compName As String
    Dim baseName As String
    Dim lastChar As String
    Dim i As Long
    Dim modulesToRemove As Collection
    Dim compToRemove As Object
    
    On Error GoTo ErrorHandler
    
    Set modulesToRemove = New Collection
    
    ' Проходим по всем компонентам проекта
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ' Проверяем только стандартные модули, классы и формы (не модули листов)
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then
            compName = vbComp.Name
            lastChar = Right(compName, 1)
            
            ' Проверяем, заканчивается ли имя на цифру
            If IsNumeric(lastChar) Then
                ' Получаем базовое имя (без последней цифры)
                baseName = Left(compName, Len(compName) - 1)
                
                ' Проверяем, существует ли модуль с базовым именем
                On Error Resume Next
                Set compToRemove = Nothing
                Set compToRemove = ThisWorkbook.VBProject.VBComponents(baseName)
                On Error GoTo ErrorHandler
                
                ' Если базовый модуль существует, добавляем дубликат в список на удаление
                If Not compToRemove Is Nothing Then
                    modulesToRemove.Add vbComp
                    Debug.Print "Найден дубликат: " & compName & " (базовый модуль: " & baseName & ")"
                End If
            End If
        End If
    Next vbComp
    
    ' Удаляем найденные дубликаты
    If modulesToRemove.count > 0 Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Найдено дубликатов модулей: " & modulesToRemove.count & vbCrLf & _
                         "Удалить дубликаты?", _
                         vbYesNo + vbQuestion, "Удаление дубликатов")
        
        If response = vbYes Then
            For i = 1 To modulesToRemove.count
                Set vbComp = modulesToRemove(i)
                ThisWorkbook.VBProject.VBComponents.Remove vbComp
                Debug.Print "Удален дубликат: " & vbComp.Name
            Next i
            
            MsgBox "Удалено дубликатов: " & modulesToRemove.count, vbInformation, "Готово"
        End If
    Else
        MsgBox "Дубликаты модулей не найдены.", vbInformation, "Проверка завершена"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при удалении дубликатов: " & Err.Description, vbCritical, "Ошибка"
End Sub
