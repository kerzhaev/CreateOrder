Attribute VB_Name = "VbaModuleManager"
Option Explicit
Option Private Module

#Const MANAGING_WORD = 0
#Const MANAGING_EXCEL = 1

Private Const MY_NAME As String = "VbaModuleManager"
Private Const ERR_SUPPORTED_APPS As String = MY_NAME & " currently only supports Microsoft Word and Excel."
Private Const DOCUMENT_MODULES_DIR As String = "workbook-modules"
Private Const DOCUMENT_FILE_EXT As String = "bas"

' === КОНСТАНТЫ ДЛЯ ПОЗДНЕГО СВЯЗЫВАНИЯ (VBIDE) ===
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_ClassModule As Long = 2
Private Const vbext_ct_MSForm As Long = 3
Private Const vbext_ct_Document As Long = 100

Dim allComponents As Object
Dim fileSys As Object
Dim alreadySaved As Boolean

Public Sub ImportModules(FromDirectory As String, Optional ShowMsgBox As Boolean = True)
    On Error GoTo ErrorHandler

    Set fileSys = CreateObject("Scripting.FileSystemObject")

    Dim fromPath As String
    Dim basePath As String
    Dim imports As Object
    Dim numFiles As Long

    fromPath = resolveFolderPath(FromDirectory, "импорта")
    basePath = getFilePath()
    Set imports = CreateObject("Scripting.Dictionary")

    numFiles = importCodeFilesFromFolder(fromPath, imports)

#If MANAGING_EXCEL Then
    Dim documentPath As String
    documentPath = fileSys.BuildPath(basePath, DOCUMENT_MODULES_DIR)
    numFiles = numFiles + importDocumentFilesFromFolder(documentPath, imports)
#End If

    If ShowMsgBox Then
        showImportExportSummary "Импорт модулей", "импортировано", fromPath, imports, numFiles
    End If

    Set imports = Nothing
    Set fileSys = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при импорте модулей: " & Err.Description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка импорта"
    Set fileSys = Nothing
End Sub

Public Sub ExportModules(ToDirectory As String)
    On Error GoTo ErrorHandler

    Set fileSys = CreateObject("Scripting.FileSystemObject")

    Dim toPath As String
    Dim exports As Object
    Dim numFiles As Long

    toPath = ensureFolderPath(ToDirectory)
    Set exports = CreateObject("Scripting.Dictionary")

    numFiles = exportCodeComponentsToFolder(toPath, exports)

#If MANAGING_EXCEL Then
    Dim documentPath As String
    documentPath = ensureFolderPath(DOCUMENT_MODULES_DIR)
    numFiles = numFiles + exportDocumentComponentsToFolder(documentPath, exports)
#End If

    showImportExportSummary "Экспорт модулей", "экспортировано", toPath, exports, numFiles

    Set exports = Nothing
    Set fileSys = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при экспорте модулей: " & Err.Description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка экспорта"
    Set fileSys = Nothing
End Sub

Public Sub RemoveModules(Optional ShowMsgBox As Boolean = True)
    On Error GoTo ErrorHandler

    If alreadySaved Then
        alreadySaved = False
        Exit Sub
    End If

    Dim removals As New Collection
    Dim vbc As Object
    Dim numModules As Long
    Dim currentComponents As Object

    Set currentComponents = getAllComponents()

    For Each vbc In currentComponents
        If isRemovableComponent(vbc) Then
            numModules = numModules + 1
            removals.Add vbc.Name
            currentComponents.Remove vbc
        End If
    Next vbc

    alreadySaved = True
    saveFile

    If ShowMsgBox Then
        Dim item As Variant
        Dim msg As String

        msg = numModules & " модулей успешно удалено:" & vbCrLf & vbCrLf
        For Each item In removals
            msg = msg & "    " & item & vbCrLf
        Next item
        msg = msg & vbCrLf & "Document-модули Excel не удаляются этой процедурой." _
                  & vbCrLf & "VbaModuleManager никогда не будет переимпортирован или экспортирован автоматически." _
                  & vbCrLf & "НИКОГДА не редактируйте код в VBE и отдельном редакторе одновременно!"
        MsgBox msg, vbOKOnly, "Удаление модулей"
    End If

    Exit Sub

ErrorHandler:
    alreadySaved = False
    MsgBox "Ошибка при удалении модулей: " & Err.Description & vbCrLf & "Код ошибки: " & Err.Number, vbCritical, "Ошибка удаления"
End Sub

Private Function importCodeFilesFromFolder(ByVal fromPath As String, ByRef imports As Object) As Long
    Dim dir As Object
    Dim f As Object
    Dim dotIndex As Long
    Dim ext As String
    Dim replaced As Boolean

    If Not fileSys.FolderExists(fromPath) Then Exit Function

    Set dir = fileSys.GetFolder(fromPath)

    For Each f In dir.Files
        dotIndex = InStrRev(f.Name, ".")
        If dotIndex > 0 Then
            ext = UCase$(Right$(f.Name, Len(f.Name) - dotIndex))
            If (ext = "BAS" Or ext = "CLS" Or ext = "FRM") And Left$(f.Name, dotIndex - 1) <> MY_NAME Then
                replaced = doImport(f)
                addOperationResult imports, "[Code] " & f.Name, replaced
                importCodeFilesFromFolder = importCodeFilesFromFolder + 1
            End If
        End If
    Next f
End Function

Private Function exportCodeComponentsToFolder(ByVal toPath As String, ByRef exports As Object) As Long
    Dim vbc As Object
    Dim currentComponents As Object
    Dim replaced As Boolean

    Set currentComponents = getAllComponents()

    For Each vbc In currentComponents
        If isExportableCodeComponent(vbc) Then
            replaced = doExport(vbc, toPath)
            addOperationResult exports, "[Code] " & vbc.Name, replaced
            exportCodeComponentsToFolder = exportCodeComponentsToFolder + 1
        End If
    Next vbc
End Function

Private Function importDocumentFilesFromFolder(ByVal fromPath As String, ByRef imports As Object) As Long
#If MANAGING_EXCEL Then
    On Error GoTo ErrorHandler

    Dim dir As Object
    Dim f As Object
    Dim targetComponent As Object
    Dim replaced As Boolean

    If Not fileSys.FolderExists(fromPath) Then Exit Function

    Set dir = fileSys.GetFolder(fromPath)
    For Each f In dir.Files
        If UCase$(fileSys.GetExtensionName(f.Name)) = UCase$(DOCUMENT_FILE_EXT) Then
            Set targetComponent = getDocumentComponentForFile(f.Name)
            If Not targetComponent Is Nothing Then
                replaced = importDocumentModule(targetComponent, f.Path)
                addOperationResult imports, "[Doc] " & f.Name, replaced
                importDocumentFilesFromFolder = importDocumentFilesFromFolder + 1
            End If
        End If
    Next f
#End If
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при импорте document-модулей: " & Err.Description, vbCritical, "Ошибка импорта"
End Function

Private Function exportDocumentComponentsToFolder(ByVal toPath As String, ByRef exports As Object) As Long
#If MANAGING_EXCEL Then
    On Error GoTo ErrorHandler

    Dim vbc As Object
    Dim currentComponents As Object
    Dim replaced As Boolean

    Set currentComponents = getAllComponents()

    For Each vbc In currentComponents
        If vbc.Type = vbext_ct_Document Then
            replaced = exportDocumentModule(vbc, toPath)
            addOperationResult exports, "[Doc] " & getDocumentModuleFileName(vbc), replaced
            exportDocumentComponentsToFolder = exportDocumentComponentsToFolder + 1
        End If
    Next vbc
#End If
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при экспорте document-модулей: " & Err.Description, vbCritical, "Ошибка экспорта"
End Function

Private Function resolveFolderPath(ByVal rawPath As String, ByVal operationName As String) As String
    Dim basePath As String
    Dim errMsg As String

    basePath = getFilePath()
    resolveFolderPath = rawPath

    If basePath = "" Then
        MsgBox "Ошибка: Файл должен быть сохранён перед операцией с модулями!", vbCritical, "Ошибка"
        Exit Function
    End If

    If Not fileSys.FolderExists(resolveFolderPath) Then
        resolveFolderPath = fileSys.BuildPath(basePath, rawPath)
    End If

    If Not fileSys.FolderExists(resolveFolderPath) Then
        errMsg = "Не удалось найти директорию " & operationName & ": " & rawPath & vbCrLf & vbCrLf
        errMsg = errMsg & "Проверенные пути:" & vbCrLf
        errMsg = errMsg & "1. " & rawPath & vbCrLf
        errMsg = errMsg & "2. " & resolveFolderPath
        MsgBox errMsg, vbCritical, "Ошибка"
        resolveFolderPath = ""
    End If
End Function

Private Function ensureFolderPath(ByVal rawPath As String) As String
    Dim basePath As String

    basePath = getFilePath()
    ensureFolderPath = rawPath

    If basePath = "" Then
        MsgBox "Ошибка: Файл должен быть сохранён перед экспортом модулей!", vbCritical, "Ошибка экспорта"
        ensureFolderPath = ""
        Exit Function
    End If

    If Not fileSys.FolderExists(ensureFolderPath) Then
        ensureFolderPath = fileSys.BuildPath(basePath, rawPath)
    End If

    If Not fileSys.FolderExists(ensureFolderPath) Then
        fileSys.CreateFolder ensureFolderPath
    End If
End Function

Private Sub showImportExportSummary(ByVal title As String, _
                                    ByVal actionVerb As String, _
                                    ByVal basePath As String, _
                                    ByRef items As Object, _
                                    ByVal numFiles As Long)
    Dim i As Long
    Dim msg As String
    Dim keys As Variant
    Dim values As Variant

    If numFiles = 0 Then
        MsgBox "Файлы не найдены: " & basePath, vbInformation, title
        Exit Sub
    End If

    msg = numFiles & " элементов " & actionVerb & ":" & vbCrLf & vbCrLf
    keys = items.Keys
    values = items.Items

    For i = 0 To items.Count - 1
        msg = msg & "    " & keys(i) & IIf(values(i), " (заменён)", " (новый)") & vbCrLf
    Next i

    MsgBox msg, vbOKOnly, title
End Sub

Private Sub addOperationResult(ByRef items As Object, ByVal key As String, ByVal replaced As Boolean)
    Dim resultKey As String

    resultKey = key
    Do While items.Exists(resultKey)
        resultKey = resultKey & " *"
    Loop
    items.Add resultKey, replaced
End Sub

Private Function getFilePath() As String
    #If MANAGING_WORD Then
        getFilePath = ThisDocument.Path
    #ElseIf MANAGING_EXCEL Then
        getFilePath = ThisWorkbook.Path
    #Else
        raiseUnsupportedAppError
    #End If
End Function

Private Function getAllComponents() As Object
    #If MANAGING_WORD Then
        Set getAllComponents = ThisDocument.VBProject.VBComponents
    #ElseIf MANAGING_EXCEL Then
        Set getAllComponents = ThisWorkbook.VBProject.VBComponents
    #Else
        raiseUnsupportedAppError
    #End If
End Function

Private Sub saveFile()
    #If MANAGING_WORD Then
        ThisDocument.Save
    #ElseIf MANAGING_EXCEL Then
        ThisWorkbook.Save
    #Else
        raiseUnsupportedAppError
    #End If
End Sub

Private Sub raiseUnsupportedAppError()
    Err.Raise Number:=vbObjectError + 1, Description:=ERR_SUPPORTED_APPS
End Sub

Private Function isExportableCodeComponent(ByRef component As Object) As Boolean
    isExportableCodeComponent = (component.Type = vbext_ct_StdModule Or _
                                 component.Type = vbext_ct_ClassModule Or _
                                 component.Type = vbext_ct_MSForm) And _
                                component.Name <> MY_NAME
End Function

Private Function isRemovableComponent(ByRef component As Object) As Boolean
    isRemovableComponent = (component.Type = vbext_ct_StdModule Or _
                            component.Type = vbext_ct_ClassModule Or _
                            component.Type = vbext_ct_MSForm) And _
                           component.Name <> MY_NAME
End Function

Private Function doImport(ByRef codeFile As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim componentName As String
    Dim currentComponents As Object
    Dim m As Object
    Dim alreadyExists As Boolean

    componentName = Left$(codeFile.Name, Len(codeFile.Name) - 4)
    Set currentComponents = getAllComponents()

    On Error Resume Next
    Set m = currentComponents.Item(componentName)
    If Err.Number <> 0 Then Set m = Nothing
    On Error GoTo ErrorHandler

    alreadyExists = Not (m Is Nothing)
    If alreadyExists Then
        currentComponents.Remove m
    End If

    currentComponents.Import codeFile.Path
    doImport = alreadyExists
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при импорте файла " & codeFile.Name & ":" & vbCrLf & Err.Description, vbCritical, "Ошибка импорта"
    doImport = False
End Function

Private Function doExport(ByRef module As Object, ByVal dirPath As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ext As String
    Dim filePath As String
    Dim alreadyExists As Boolean
    Dim f As Object

    Select Case module.Type
        Case vbext_ct_MSForm
            ext = "frm"
        Case vbext_ct_ClassModule
            ext = "cls"
        Case Else
            ext = "bas"
    End Select

    filePath = fileSys.BuildPath(dirPath, module.Name & "." & ext)
    alreadyExists = fileSys.FileExists(filePath)

    If alreadyExists Then
        Set f = fileSys.GetFile(filePath)
        If (f.Attributes And 1) Then
            f.Attributes = f.Attributes - 1
        End If
        fileSys.DeleteFile filePath
    End If

    module.Export filePath
    doExport = alreadyExists
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при экспорте модуля " & module.Name & ":" & vbCrLf & Err.Description, vbCritical, "Ошибка экспорта"
    doExport = False
End Function

Private Function exportDocumentModule(ByRef component As Object, ByVal dirPath As String) As Boolean
#If MANAGING_EXCEL Then
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim textStream As Object

    filePath = fileSys.BuildPath(dirPath, getDocumentModuleFileName(component))
    exportDocumentModule = fileSys.FileExists(filePath)

    Set textStream = fileSys.CreateTextFile(filePath, True, False)
    If component.CodeModule.CountOfLines > 0 Then
        textStream.Write component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
    End If
    textStream.Close
#End If
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при экспорте document-модуля " & component.Name & ":" & vbCrLf & Err.Description, vbCritical, "Ошибка экспорта"
    exportDocumentModule = False
End Function

Private Function importDocumentModule(ByRef component As Object, ByVal filePath As String) As Boolean
#If MANAGING_EXCEL Then
    On Error GoTo ErrorHandler

    Dim textStream As Object
    Dim codeText As String

    importDocumentModule = (component.CodeModule.CountOfLines > 0)

    Set textStream = fileSys.OpenTextFile(filePath, 1, False)
    codeText = textStream.ReadAll
    textStream.Close

    codeText = normalizeDocumentCode(codeText)

    If component.CodeModule.CountOfLines > 0 Then
        component.CodeModule.DeleteLines 1, component.CodeModule.CountOfLines
    End If

    If codeText <> "" Then
        component.CodeModule.AddFromString codeText
    End If
#End If
    Exit Function

ErrorHandler:
    MsgBox "Ошибка при импорте document-модуля " & component.Name & ":" & vbCrLf & Err.Description, vbCritical, "Ошибка импорта"
    importDocumentModule = False
End Function

Private Function getDocumentModuleFileName(ByRef component As Object) As String
#If MANAGING_EXCEL Then
    Dim displayName As String

    If component.Name = "ЭтаКнига" Or component.Name = "ThisWorkbook" Then
        getDocumentModuleFileName = component.Name & "." & DOCUMENT_FILE_EXT
        Exit Function
    End If

    displayName = getDocumentDisplayName(component)
    If displayName <> "" Then
        getDocumentModuleFileName = component.Name & " (" & sanitizeFileName(displayName) & ")." & DOCUMENT_FILE_EXT
    Else
        getDocumentModuleFileName = component.Name & "." & DOCUMENT_FILE_EXT
    End If
#End If
End Function

Private Function getDocumentDisplayName(ByRef component As Object) As String
#If MANAGING_EXCEL Then
    On Error Resume Next
    getDocumentDisplayName = CStr(component.Properties("Name").Value)
    On Error GoTo 0
#End If
End Function

Private Function getDocumentComponentForFile(ByVal fileName As String) As Object
#If MANAGING_EXCEL Then
    Dim baseName As String
    Dim targetName As String
    Dim splitPos As Long
    Dim currentComponents As Object

    baseName = fileSys.GetBaseName(fileName)
    targetName = baseName

    splitPos = InStr(baseName, " (")
    If splitPos > 0 Then
        targetName = Left$(baseName, splitPos - 1)
    End If

    Set currentComponents = getAllComponents()

    On Error Resume Next
    Set getDocumentComponentForFile = currentComponents.Item(targetName)
    If Err.Number <> 0 Then Set getDocumentComponentForFile = Nothing
    On Error GoTo 0

    If Not getDocumentComponentForFile Is Nothing Then
        If getDocumentComponentForFile.Type <> vbext_ct_Document Then
            Set getDocumentComponentForFile = Nothing
        End If
    End If
#End If
End Function

Private Function normalizeDocumentCode(ByVal rawCode As String) As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim currentLine As String

    If rawCode = "" Then Exit Function

    rawCode = Replace(rawCode, vbCrLf, vbLf)
    rawCode = Replace(rawCode, vbCr, vbLf)
    lines = Split(rawCode, vbLf)

    For i = LBound(lines) To UBound(lines)
        currentLine = lines(i)
        If Left$(currentLine, 18) = "Attribute VB_Name" Then
            ' Пропускаем экспортные атрибуты, они недопустимы в document-модулях.
        Else
            If result = "" Then
                result = currentLine
            Else
                result = result & vbCrLf & currentLine
            End If
        End If
    Next i

    Do While Left$(result, 2) = vbCrLf
        result = Mid$(result, 3)
    Loop

    normalizeDocumentCode = result
End Function

Private Function sanitizeFileName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    sanitizeFileName = rawName
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For Each item In invalidChars
        sanitizeFileName = Replace(sanitizeFileName, CStr(item), "_")
    Next item

    sanitizeFileName = Trim$(sanitizeFileName)
    If sanitizeFileName = "" Then sanitizeFileName = "Unnamed"
End Function

' =============================================
' ПУБЛИЧНЫЕ ПРОЦЕДУРЫ ДЛЯ ВЫЗОВА ИЗ МАКРОСОВ EXCEL
' =============================================

Public Sub ИмпортМодулей()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    folderPath = ThisWorkbook.Name & ".modules"

    ImportModules folderPath, ShowMsgBox:=True
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при импорте модулей: " & Err.Description, vbCritical, "Ошибка"
End Sub

Public Sub ЭкспортМодулей()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    folderPath = ThisWorkbook.Name & ".modules"

    ExportModules folderPath
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при экспорте модулей: " & Err.Description, vbCritical, "Ошибка"
End Sub

Public Sub УдалениеМодулей()
    On Error GoTo ErrorHandler

    Dim response As VbMsgBoxResult
    response = MsgBox("Вы уверены, что хотите удалить все модули?" & vbCrLf & _
                      "VbaModuleManager и document-модули будут сохранены.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, _
                      "Подтверждение удаления")

    If response = vbYes Then
        RemoveModules ShowMsgBox:=True
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при удалении модулей: " & Err.Description, vbCritical, "Ошибка"
End Sub
