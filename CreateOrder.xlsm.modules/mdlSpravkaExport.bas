Attribute VB_Name = "mdlSpravkaExport"
'===============================================================================
' Module mdlSpravkaExport for creating a certificate of presence in the SVO zone based on a template
' Version: 2.2.2 (WordSafe, process control)
'===============================================================================

Option Explicit

'/**
 ' Procedure for creating certificates of presence of a serviceman in the SVO zone with safe Word instance management.
 ' Guarantees correct closing of the document and Word application, removes "hanging" Word processes.
 '*/
Sub ExportToWordSpravkaFromTemplate()
    If mdlHelper.hasCriticalErrors() Then
        MsgBox "Экспорт справок заблокирован из-за критических ошибок в данных!" & vbCrLf & _
               "Исправьте все ошибки (красные ячейки) в листе ДСО.", vbCritical, "Экспорт невозможен"
        Exit Sub
    End If

    Dim wdApp As Object
    Dim wordWasNotRunning As Boolean
    Dim wdDoc As Object
    Dim wsMain As Worksheet
    Dim wsStaff As Worksheet
    Dim lastRowMain As Long
    Dim i As Long, j As Long
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim currentFIO As String
    Dim currentLichniyNomer As String
    Dim fio As String, lichniyNomer As String, zvanie As String, dolzhnost As String, VoinskayaChast As String
    Dim templatePath As String, savePath As String, fileName As String

    Dim periodList As Collection
    Dim periodArr() As Variant
    Dim cutoffDate As Date

    Dim periodsText As String
    Dim firstDate As String, lastDate As String
    Dim hasValidPeriods As Boolean

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Создание справок ДСО..."

    templatePath = ThisWorkbook.Path & "\Шаблон_Справка.docx"
    If dir(templatePath) = "" Then
        MsgBox "Файл шаблона не найден: " & templatePath, vbCritical
        GoTo CleanUp
    End If

    ' === CORRECTLY create Word instance ===
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
        wordWasNotRunning = True
    Else
        wordWasNotRunning = False
    End If
    On Error GoTo 0

    If wdApp Is Nothing Then
        MsgBox "Не удалось создать экземпляр Word. Операция невозможна.", vbCritical
        GoTo CleanUp
    End If

    wdApp.Visible = False

    Set wsMain = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")

    If Not mdlHelper.FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast) Then
        MsgBox "Ошибка: Не удалось найти необходимые столбцы в листе 'Штат'.", vbCritical
        GoTo CleanUp
    End If

    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastRowMain = wsMain.Cells(wsMain.Rows.count, "C").End(xlUp).Row

    For i = 2 To lastRowMain
        Application.StatusBar = "Создание справки " & (i - 1) & " из " & (lastRowMain - 1)
        currentFIO = Trim(wsMain.Cells(i, 2).value)
        currentLichniyNomer = Trim(wsMain.Cells(i, 3).value)
        If currentLichniyNomer <> "" Then
            Dim staffRow As Long
            staffRow = mdlHelper.FindStaffRow(wsStaff, currentLichniyNomer, colLichniyNomer)
            If staffRow > 0 Then
                Set periodList = New Collection
                mdlHelper.CollectAllPersonPeriods wsMain, i, periodList
                For j = 1 To periodList.count
                    If periodList(j)(2) < periodList(j)(1) Then
                        MsgBox "Обнаружена ошибка: дата окончания меньше даты начала. Исправьте периоды для " & currentFIO & " (" & currentLichniyNomer & ")." & vbCrLf & _
                        "Экспорт не будет выполнен!", vbCritical, "Ошибка данных"
                        GoTo CleanUp
                    End If
                Next j

                Set wdDoc = wdApp.Documents.Open(templatePath)
                lichniyNomer = wsStaff.Cells(staffRow, colLichniyNomer).value
                zvanie = wsStaff.Cells(staffRow, colZvanie).value
                fio = wsStaff.Cells(staffRow, colFIO).value
                dolzhnost = wsStaff.Cells(staffRow, colDolzhnost).value
                VoinskayaChast = mdlHelper.ExtractVoinskayaChast(wsStaff.Cells(staffRow, colVoinskayaChast).value)

                periodsText = ""
                hasValidPeriods = False
                firstDate = ""
                lastDate = ""
                If periodList.count > 0 Then
                    ReDim periodArr(1 To periodList.count, 1 To 3)
                    For j = 1 To periodList.count
                        periodArr(j, 1) = periodList(j)(1)
                        periodArr(j, 2) = periodList(j)(2)
                        periodArr(j, 3) = periodList(j)(3)
                    Next j
                    Dim swap As Boolean
                    Do
                        swap = False
                        For j = 1 To UBound(periodArr) - 1
                            If periodArr(j, 1) > periodArr(j + 1, 1) Then
                                Dim t1, t2, t3
                                t1 = periodArr(j, 1)
                                t2 = periodArr(j, 2)
                                t3 = periodArr(j, 3)
                                periodArr(j, 1) = periodArr(j + 1, 1)
                                periodArr(j, 2) = periodArr(j + 1, 2)
                                periodArr(j, 3) = periodArr(j + 1, 3)
                                periodArr(j + 1, 1) = t1
                                periodArr(j + 1, 2) = t2
                                periodArr(j + 1, 3) = t3
                                swap = True
                            End If
                        Next j
                    Loop While swap

                    For j = 1 To UBound(periodArr)
                        hasValidPeriods = True
                        If firstDate = "" Then firstDate = Format(periodArr(j, 1), "dd.mm.yyyy")
                        lastDate = Format(periodArr(j, 2), "dd.mm.yyyy")
                        periodsText = periodsText & "- с " & Format(periodArr(j, 1), "dd.mm.yyyy") & _
                            " по " & Format(periodArr(j, 2), "dd.mm.yyyy")
                        If periodArr(j, 2) < cutoffDate Then
                            periodsText = periodsText & " (НЕ АКТУАЛЕН — старше 3 лет + 1 месяц!)"
                        End If
                        periodsText = periodsText & vbCrLf
                    Next j
                End If

                If Not hasValidPeriods Then
                    periodsText = "Нет актуальных периодов службы в зоне СВО." & vbCrLf
                End If

                With wdDoc.content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting

                    .Text = "[ЗВАНИЕ]"
                    .Replacement.Text = mdlHelper.GetZvanieImenitelny(zvanie)
                    .Execute Replace:=2

                    .Text = "[ФИО]"
                    .Replacement.Text = fio
                    .Execute Replace:=2

                    .Text = "[ЛИЧНЫЙ_НОМЕР]"
                    .Replacement.Text = lichniyNomer
                    .Execute Replace:=2

                    .Text = "[ДОЛЖНОСТЬ]"
                    .Replacement.Text = mdlHelper.GetDolzhnostImenitelny(dolzhnost, VoinskayaChast)
                    .Execute Replace:=2
                End With

                ' === Inserting periodsText via Range.InsertAfter in chunks of 230 characters ===
                Dim rng As Object
                Set rng = wdDoc.content
                With rng.Find
                    .Text = "[ПЕРИОДЫ]"
                    If .Execute Then
                        rng.Select
                        rng.Text = "" ' Clear placeholder
                        Dim partLen As Integer, startPos As Integer, periodChunk As String
                        partLen = 230
                        For startPos = 1 To Len(periodsText) Step partLen
                            periodChunk = Mid(periodsText, startPos, partLen)
                            rng.InsertAfter periodChunk
                        Next startPos
                    End If
                End With

                Dim cleanFIO As String, periodForFileName As String
                cleanFIO = Replace(Replace(Replace(fio, " ", "_"), ".", ""), ",", "")
                If firstDate <> "" And lastDate <> "" Then
                    periodForFileName = firstDate & "_по_" & lastDate
                Else
                    periodForFileName = "нет_актуальных_периодов"
                End If
                fileName = "СправкаДСО_" & lichniyNomer & "_" & cleanFIO & "_" & periodForFileName & ".docx"
                savePath = ThisWorkbook.Path & "\" & fileName

                Call mdlHelper.SaveWordDocumentSafe(wdDoc, savePath)
                wdDoc.Close
                Set wdDoc = Nothing
            End If
        End If
    Next i

    MsgBox "Справки созданы и сохранены в папке: " & ThisWorkbook.Path, vbInformation, "Справки готовы"
    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка при создании справок: " & Err.Description, vbCritical, "Ошибка"
    If Not wdDoc Is Nothing Then
        wdDoc.Close False
        Set wdDoc = Nothing
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ' Additionally: Forced reset of hanging Word instances
    On Error Resume Next
    If Not wdApp Is Nothing Then
        If wordWasNotRunning Then
            wdApp.Quit
        Else
            ' Ensure no open documents
            If wdApp.Documents.count = 0 Then
                wdApp.Quit
            End If
        End If
        Set wdApp = Nothing
    End If

    ' Forced termination of extraneous Word processes via WMI
    Dim wmi As Object, procs As Object, proc As Object
    Set wmi = GetObject("winmgmts:")
    Set procs = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='WINWORD.EXE'")
    For Each proc In procs
        On Error Resume Next
        proc.Terminate
        On Error GoTo 0
    Next proc

    Set wsMain = Nothing
    Set wsStaff = Nothing
End Sub

