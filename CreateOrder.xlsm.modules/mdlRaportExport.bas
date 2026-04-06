Attribute VB_Name = "mdlRaportExport"
' ===============================================================================
' Module mdlRaportExport
' Version: 1.7.3 (Unified Raport & Risk Export, Fixed duplicates)
' Date: 25.02.2026
' Description: Export of reports (raports), correct insertion of periodsText via Range for large strings.
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' ===============================================================================

Option Explicit

Sub ExportToWordRaportFromTemplateByLichniyNomer(Optional RaportType As String = "DSO")

    Call mdlHelper.EnsureStaffColumnsInitialized
    
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
    Dim hasAnyPeriods As Boolean

    ' --- ОПРЕДЕЛЯЕМ ШАБЛОН И ПРЕФИКС В ЗАВИСИМОСТИ ОТ ТИПА РАПОРТА ---
    Dim templateName As String
    Dim filePrefix As String
    
    If RaportType = "Risk" Then
        templateName = "Шаблон_РапортРиск.docx"
        filePrefix = "Рапорт_Риск_"
    Else
        templateName = "Шаблон_Рапорт.docx"
        filePrefix = "Рапорт_Отгулы_"
    End If

    templatePath = ThisWorkbook.Path & "\" & templateName
    ' -----------------------------------------------------------------

    If mdlHelper.hasCriticalErrors() Then
        MsgBox "Экспорт рапортов заблокирован из-за критических ошибок в данных!" & vbCrLf & _
               "Исправьте все ошибки (красные ячейки) в листе ДСО.", vbCritical, "Экспорт невозможен"
        Exit Sub
    End If
    
    If Not modActivation.CheckLicenseAndPrompt() Then Exit Sub

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Создание рапортов..."

    If dir(templatePath) = "" Then
        MsgBox "Файл шаблона не найден: " & templatePath, vbCritical
        GoTo CleanUp
    End If

    Set wdApp = CreateObject("Word.Application")
    If wdApp Is Nothing Then
        MsgBox "Не удалось создать отдельный экземпляр Word.", vbCritical, "Ошибка Word"
        GoTo CleanUp
    End If
    wordWasNotRunning = True

    wdApp.Visible = False ' Hide Word for batch processing

    Set wsMain = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")

    If Not mdlHelper.FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast) Then
        MsgBox "Ошибка: Не удалось найти необходимые столбцы в листе 'Штат'.", vbCritical
        GoTo CleanUp
    End If

    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastRowMain = wsMain.Cells(wsMain.Rows.count, "C").End(xlUp).Row

    For i = 2 To lastRowMain
        Application.StatusBar = "Создание рапорта " & (i - 1) & " из " & (lastRowMain - 1)

        currentFIO = Trim(wsMain.Cells(i, 2).value)
        currentLichniyNomer = Trim(wsMain.Cells(i, 3).value)

        If currentLichniyNomer <> "" Then
            Dim staffRow As Long
            staffRow = mdlHelper.FindStaffRow(wsStaff, currentLichniyNomer, colLichniyNomer)
            If staffRow > 0 Then
                Set wdDoc = wdApp.Documents.Open(templatePath)

                ' Get personal data
                lichniyNomer = wsStaff.Cells(staffRow, colLichniyNomer).value
                zvanie = wsStaff.Cells(staffRow, colZvanie).value
                fio = wsStaff.Cells(staffRow, colFIO).value
                dolzhnost = wsStaff.Cells(staffRow, colDolzhnost).value
                VoinskayaChast = mdlHelper.ExtractVoinskayaChast(wsStaff.Cells(staffRow, colVoinskayaChast).value)

                ' Collect all periods for the row
                Set periodList = New Collection
                mdlHelper.CollectAllPersonPeriods wsMain, i, periodList

                ' Check for errors (end < start)
                For j = 1 To periodList.count
                    If periodList(j)(2) < periodList(j)(1) Then
                        MsgBox "Обнаружена ошибка: дата окончания меньше даты начала. Исправьте периоды для " & fio & " (" & lichniyNomer & ")." & vbCrLf & _
                        "Экспорт не будет выполнен!", vbCritical, "Ошибка данных"
                        If Not wdDoc Is Nothing Then wdDoc.Close False
                        GoTo CleanUp
                    End If
                Next j

                ' Convert to array and sort by start date
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

                    periodsText = ""
                    hasAnyPeriods = False
                    firstDate = ""
                    lastDate = ""

                    For j = 1 To UBound(periodArr)
                        hasAnyPeriods = True
                        If firstDate = "" Then firstDate = Format(periodArr(j, 1), "dd.mm.yyyy")
                        lastDate = Format(periodArr(j, 2), "dd.mm.yyyy")
                        
                        ' --- НОВАЯ ЛОГИКА ОТОБРАЖЕНИЯ СУТОК ИЛИ ПРОЦЕНТОВ ---
                        Dim periodValue As String
                        If RaportType = "Risk" Then
                            Dim calcPercent As Long
                            calcPercent = periodArr(j, 3) * 2
                            If calcPercent > 60 Then calcPercent = 60 ' Ограничение максимум 60%
                            periodValue = calcPercent & "%"
                        Else
                            periodValue = periodArr(j, 3) & " сут."
                        End If
                        ' ----------------------------------------------------
                        
                        periodsText = periodsText & "- с " & Format(periodArr(j, 1), "dd.mm.yyyy") & _
                            " по " & Format(periodArr(j, 2), "dd.mm.yyyy") & " (" & periodValue & ")"
                            
                        If periodArr(j, 2) < cutoffDate Then
                            periodsText = periodsText & " (НЕ АКТУАЛЕН — старше 3 лет + 1 месяц!)"
                        End If
                        periodsText = periodsText & vbCrLf
                    Next j
                    
                    ' --- РАЗДЕЛЬНАЯ ЛОГИКА ДЛЯ ОТГУЛОВ И РИСКА ---
                    ' Calculate total days and rest calculation
                    Dim totalDays As Long, restDays As Long, daysList As String
                    Dim periodForRaport As String, calculationText As String
                    totalDays = 0
                    daysList = ""
                    For j = 1 To UBound(periodArr)
                        totalDays = totalDays + periodArr(j, 3)
                        If daysList = "" Then
                            daysList = periodArr(j, 3)
                        Else
                            daysList = daysList & "+" & periodArr(j, 3)
                        End If
                    Next j
                    
                    ' --- РАЗДЕЛЬНАЯ ЛОГИКА ДЛЯ ОТГУЛОВ И РИСКА ---
                    If totalDays > 0 Then
                        If RaportType = "Risk" Then
                            ' Расчет для риска не нужен (оставляем пустым)
                            calculationText = ""
                        Else
                            ' Расчет для ДСО (отгулы)
                            restDays = Int(totalDays / 3) * 2
                            calculationText = "(" & daysList & ") = " & totalDays & " суток привлечения/3*2 = " & restDays & " суток отдыха."
                        End If
                    Else
                        If RaportType = "Risk" Then
                            calculationText = ""
                        Else
                            calculationText = "Нет актуальных периодов для расчета."
                        End If
                    End If
                    ' ---------------------------------------------
                    
                    ' Form participation period string
                    If firstDate <> "" And lastDate <> "" Then
                        periodForRaport = "с " & firstDate & " по " & lastDate
                    Else
                        periodForRaport = "период не указан"
                    End If
                Else
                    periodsText = "Нет актуальных периодов для расчета." & vbCrLf
                    firstDate = "нет"
                    lastDate = "периодов"
                    periodForRaport = "период не указан"
                    calculationText = "Нет актуальных периодов для расчета."
                End If

                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ФИО_ИМЕНИТЕЛЬНЫЙ]", fio)
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ЛИЧНЫЙ_НОМЕР]", lichniyNomer)
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ЗВАНИЕ_СОКРАЩЕННО]", mdlHelper.GetZvanieSkrasheno(zvanie))
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ФИО_ИНИЦИАЛЫ]", mdlHelper.GetFIOWithInitials(fio))
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ДОЛЖНОСТЬ]", mdlHelper.GetDolzhnostImenitelny(dolzhnost, VoinskayaChast))
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ПЕРИОД_УЧАСТИЯ]", periodForRaport)
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[РАСЧЕТ]", calculationText)
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ЗВАНИЕ_ИМЕНИТЕЛЬНЫЙ]", mdlHelper.GetZvanieImenitelnyForSignature(zvanie))
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ФИО_ИНИЦИАЛЫ_ИМЕНИТЕЛЬНЫЙ]", mdlHelper.GetFIOWithInitialsImenitelny(fio))
                Call mdlWordTemplateSafe.ReplacePlaceholderText(wdDoc, "[ПЕРИОДЫ_СЛУЖБЫ]", periodsText)

                ' Generate file name
                Dim cleanFIO As String, periodForFileName As String
                cleanFIO = Replace(Replace(Replace(fio, " ", "_"), ".", ""), ",", "")
                periodForFileName = firstDate & "_по_" & lastDate
                
                ' --- ИСПОЛЬЗУЕМ ПРЕФИКС В НАЗВАНИИ ФАЙЛА ---
                fileName = filePrefix & lichniyNomer & "_" & cleanFIO & "_" & periodForFileName & ".docx"
                savePath = ThisWorkbook.Path & "\" & fileName

                Call mdlHelper.SaveWordDocumentSafe(wdDoc, savePath)
                wdDoc.Close
                Debug.Print "Создан рапорт: " & fileName
            End If
        End If
    Next i

    MsgBox "Все рапорты сформированы и сохранены в папке: " & ThisWorkbook.Path, vbInformation, "Рапорты готовы"
    GoTo CleanUp

ErrorHandler:
    MsgBox "Ошибка на строке " & i & " (ФИО: " & currentFIO & "): " & Err.description, vbCritical, "Ошибка"
    If Not wdDoc Is Nothing Then wdDoc.Close False
CleanUp:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error Resume Next
    If Not wdDoc Is Nothing Then Set wdDoc = Nothing
    If Not wdApp Is Nothing Then
        wdApp.Quit
        Set wdApp = Nothing
    End If

    Set wdDoc = Nothing
    Set wsMain = Nothing
    Set wsStaff = Nothing
End Sub

