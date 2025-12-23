Attribute VB_Name = "mdlMainExport"
' ===============================================================================
' модуль mdlMainExport
' Версия: 5.4.0
'Дата: 30.10.2025
' Описание: Экспорт основного приказа c заполнителями для отсутствующих сотрудников,
' файл сохраняется как "Основной приказ.docx" в папке с макросом и сразу открывается
' ===============================================================================

Option Explicit

'If GetProductStatus() = 2 Then
'    frmAbout.Show
'    Exit Sub
'End If

Sub ExportToWordFromStaffByLichniyNomer()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wsMain As Worksheet
    Dim wsStaff As Worksheet
    Dim lastRowMain As Long
    Dim i As Long, j As Long
    Dim colLichniyNomer As Long, colZvanie As Long, colFIO As Long, colDolzhnost As Long, colVoinskayaChast As Long
    Dim currentFIO As String
    Dim currentLichniyNomer As String
    Dim fio As String, lichniyNomer As String, zvanie As String, dolzhnost As String, VoinskayaChast As String, osnovanie As String
    Dim staffRow As Long
    Dim periodList As Collection
    Dim periodArr() As Variant
    Dim cutoffDate As Date
    Dim textLine As String
    Dim fileName As String, savePath As String
    Dim wdVisibleState As Boolean

    On Error GoTo ErrorHandler
    
    Call mdlHelper.InitStaffColumnIndexes

    Set wsMain = ThisWorkbook.Sheets("ДСО")
    Set wsStaff = ThisWorkbook.Sheets("Штат")

    If Not mdlHelper.FindColumnNumbers(wsStaff, colLichniyNomer, colZvanie, colFIO, colDolzhnost, colVoinskayaChast) Then
        MsgBox "Ошибка: Не удалось найти необходимые столбцы в листе 'Штат'. Проверьте заголовки.", vbCritical
        Exit Sub
    End If

    cutoffDate = mdlHelper.GetExportCutoffDate()
    lastRowMain = wsMain.Cells(wsMain.Rows.count, "C").End(xlUp).Row ' Поиск по столбцу C (Личный номер)

    Set wdApp = CreateObject("Word.Application")
    wdVisibleState = wdApp.Visible
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add

    For i = 2 To lastRowMain
        currentFIO = wsMain.Cells(i, 2).value ' ФИО из столбца B
        currentLichniyNomer = wsMain.Cells(i, 3).value ' Личный номер из столбца C
        osnovanie = wsMain.Cells(i, 4).value ' Основание

        staffRow = mdlHelper.FindStaffRow(wsStaff, currentLichniyNomer, colLichniyNomer)
        If staffRow > 0 Then
            lichniyNomer = wsStaff.Cells(staffRow, colLichniyNomer).value
            zvanie = wsStaff.Cells(staffRow, colZvanie).value
            fio = wsStaff.Cells(staffRow, colFIO).value
            dolzhnost = wsStaff.Cells(staffRow, colDolzhnost).value
            VoinskayaChast = mdlHelper.ExtractVoinskayaChast(wsStaff.Cells(staffRow, colVoinskayaChast).value)
        Else
            lichniyNomer = "Заполните личный номер"
            zvanie = "Заполните воинское звание"
            fio = currentFIO
            dolzhnost = "Заполните воинскую должность"
            VoinskayaChast = "Заполните наименование части"
        End If

        textLine = wsMain.Cells(i, 1).value & ". " & mdlHelper.SklonitZvanie(zvanie) & " " & _
                              mdlHelper.SklonitFIO(fio) & ", личный номер " & lichniyNomer & ", " & _
                              mdlHelper.SklonitDolzhnost(dolzhnost, VoinskayaChast) & vbCrLf

        ' Сбор всех периодов для строки
        Set periodList = New Collection
        mdlHelper.CollectAllPersonPeriods wsMain, i, periodList

        ' Проверка на ошибочные пары (разрешены совпадающие даты — только < запрещено)
        For j = 1 To periodList.count
            If periodList(j)(2) < periodList(j)(1) Then
                MsgBox "Обнаружена ошибка: дата окончания меньше даты начала. Исправьте периоды для " & fio & " (" & lichniyNomer & ")." & vbCrLf & _
                "Экспорт не будет выполнен!", vbCritical, "Ошибка данных"
                wdDoc.Close False
                wdApp.Quit
                Exit Sub
            End If
        Next j

        ' Сортировка и сборка массивов периодов
        If periodList.count > 0 Then
            ReDim periodArr(1 To periodList.count, 1 To 3)
            For j = 1 To periodList.count
                periodArr(j, 1) = periodList(j)(1)
                periodArr(j, 2) = periodList(j)(2)
                periodArr(j, 3) = periodList(j)(3)
            Next j
            ' Сортировка пузырьком
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

            Dim totalDays As Long, daysList As String, restDays As Long
            totalDays = 0: daysList = ""
            For j = 1 To UBound(periodArr)
                totalDays = totalDays + periodArr(j, 3)
                If daysList = "" Then
                    daysList = periodArr(j, 3)
                Else
                    daysList = daysList & "+" & periodArr(j, 3)
                End If

                textLine = textLine & "- с " & Format(periodArr(j, 1), "dd.mm.yyyy") & " по " & _
                                Format(periodArr(j, 2), "dd.mm.yyyy") & " в количестве " & periodArr(j, 3) & " суток"
                If periodArr(j, 2) < cutoffDate Then
                    textLine = textLine & " (НЕ АКТУАЛЕН — старше 3 лет + 1 месяц!)"
                End If
                textLine = textLine & vbCrLf
            Next j

            If totalDays > 0 Then
                restDays = Int(totalDays / 3) * 2
                textLine = textLine & "(" & daysList & ") = " & totalDays & " суток привлечения /3*2 = " & restDays & " суток отдыха" & vbCrLf
            End If
        Else
            textLine = textLine & "Нет периодов для вывода." & vbCrLf
        End If

        If osnovanie <> "" Then
            textLine = textLine & "Основание: " & osnovanie & vbCrLf
        End If
        textLine = textLine & vbCrLf
        wdDoc.Content.InsertAfter textLine
    Next i

    fileName = "Основной приказ.docx"
    savePath = ThisWorkbook.Path & "\" & fileName
    Call mdlHelper.SaveWordDocumentSafe(wdDoc, savePath)
    wdDoc.Activate
    wdApp.Visible = True

    MsgBox "Основной приказ успешно создан и открыт для просмотра: " & savePath, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка экспорта: " & Err.description, vbCritical, "Ошибка"
    If Not wdDoc Is Nothing Then wdDoc.Close False

End Sub


