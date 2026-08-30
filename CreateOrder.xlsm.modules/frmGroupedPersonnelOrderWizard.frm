VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGroupedPersonnelOrderWizard 
   Caption         =   "Единый кадровый приказ"
   ClientHeight    =   9810.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "frmGroupedPersonnelOrderWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGroupedPersonnelOrderWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = t("personnel.grouped.title", "Единый кадровый приказ")
    lblDescription.Caption = t("personnel.grouped.description", "Выберите сохранённые EventID через запятую. Порядок строк и параграфов будет сохранён.")
    lblSelection.Caption = t("personnel.grouped.selection", "EventID (необязательно)")
    lblReadOnly.Caption = t("personnel.grouped.read_only", "Проверка не изменяет реестры; экспорт блокируется при неполной записи.")
    btnPreview.Caption = t("personnel.grouped.preview", "Проверить")
    btnExport.Caption = t("personnel.grouped.export", "Сформировать DOCX")
    btnClose.Caption = t("common.close", "Закрыть")
    lblSummary.Caption = t("personnel.grouped.not_loaded", "Проверка ещё не запускалась.")
    txtReport.Text = vbNullString
End Sub

Private Sub btnPreview_Click()
    On Error GoTo Failed
    txtReport.Text = mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport(Trim$(txtEventIDs.Text))
    If Left$(txtReport.Text, 3) = "OK|" Then
        lblSummary.Caption = t("personnel.grouped.valid", "Данные готовы к формированию DOCX.")
    Else
        lblSummary.Caption = t("personnel.grouped.invalid", "Есть ошибки. Исправьте строки, отмеченные в отчёте.")
    End If
    Exit Sub
Failed:
    lblSummary.Caption = tf("personnel.grouped.failed", "Проверка не выполнена: {error}", "{error}", Err.description)
    txtReport.Text = vbNullString
End Sub

Private Sub btnExport_Click()
    Dim outputPath As String
    On Error GoTo Failed
    outputPath = mdlGroupedPersonnelOrderExport.ExportGroupedPersonnelOrder(Trim$(txtEventIDs.Text))
    lblSummary.Caption = tf("personnel.grouped.exported", "DOCX сформирован: {path}", "{path}", outputPath)
    txtReport.Text = mdlGroupedPersonnelOrderExport.BuildGroupedPersonnelOrderReport(Trim$(txtEventIDs.Text))
    Exit Sub
Failed:
    lblSummary.Caption = tf("personnel.grouped.failed", "Экспорт не выполнен: {error}", "{error}", Err.description)
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
