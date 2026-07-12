VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLicenseActions
   Caption         =   "Файлы лицензии"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmLicenseActions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLicenseActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = t("form.license_actions.title", "Файлы лицензии")
    lblDescription.Caption = t("form.license_actions.description", "Выберите нужное действие для лицензии.")
    btnExportRequest.Caption = t("form.license_actions.export", "Подготовить запрос")
    btnImportResponse.Caption = t("form.license_actions.import", "Загрузить лицензию")
    btnLicenseStatus.Caption = t("form.license_actions.status", "Состояние лицензии")
    btnClose.Caption = t("form.license_actions.close", "Закрыть")
End Sub

Private Sub btnExportRequest_Click()
    modActivation.ExportActivationRequestUI
End Sub

Private Sub btnImportResponse_Click()
    modActivation.ImportActivationResponseUI
End Sub

Private Sub btnLicenseStatus_Click()
    modActivation.ShowLicenseStatusUI
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
