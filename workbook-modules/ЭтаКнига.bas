Option Explicit

Private Sub Workbook_Open()
    ' Автоматическая проверка структуры файла при каждом запуске.
    Call mdlDataValidation.SilentCheckStructure

    ' После окончания встроенного trial или истечения кода сразу показываем активацию.
    Call modActivation.EnsureLicenseOnOpen
End Sub
