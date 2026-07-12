Attribute VB_Name = "mdlHelper"
' ==============================================================================
' Module: mdlHelper
' Author: Kerzhaev Evgeniy, FKU "95 FES" MO RF
' Version: 1.6.0 (UX Update: Smart Column Detection & Detailed Error Messages)
' Description: Universal utility functions, Smart Position Parser & FIO Engine.
' ==============================================================================

Option Explicit

Public colFIO_Global As Long
Public colLichniyNomer_Global As Long
Public colZvanie_Global As Long
Public colDolzhnost_Global As Long
Public colVoinskayaChast_Global As Long
Private Const FIO_CASE_DATIVE As String = "D"
Private Const FIO_CASE_GENITIVE As String = "R"
Private Const FIO_CASE_NOMINATIVE As String = "N"
Private Const FIO_GENDER_UNKNOWN As Long = 0
Private Const FIO_GENDER_MALE As Long = 1
Private Const FIO_GENDER_FEMALE As Long = 2

Private Function StaffHeaderPersonalNumber() As String
    StaffHeaderPersonalNumber = Ru(1083, 1080, 1095, 1085, 1099, 1081, 32, 1085, 1086, 1084, 1077, 1088)
End Function

Private Function StaffHeaderRank() As String
    StaffHeaderRank = Ru(1074, 1086, 1080, 1085, 1089, 1082, 1086, 1077, 32, 1079, 1074, 1072, 1085, 1080, 1077)
End Function

Private Function StaffHeaderShortRank() As String
    StaffHeaderShortRank = Ru(1079, 1074, 1072, 1085, 1080, 1077)
End Function

Private Function StaffHeaderPosition() As String
    StaffHeaderPosition = Ru(1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100)
End Function

Private Function StaffHeaderStaffPosition() As String
    StaffHeaderStaffPosition = Ru(1096, 1090, 1072, 1090, 1085, 1072, 1103, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100)
End Function

Private Function StaffHeaderFio() As String
    StaffHeaderFio = Ru(1083, 1080, 1094, 1086)
End Function

Private Function StaffHeaderUnit() As String
    StaffHeaderUnit = Ru(1095, 1072, 1089, 1090, 1100)
End Function

Private Function StaffHeaderMilitaryUnit() As String
    StaffHeaderMilitaryUnit = Ru(1074, 1086, 1080, 1085, 1089, 1082, 1072, 1103, 32, 1095, 1072, 1089, 1090, 1100)
End Function

Private Function StaffHeaderPersonnelSection() As String
    StaffHeaderPersonnelSection = Ru(1088, 1072, 1079, 1076, 1077, 1083, 32, 1087, 1077, 1088, 1089, 1086, 1085, 1072, 1083, 1072)
End Function

Private Function StaffHeaderArmyUnitAlt() As String
    StaffHeaderArmyUnitAlt = Ru(1074, 1086, 1081, 1089, 1082, 1086, 1074, 1072, 1103, 32, 1095, 1072, 1089, 1090, 1100)
End Function

Private Function StaffHeaderTableNumber() As String
    StaffHeaderTableNumber = Ru(1090, 1072, 1073, 1077, 1083, 1100, 1085, 1099, 1081, 32, 1085, 1086, 1084, 1077, 1088)
End Function

Private Function StaffHeaderBirthDate() As String
    StaffHeaderBirthDate = Ru(1076, 1072, 1090, 1072, 32, 1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103)
End Function

Private Function StaffHeaderCitizenship() As String
    StaffHeaderCitizenship = Ru(1075, 1088, 1072, 1078, 1076, 1072, 1085, 1089, 1090, 1074, 1086)
End Function

Private Function StaffHeaderEmployeeGroup() As String
    StaffHeaderEmployeeGroup = Ru(1075, 1088, 1091, 1087, 1087, 1072, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1086, 1074)
End Function

Private Function StaffHeaderContractKind() As String
    StaffHeaderContractKind = Ru(1074, 1080, 1076, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffHeaderContractType() As String
    StaffHeaderContractType = Ru(1090, 1080, 1087, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffHeaderContractStartDate() As String
    StaffHeaderContractStartDate = Ru(1076, 1072, 1090, 1072, 32, 1085, 1072, 1095, 1072, 1083, 1072, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffHeaderVus() As String
    StaffHeaderVus = Ru(1042, 1059, 1057)
End Function

Private Function StaffHeaderTariffRank() As String
    StaffHeaderTariffRank = Ru(1090, 1072, 1088, 1080, 1092, 1085, 1099, 1081, 32, 1088, 1072, 1079, 1088, 1103, 1076)
End Function

Private Function StaffHeaderBankAccount() As String
    StaffHeaderBankAccount = Ru(1085, 1086, 1084, 1077, 1088, 32, 1089, 1095, 1077, 1090, 1072, 32, 1074, 32, 1073, 1072, 1085, 1082, 1077)
End Function

Private Function StaffMessageMissingColumnsTitle() As String
    StaffMessageMissingColumnsTitle = Ru(1057, 1090, 1088, 1091, 1082, 1090, 1091, 1088, 1072, 32, 1083, 1080, 1089, 1090, 1072)
End Function

Private Function StaffMessageMissingColumnsIntro() As String
    StaffMessageMissingColumnsIntro = Ru(1054, 1096, 1080, 1073, 1082, 1072, 32, 1087, 1088, 1080, 32, 1087, 1086, 1080, 1089, 1082, 1077, 32, 1086, 1073, 1103, 1079, 1072, 1090, 1077, 1083, 1100, 1085, 1099, 1093, 32, 1082, 1086, 1083, 1086, 1085, 1086, 1082, 32, 1083, 1080, 1089, 1090, 1072, 32, 39, 1064, 1090, 1072, 1090, 39, 46)
End Function

Private Function StaffMessageMissingColumnsList() As String
    StaffMessageMissingColumnsList = Ru(1053, 1077, 32, 1091, 1076, 1072, 1083, 1086, 1089, 1100, 32, 1086, 1087, 1088, 1077, 1076, 1077, 1083, 1080, 1090, 1100, 32, 1089, 1083, 1077, 1076, 1091, 1102, 1097, 1080, 1077, 32, 1089, 1090, 1086, 1083, 1073, 1094, 1099, 58)
End Function

Private Function StaffMessageMissingColumnsHint() As String
    StaffMessageMissingColumnsHint = Ru(1055, 1088, 1086, 1074, 1077, 1088, 1100, 1090, 1077, 32, 1079, 1072, 1075, 1086, 1083, 1086, 1074, 1082, 1080, 32, 1085, 1072, 32, 1083, 1080, 1089, 1090, 1077, 32, 39, 1064, 1090, 1072, 1090, 39, 46)
End Function

Private Function StaffKeyFio() As String
    StaffKeyFio = Ru(1051, 1080, 1094, 1086)
End Function

Private Function StaffKeyPersonalNumber() As String
    StaffKeyPersonalNumber = Ru(1051, 1080, 1095, 1085, 1099, 1081, 32, 1085, 1086, 1084, 1077, 1088)
End Function

Private Function StaffKeyRank() As String
    StaffKeyRank = Ru(1042, 1086, 1080, 1085, 1089, 1082, 1086, 1077, 32, 1079, 1074, 1072, 1085, 1080, 1077)
End Function

Private Function StaffKeyUnit() As String
    StaffKeyUnit = Ru(1063, 1072, 1089, 1090, 1100)
End Function

Private Function StaffKeyPosition() As String
    StaffKeyPosition = Ru(1064, 1090, 1072, 1090, 1085, 1072, 1103, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100)
End Function

Private Function StaffKeyTableNumber() As String
    StaffKeyTableNumber = Ru(1058, 1072, 1073, 1077, 1083, 1100, 1085, 1099, 1081, 32, 1085, 1086, 1084, 1077, 1088)
End Function

Private Function StaffKeyBirthDate() As String
    StaffKeyBirthDate = Ru(1044, 1072, 1090, 1072, 32, 1088, 1086, 1078, 1076, 1077, 1085, 1080, 1103)
End Function

Private Function StaffKeyCitizenship() As String
    StaffKeyCitizenship = Ru(1043, 1088, 1072, 1078, 1076, 1072, 1085, 1089, 1090, 1074, 1086)
End Function

Private Function StaffKeyServiceCategory() As String
    StaffKeyServiceCategory = Ru(1043, 1088, 1091, 1087, 1087, 1072, 32, 1089, 1086, 1090, 1088, 1091, 1076, 1085, 1080, 1082, 1086, 1074)
End Function

Private Function StaffKeyContractKind() As String
    StaffKeyContractKind = Ru(1042, 1080, 1076, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffKeyContractType() As String
    StaffKeyContractType = Ru(1058, 1080, 1087, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffKeyContractStartDate() As String
    StaffKeyContractStartDate = Ru(1044, 1072, 1090, 1072, 32, 1085, 1072, 1095, 1072, 1083, 1072, 32, 1082, 1086, 1085, 1090, 1088, 1072, 1082, 1090, 1072)
End Function

Private Function StaffKeyVus() As String
    StaffKeyVus = Ru(1042, 1059, 1057)
End Function

Private Function StaffKeyTariffRank() As String
    StaffKeyTariffRank = Ru(1058, 1072, 1088, 1080, 1092, 1085, 1099, 1081, 32, 1088, 1072, 1079, 1088, 1103, 1076)
End Function

Private Function StaffKeyBankAccount() As String
    StaffKeyBankAccount = Ru(1053, 1086, 1084, 1077, 1088, 32, 1089, 1095, 1077, 1090, 1072, 32, 1074, 32, 1073, 1072, 1085, 1082, 1077)
End Function

' ==============================================================================
' 1. INITIALIZATION & COLUMN FINDING
' ==============================================================================

Public Sub InitStaffColumnIndexes()
    Dim wsStaff As Worksheet
    Set wsStaff = GetStaffWorksheet()
    If wsStaff Is Nothing Then Exit Sub
    
    If Not FindColumnNumbers(wsStaff, colLichniyNomer_Global, colZvanie_Global, colFIO_Global, colDolzhnost_Global, colVoinskayaChast_Global) Then
        ' ????????? ??? ????????? ?????? FindColumnNumbers, ????? ?????? ????????? ??????
        End
    End If
End Sub

Public Sub EnsureStaffColumnsInitialized()
    If colLichniyNomer_Global = 0 Or colFIO_Global = 0 Then
        InitStaffColumnIndexes
    End If
End Sub

Public Function FindStaffRow(ws As Worksheet, lichniyNomer As String, colNum As Long) As Long
    Dim res As Variant
    res = Application.Match(lichniyNomer, ws.Columns(colNum), 0)
    If isError(res) Then
        FindStaffRow = 0
    Else
        FindStaffRow = CLng(res)
    End If
End Function

Public Function FindColumnNumbers(ws As Worksheet, ByRef colLichniyNomer As Long, ByRef colZvanie As Long, ByRef colFIO As Long, ByRef colDolzhnost As Long, ByRef colVoinskayaChast As Long) As Boolean
    Dim lastCol As Long, i As Long, headerText As String
    Dim fioCandidate As Long, personnelSectionFallback As Long
    Dim personalNumberHeader As String
    Dim rankHeader As String
    Dim shortRankHeader As String
    Dim positionHeader As String
    Dim staffPositionHeader As String
    Dim fioHeader As String
    Dim unitHeader As String
    Dim militaryUnitHeader As String
    Dim personnelSectionHeader As String

    colLichniyNomer = 0: colZvanie = 0: colFIO = 0: colDolzhnost = 0: colVoinskayaChast = 0

    personalNumberHeader = NormalizeHeaderText(StaffHeaderPersonalNumber())
    rankHeader = NormalizeHeaderText(StaffHeaderRank())
    shortRankHeader = NormalizeHeaderText(StaffHeaderShortRank())
    positionHeader = NormalizeHeaderText(StaffHeaderPosition())
    staffPositionHeader = NormalizeHeaderText(StaffHeaderStaffPosition())
    fioHeader = NormalizeHeaderText(StaffHeaderFio())
    unitHeader = NormalizeHeaderText(StaffHeaderUnit())
    militaryUnitHeader = NormalizeHeaderText(StaffHeaderMilitaryUnit())
    personnelSectionHeader = NormalizeHeaderText(StaffHeaderPersonnelSection())

    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For i = 1 To lastCol
        headerText = NormalizeHeaderText(CStr(ws.Cells(1, i).Value))

        If headerText = personalNumberHeader Then
            colLichniyNomer = i
        ElseIf headerText = rankHeader Or headerText = shortRankHeader Then
            colZvanie = i
        ElseIf headerText = positionHeader Or headerText = staffPositionHeader Then
            If IsTextColumn(ws, i) Or IsLongPositionColumn(ws, i) Then
                colDolzhnost = i
            ElseIf colDolzhnost = 0 Then
                colDolzhnost = i
            End If
        ElseIf headerText = fioHeader Then
            If IsTextFIOColumn(ws, i) Then
                colFIO = i
            ElseIf fioCandidate = 0 Then
                fioCandidate = i
            End If
        ElseIf headerText = unitHeader Or headerText = militaryUnitHeader Then
            colVoinskayaChast = i
        ElseIf headerText = personnelSectionHeader Then
            If personnelSectionFallback = 0 Then personnelSectionFallback = i
            If ColumnContainsUnitText(ws, i) Then colVoinskayaChast = i
        End If
    Next i

    If colFIO = 0 Then colFIO = fioCandidate
    If colVoinskayaChast = 0 Then colVoinskayaChast = personnelSectionFallback

    If colLichniyNomer > 0 And colZvanie > 0 And colFIO > 0 And colDolzhnost > 0 And colVoinskayaChast > 0 Then
        FindColumnNumbers = True
    Else
        FindColumnNumbers = False
        
        ' ????????? ????????? ????? ?? ????????????? ????????
        Dim missingCols As String
        missingCols = ""
        If colLichniyNomer = 0 Then missingCols = missingCols & "- " & StaffKeyPersonalNumber() & vbCrLf
        If colZvanie = 0 Then missingCols = missingCols & "- " & StaffKeyRank() & vbCrLf
        If colFIO = 0 Then missingCols = missingCols & "- " & Ru(1051, 1080, 1094, 1086, 32, 40, 1060, 1048, 1054, 41) & vbCrLf
        If colDolzhnost = 0 Then missingCols = missingCols & "- " & StaffKeyPosition() & vbCrLf
        If colVoinskayaChast = 0 Then missingCols = missingCols & "- " & Ru(1063, 1072, 1089, 1090, 1100, 32, 1080, 1083, 1080, 32, 1056, 1072, 1079, 1076, 1077, 1083, 32, 1087, 1077, 1088, 1089, 1086, 1085, 1072, 1083, 1072) & vbCrLf
        
        MsgBox StaffMessageMissingColumnsIntro() & vbCrLf & vbCrLf & _
               StaffMessageMissingColumnsList() & vbCrLf & missingCols & vbCrLf & _
               StaffMessageMissingColumnsHint(), vbCritical, StaffMessageMissingColumnsTitle()
    End If
End Function

Public Function GetStaffWorksheet() As Worksheet
    On Error Resume Next
    Set GetStaffWorksheet = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_STAFF)
    On Error GoTo 0
End Function

Public Function GetDsoWorksheet() As Worksheet
    Dim ws As Worksheet
    Dim targetName As String

    targetName = Ru(1044, 1057, 1054)

    On Error Resume Next
    Set GetDsoWorksheet = ThisWorkbook.Worksheets(targetName)
    On Error GoTo 0
    If Not GetDsoWorksheet Is Nothing Then Exit Function

    For Each ws In ThisWorkbook.Worksheets
        If LCase$(Trim$(ws.CodeName)) = LCase$(Trim$(Ru(1051, 1080, 1089, 1090) & "1")) Then
            Set GetDsoWorksheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function NormalizeHeaderText(ByVal rawValue As String) As String
    NormalizeHeaderText = LCase$(Trim$(Replace$(Replace$(rawValue, vbCr, " "), vbLf, " ")))
    Do While InStr(NormalizeHeaderText, "  ") > 0
        NormalizeHeaderText = Replace$(NormalizeHeaderText, "  ", " ")
    Loop
End Function

Private Function ColumnContainsUnitText(ByVal ws As Worksheet, ByVal colNum As Long) As Boolean
    Dim lastRow As Long, i As Long
    Dim valueText As String

    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    If lastRow > 20 Then lastRow = 20

    For i = 2 To lastRow
        valueText = LCase$(Trim$(CStr(ws.Cells(i, colNum).Value)))
        If valueText <> "" Then
            If InStr(valueText, LCase$(StaffHeaderMilitaryUnit())) > 0 _
               Or InStr(valueText, LCase$(StaffHeaderArmyUnitAlt())) > 0 _
               Or InStr(valueText, "РІ/С‡") > 0 Then
                ColumnContainsUnitText = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function IsTextFIOColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim lastRow As Long, i As Long, value As String
    lastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
    Dim checkLimit As Long: checkLimit = IIf(lastRow > 50, 50, lastRow)
    Dim textCount As Long
    For i = 2 To checkLimit
        value = Trim(ws.Cells(i, colNum).value)
        If value <> "" And ContainsLetters(value) And InStr(value, " ") > 0 And Not IsNumeric(value) Then
            textCount = textCount + 1
        End If
    Next i
    IsTextFIOColumn = (textCount > 0)
End Function

Private Function IsTextColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, value As String
    For i = 2 To 20
        value = Trim(ws.Cells(i, colNum).value)
        If value <> "" And ContainsLetters(value) Then IsTextColumn = True: Exit Function
    Next i
    IsTextColumn = False
End Function

Private Function IsLongPositionColumn(ws As Worksheet, colNum As Long) As Boolean
    Dim i As Long, numCount As Long, totalCount As Long, val As String
    Dim lastRow As Long: lastRow = IIf(ws.Cells(ws.Rows.count, colNum).End(xlUp).Row > 50, 50, ws.Cells(ws.Rows.count, colNum).End(xlUp).Row)
    For i = 2 To lastRow
        val = Trim(ws.Cells(i, colNum).value)
        If val <> "" Then
            totalCount = totalCount + 1
            If ContainsLetters(val) And ContainsNumbers(val) Then numCount = numCount + 1
        End If
    Next i
    If totalCount > 0 Then IsLongPositionColumn = (numCount / totalCount) > 0.3 Else IsLongPositionColumn = False
End Function

' ==============================================================================
' 2. GENERAL UTILITIES
' ==============================================================================

Public Function ContainsLetters(Text As String) As Boolean
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(Text)
        currentChar = Mid$(Text, i, 1)
        If currentChar <> "" Then
            If LCase$(currentChar) <> UCase$(currentChar) Then
                ContainsLetters = True
                Exit Function
            End If
        End If
    Next i

    ContainsLetters = False
End Function

Public Function ContainsNumbers(Text As String) As Boolean
    Dim i As Long, char As String
    For i = 1 To Len(Text)
        char = Mid(Text, i, 1)
        If char >= "0" And char <= "9" Then ContainsNumbers = True: Exit Function
    Next i
    ContainsNumbers = False
End Function

Public Function GetLastRow(ws As Worksheet, colIndex As Variant) As Long
    Dim colNum As Long
    If VarType(colIndex) = vbString Then colNum = ws.Range(CStr(colIndex) & "1").Column Else colNum = CLng(colIndex)
    GetLastRow = ws.Cells(ws.Rows.count, colNum).End(xlUp).Row
End Function

Public Function ExtractVoinskayaChast(inputText As String) As String
    ExtractVoinskayaChast = RegExpExtract(inputText, "\d{4,5}")
    If ExtractVoinskayaChast = "" Then ExtractVoinskayaChast = inputText
End Function

Public Function GetExportCutoffDate() As Date
    GetExportCutoffDate = DateAdd("m", -1, DateAdd("yyyy", -3, Date))
End Function

' ==============================================================================
' 3. SETTINGS & CONFIGURATION
' ==============================================================================

Public Function GetSettingCutBattalion() As Boolean
    On Error Resume Next
    Dim ws As Worksheet, val As Variant
    Set ws = ThisWorkbook.Sheets("?????????")
    If ws Is Nothing Then GetSettingCutBattalion = True: Exit Function
    val = ws.Range("B2").value
    If IsEmpty(val) Then GetSettingCutBattalion = True: Exit Function
    If UCase(CStr(val)) = "???" Or val = False Or val = 0 Then GetSettingCutBattalion = False Else GetSettingCutBattalion = True
End Function

Public Sub SetupSettingsSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("?????????")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "?????????"
        ws.Cells(1, 1).value = "????????": ws.Cells(1, 2).value = "????????"
        ws.Cells(2, 1).value = "???????? ???????? ??????????": ws.Cells(2, 2).value = "??"
        ws.Columns("A:B").AutoFit
        MsgBox "???? '?????????' ??????.", vbInformation
    Else
        MsgBox "???? '?????????' ??? ??????????.", vbInformation
    End If
End Sub

' ==============================================================================
' 4. DATE PARSING & SEARCH
' ==============================================================================

Public Function ParseDateSafe(val As Variant) As Date
    On Error Resume Next
    ParseDateSafe = 0
    If IsEmpty(val) Or Trim(CStr(val)) = "" Then Exit Function
    
    Dim d As Date, sVal As String
    sVal = Trim(CStr(val))

    ' 0. Excel date serials (for example, 45220)
    If IsNumeric(val) Or IsNumeric(sVal) Then
        d = CDate(CDbl(val))
        If Year(d) > 2000 And Year(d) < 2100 Then
            ParseDateSafe = d
            Exit Function
        End If
    End If
    
    ' 1. ??????????? ???????????
    If IsDate(sVal) Then
        d = CDate(sVal)
        If Year(d) < 2000 And Year(d) > 1900 Then d = DateSerial(Year(d) + 100, Month(d), Day(d))
        If Year(d) > 2000 And Year(d) < 2100 Then ParseDateSafe = d: Exit Function
    End If
    
    ' 2. ?????? ??????
    Dim parts() As String
    If InStr(sVal, ".") > 0 Then
        parts = Split(sVal, ".")
    ElseIf InStr(sVal, "/") > 0 Then
        parts = Split(sVal, "/")
    End If
    
    If (Not Not parts) <> 0 Then
        If UBound(parts) = 2 Then
            Dim y As Integer: y = CInt(parts(2))
            If y < 100 Then y = 2000 + y
            If y > 2000 And y < 2100 Then ParseDateSafe = DateSerial(y, CInt(parts(1)), CInt(parts(0)))
        End If
    End If
End Function

Public Function GetStaffData(queryValue As String, Optional byLichniyNomer As Boolean = True) As Object
    Call EnsureStaffColumnsInitialized
    Dim ws As Worksheet, d As Object, r As Long, searchCol As Long
    Set ws = GetStaffWorksheet()
    Set d = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set GetStaffData = d
        Exit Function
    End If
    
    If byLichniyNomer Then searchCol = colLichniyNomer_Global Else searchCol = colFIO_Global
    r = FindStaffRow(ws, queryValue, searchCol)
    
    If r > 0 Then
        FillStaffDictionaryFromRow ws, r, d
    End If
    Set GetStaffData = d
End Function

Public Function FindEmployeeByAnyNumber(number As String) As Object
    Dim res As Object
    Set res = GetStaffData(number, True)
    
    ' ???? ?? ????? ?? ???????, ??????? ?? ??????????
    If res.count = 0 Then
        Dim wsStaff As Worksheet
        Dim colTable As Long
        Set wsStaff = GetStaffWorksheet()
        If wsStaff Is Nothing Then
            Set FindEmployeeByAnyNumber = res
            Exit Function
        End If
        colTable = FindTableNumberColumn(wsStaff)
        
        If colTable > 0 Then
            Dim r As Long
            r = FindStaffRow(wsStaff, number, colTable)
            If r > 0 Then
                Set res = CreateObject("Scripting.Dictionary")
                FillStaffDictionaryFromRow wsStaff, r, res
            End If
        End If
    End If
    Set FindEmployeeByAnyNumber = res
End Function

Private Sub FillStaffDictionaryFromRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal targetDictionary As Object)
    Dim colTable As Long
    Dim colBirthDate As Long
    Dim colCitizenship As Long
    Dim colEmployeeGroup As Long
    Dim colContractKind As Long
    Dim colContractType As Long
    Dim colContractStartDate As Long
    Dim colVus As Long
    Dim colTariffRank As Long
    Dim colBankAccount As Long

    If targetDictionary Is Nothing Then Exit Sub
    If rowNum < 2 Then Exit Sub

    targetDictionary(StaffKeyFio()) = ws.Cells(rowNum, colFIO_Global).Value
    targetDictionary(StaffKeyPersonalNumber()) = ws.Cells(rowNum, colLichniyNomer_Global).Value
    targetDictionary(StaffKeyRank()) = ws.Cells(rowNum, colZvanie_Global).Value
    targetDictionary(StaffKeyUnit()) = ws.Cells(rowNum, colVoinskayaChast_Global).Value
    targetDictionary(StaffKeyPosition()) = ws.Cells(rowNum, colDolzhnost_Global).Value

    colTable = FindTableNumberColumn(ws)
    If colTable > 0 Then targetDictionary(StaffKeyTableNumber()) = ws.Cells(rowNum, colTable).Value

    colBirthDate = FindOptionalStaffColumn(ws, StaffHeaderBirthDate())
    If colBirthDate > 0 Then targetDictionary(StaffKeyBirthDate()) = ws.Cells(rowNum, colBirthDate).Value

    colCitizenship = FindOptionalStaffColumn(ws, StaffHeaderCitizenship())
    If colCitizenship > 0 Then targetDictionary(StaffKeyCitizenship()) = ws.Cells(rowNum, colCitizenship).Value

    colEmployeeGroup = FindOptionalStaffColumn(ws, StaffHeaderEmployeeGroup())
    If colEmployeeGroup > 0 Then targetDictionary(StaffKeyServiceCategory()) = ws.Cells(rowNum, colEmployeeGroup).Value

    colContractKind = FindOptionalStaffColumn(ws, StaffHeaderContractKind())
    If colContractKind > 0 Then targetDictionary(StaffKeyContractKind()) = ws.Cells(rowNum, colContractKind).Value

    colContractType = FindOptionalStaffColumn(ws, StaffHeaderContractType())
    If colContractType > 0 Then targetDictionary(StaffKeyContractType()) = ws.Cells(rowNum, colContractType).Value

    colContractStartDate = FindOptionalStaffColumn(ws, StaffHeaderContractStartDate())
    If colContractStartDate > 0 Then targetDictionary(StaffKeyContractStartDate()) = ws.Cells(rowNum, colContractStartDate).Value

    colVus = FindOptionalStaffColumn(ws, StaffHeaderVus())
    If colVus > 0 Then targetDictionary(StaffKeyVus()) = ws.Cells(rowNum, colVus).Value

    colTariffRank = FindOptionalStaffColumn(ws, StaffHeaderTariffRank())
    If colTariffRank > 0 Then targetDictionary(StaffKeyTariffRank()) = ws.Cells(rowNum, colTariffRank).Value

    colBankAccount = FindOptionalStaffColumn(ws, StaffHeaderBankAccount())
    If colBankAccount > 0 Then targetDictionary(StaffKeyBankAccount()) = ws.Cells(rowNum, colBankAccount).Value
End Sub

Private Function FindOptionalStaffColumn(ByVal ws As Worksheet, ParamArray headerCandidates()) As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim headerText As String
    Dim candidateText As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        headerText = NormalizeHeaderText(CStr(ws.Cells(1, i).Value))
        If headerText <> "" Then
            For j = LBound(headerCandidates) To UBound(headerCandidates)
                candidateText = NormalizeHeaderText(CStr(headerCandidates(j)))
                If candidateText <> "" Then
                    If headerText = candidateText Or InStr(headerText, candidateText) > 0 Then
                        FindOptionalStaffColumn = i
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i
End Function

Public Function FindTableNumberColumn(ws As Worksheet) As Long
    Dim i As Long, val As Variant, headerText As String
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        headerText = NormalizeHeaderText(CStr(ws.Cells(1, i).Value))
        If headerText = "в„–" Or headerText = NormalizeHeaderText(StaffHeaderTableNumber()) Or InStr(headerText, NormalizeHeaderText(StaffHeaderTableNumber())) > 0 Then
            val = ws.Cells(2, i).value
            If IsNumeric(val) And Not IsEmpty(val) Then
                FindTableNumberColumn = i
                Exit Function
            End If
        End If
    Next i
    FindTableNumberColumn = 0
End Function

Public Function GetStaffDataByTableNumber(tableNumber As String) As Object
    ' ????? ??? ?????????????
    Set GetStaffDataByTableNumber = FindEmployeeByAnyNumber(tableNumber)
End Function

Public Sub SaveWordDocumentSafe(wdDoc As Object, filePath As String)
    On Error GoTo SaveAsFallback
    wdDoc.SaveAs2 filePath, 16 ' wdFormatXMLDocument
    Exit Sub

SaveAsFallback:
    Err.Clear
    On Error GoTo SaveFailed
    wdDoc.SaveAs filePath
    Exit Sub

SaveFailed:
    Err.Raise vbObjectError + 901, "SaveWordDocumentSafe", "РќРµ СѓРґР°Р»РѕСЃСЊ СЃРѕС…СЂР°РЅРёС‚СЊ РґРѕРєСѓРјРµРЅС‚ Word: " & filePath & ". " & Err.Description
End Sub

Public Function IsWordAvailable() As Boolean
    On Error Resume Next
    Dim app As Object: Set app = CreateObject("Word.Application")
    IsWordAvailable = Not app Is Nothing
    If Not app Is Nothing Then app.Quit
End Function

Public Sub CollectAllPersonPeriods(ws As Worksheet, rowNum As Long, periods As Collection)
    Dim lastCol As Long, j As Long, d1 As Date, d2 As Date
    lastCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).Column
    For j = 5 To lastCol Step 2
        d1 = ParseDateSafe(ws.Cells(rowNum, j).value)
        d2 = ParseDateSafe(ws.Cells(rowNum, j + 1).value)
        If d1 > 0 And d2 > 0 Then
            Dim p As Collection: Set p = New Collection
            p.Add d1: p.Add d2: p.Add (DateDiff("d", d1, d2) + 1)
            periods.Add p
        End If
    Next j
End Sub

Public Function FormatPeriodsForWord(periods As Collection, cutoff As Date, ByRef t As Long, ByRef r As Long) As String
    Dim p As Collection, s As String, i As Long
    t = 0: r = 0
    If periods.count = 0 Then Exit Function
    
    Dim sorted As Collection
    Set sorted = SortPeriodsByDateStart(periods)
    
    For i = 1 To sorted.count
        Set p = sorted(i)
        t = t + p(3)
        s = s & "- ? " & Format(p(1), "dd.mm.yy") & " ?? " & Format(p(2), "dd.mm.yy") & " (" & p(3) & " ???.)"
        If p(2) < cutoff Then s = s & " (?? ????????)"
        s = s & vbCrLf
    Next i
    r = (t \ 3) * 2
    FormatPeriodsForWord = s
End Function

Public Function HasInvalidPair(periods As Collection) As Boolean
    Dim p As Collection
    For Each p In periods
        If p(2) < p(1) Then HasInvalidPair = True: Exit Function
    Next p
    HasInvalidPair = False
End Function

Public Function SortPeriodsByDateStart(periods As Collection) As Collection
    Dim arr() As Variant, i As Long, j As Long, n As Long
    n = periods.count
    If n <= 1 Then Set SortPeriodsByDateStart = periods: Exit Function
    ReDim arr(1 To n)
    For i = 1 To n: Set arr(i) = periods(i): Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i)(1) > arr(j)(1) Then
                Dim tmp As Collection: Set tmp = arr(i)
                Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
    Dim resCol As Collection: Set resCol = New Collection
    For i = 1 To n: resCol.Add arr(i): Next i
    Set SortPeriodsByDateStart = resCol
End Function

Public Function hasCriticalErrors() As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("???")
    If ws Is Nothing Then hasCriticalErrors = True: Exit Function
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "C").End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    Dim i As Long, j As Long
    For i = 2 To lastRow
        For j = 5 To lastCol Step 2
            If ws.Cells(i, j).Interior.Color = RGB(255, 100, 100) Or ws.Cells(i, j).Interior.Color = RGB(255, 200, 200) Then
                hasCriticalErrors = True: Exit Function
            End If
        Next j
    Next i
    hasCriticalErrors = False
End Function

' ==============================================================================
' 5. SMART POSITION PARSER
' ==============================================================================

Public Function SklonitDolzhnost(dolzhnost As String, VoinskayaChast As String) As String
    Dim clean As String, role As String, body As String, res As String
    clean = LCase(Trim(dolzhnost))
    
    ' 1. ????? ?????
    clean = CutUnitTail(clean)
    
    ' 2. ?????????
    Call SplitRoleAndBody(clean, role, body)
    
    ' 3. ???????? ????
    role = SklonitVoennayaDolzhnost(role)
    
    res = role
    If body <> "" Then res = res & " " & body
    SklonitDolzhnost = res & " ????????? ????? " & VoinskayaChast
End Function

Private Function CutUnitTail(Text As String) As String
    Dim patterns As Variant, i As Long, cutBattalion As Boolean
    cutBattalion = GetSettingCutBattalion()
    Dim t As String: t = Text
    
    ' 1. ???????????? ?????: ??????
    patterns = Array( _
        "(\d+\s+)?(??????????|????????????|????????????????)\s+.*", _
        "\d+\s+(?????|???????|???????|?????|?????????).*" _
    )
    For i = LBound(patterns) To UBound(patterns)
        t = RegExpReplace(t, patterns(i), "")
    Next i
    
    ' 2. ??????? ??????????
    patterns = Array( _
        "(??????????|?????)\s+(?????|???????|???????).*" _
    )
    For i = LBound(patterns) To UBound(patterns)
        t = RegExpReplace(t, patterns(i), "")
    Next i
    
    ' 3. ???????? ?????
    If cutBattalion Then
        patterns = Array("(\d+\s+)?(?????????|?????????|??????????).*")
        For i = LBound(patterns) To UBound(patterns)
            t = RegExpReplace(t, patterns(i), "")
        Next i
    End If
    
    CutUnitTail = Trim(t)
End Function

Private Sub SplitRoleAndBody(Text As String, ByRef roleOut As String, ByRef bodyOut As String)
    Dim words() As String, i As Long, splitIdx As Long
    words = Split(Text, " ")
    splitIdx = UBound(words)
    
    For i = 0 To UBound(words)
        If IsUnitKeyword(words(i)) Then
            splitIdx = i - 1
            Exit For
        End If
    Next i
    
    Dim r As String, b As String
    For i = 0 To UBound(words)
        If i <= splitIdx Then r = r & words(i) & " " Else b = b & words(i) & " "
    Next i
    roleOut = Trim(r): bodyOut = Trim(b)
End Sub

Private Function IsUnitKeyword(word As String) As Boolean
    Select Case LCase(word)
        Case "??????", "????", "?????????", "???????", "???????", "??????", "???????", "???????", "???????", "??????", "??????", "????", "??????", "?????", "???????", "?????", "?????????", "??????", "????????", "??????????"
            IsUnitKeyword = True
        Case Else
            IsUnitKeyword = False
    End Select
End Function

Public Function SklonitVoennayaDolzhnost(dolzhnost As String) As String
    Dim res As String: res = dolzhnost
    If Left(res, 8) = "??????? " Then res = "???????? " & Mid(res, 9)
    If Left(res, 8) = "??????? " Then res = "???????? " & Mid(res, 9)
    If Left(res, 8) = "??????? " Then res = "???????? " & Mid(res, 9)
    If Left(res, 8) = "??????? " Then res = "???????? " & Mid(res, 9)
    
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "?????????", "??????????")
    res = Replace(res, "???????????", "???????????")
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "???????", "????????")
    res = Replace(res, "????????", "????????")
    res = Replace(res, "???????????????", "????????????????")
    res = Replace(res, "?????????", "??????????")
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "??????????", "???????????")
    res = Replace(res, "??????", "???????")
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "??????????", "???????????")
    res = Replace(res, "?????????????", "??????????????")
    res = Replace(res, "???????", "???????")
    res = Replace(res, "?????", "??????")
    res = Replace(res, "???????", "????????")
    res = Replace(res, "???????", "????????")
    res = Replace(res, "?????????????????", "?????????????????")
    res = Replace(res, "????????", "?????????")
    res = Replace(res, "????????", "????????")
    res = Replace(res, "?????????", "??????????")
    res = Replace(res, "????", "?????")
    res = Replace(res, "????????", "?????????")
    SklonitVoennayaDolzhnost = res
End Function

Public Function GetDolzhnostImenitelny(dolzhnost As String, VoinskayaChast As String) As String
    GetDolzhnostImenitelny = CutUnitTail(LCase(Trim(dolzhnost))) & " ????????? ????? " & VoinskayaChast
End Function

Public Function SklonitZvanie(zvanie As String) As String
    Dim z As String: z = LCase(Trim(zvanie))
    Select Case z
        Case "???????": SklonitZvanie = "????????"
        Case "????????": SklonitZvanie = "?????????"
        Case "??????? ???????": SklonitZvanie = "???????? ????????"
        Case "???????": SklonitZvanie = "????????"
        Case "??????? ???????": SklonitZvanie = "???????? ????????"
        Case "????????": SklonitZvanie = "????????"
        Case "?????????": SklonitZvanie = "??????????"
        Case "??????? ?????????": SklonitZvanie = "???????? ??????????"
        Case "??????? ?????????": SklonitZvanie = "???????? ??????????"
        Case "?????????": SklonitZvanie = "??????????"
        Case "??????? ?????????": SklonitZvanie = "???????? ??????????"
        Case "???????": SklonitZvanie = "????????"
        Case "?????": SklonitZvanie = "??????"
        Case "????????????": SklonitZvanie = "?????????????"
        Case "?????????": SklonitZvanie = "??????????"
        Case "???????-?????": SklonitZvanie = "???????-??????"
        Case Else: SklonitZvanie = UCase(Left(z, 1)) & Mid(z, 2)
    End Select
End Function

Public Function GetZvanieImenitelny(zvanie As String) As String
    GetZvanieImenitelny = LCase(Trim(zvanie))
End Function

Public Function GetZvanieSkrasheno(zvanie As String) As String
    Dim z As String: z = LCase(Trim(zvanie))
    Select Case z
        Case "??????? ???????": GetZvanieSkrasheno = "??. ????????"
        Case "??????? ???????": GetZvanieSkrasheno = "??. ????????"
        Case "??????? ?????????": GetZvanieSkrasheno = "??. ??????????"
        Case "??????? ?????????": GetZvanieSkrasheno = "??. ??????????"
        Case "??????? ?????????": GetZvanieSkrasheno = "??. ??????????"
        Case Else: GetZvanieSkrasheno = SklonitZvanie(z)
    End Select
End Function

Public Function GetZvanieImenitelnyForSignature(zvanie As String) As String
    GetZvanieImenitelnyForSignature = UCase(Left(zvanie, 1)) & LCase(Mid(zvanie, 2))
End Function

Public Function GetFIOWithInitials(sName As String) As String
    GetFIOWithInitials = BuildInitialsDisplay(sName, FIO_CASE_DATIVE)
End Function

Public Function GetFIOWithInitialsImenitelny(sName As String) As String
    GetFIOWithInitialsImenitelny = BuildInitialsDisplay(sName, FIO_CASE_NOMINATIVE)
End Function

Public Function SklonitFIO(sName As String) As String
    SklonitFIO = fio(sName, FIO_CASE_DATIVE)
End Function

Public Function GetMonthNameRussian(monthNumber As Integer) As String
    Select Case monthNumber
        Case 1: GetMonthNameRussian = BuildUnicodeTextHelper(1103, 1085, 1074, 1072, 1088, 1103)
        Case 2: GetMonthNameRussian = BuildUnicodeTextHelper(1092, 1077, 1074, 1088, 1072, 1083, 1103)
        Case 3: GetMonthNameRussian = BuildUnicodeTextHelper(1084, 1072, 1088, 1090, 1072)
        Case 4: GetMonthNameRussian = BuildUnicodeTextHelper(1072, 1087, 1088, 1077, 1083, 1103)
        Case 5: GetMonthNameRussian = BuildUnicodeTextHelper(1084, 1072, 1103)
        Case 6: GetMonthNameRussian = BuildUnicodeTextHelper(1080, 1102, 1085, 1103)
        Case 7: GetMonthNameRussian = BuildUnicodeTextHelper(1080, 1102, 1083, 1103)
        Case 8: GetMonthNameRussian = BuildUnicodeTextHelper(1072, 1074, 1075, 1091, 1089, 1090, 1072)
        Case 9: GetMonthNameRussian = BuildUnicodeTextHelper(1089, 1077, 1085, 1090, 1103, 1073, 1088, 1103)
        Case 10: GetMonthNameRussian = BuildUnicodeTextHelper(1086, 1082, 1090, 1103, 1073, 1088, 1103)
        Case 11: GetMonthNameRussian = BuildUnicodeTextHelper(1085, 1086, 1103, 1073, 1088, 1103)
        Case 12: GetMonthNameRussian = BuildUnicodeTextHelper(1076, 1077, 1082, 1072, 1073, 1088, 1103)
        Case Else: GetMonthNameRussian = BuildUnicodeTextHelper(1085, 1077, 1080, 1079, 1074, 1077, 1089, 1090, 1085, 1099, 1081, 32, 1084, 1077, 1089, 1103, 1094)
    End Select
End Function

Private Function BuildUnicodeTextHelper(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(codePoints) To UBound(codePoints)
        result = result & ChrW$(CLng(codePoints(i)))
    Next i

    BuildUnicodeTextHelper = result
End Function

Public Function Ru(ParamArray codePoints() As Variant) As String
    Dim i As Long
    Dim resultText As String

    For i = LBound(codePoints) To UBound(codePoints)
        resultText = resultText & ChrW$(CLng(codePoints(i)))
    Next i

    Ru = resultText
End Function

Private Function NormalizeFioCaseCode(ByVal nameCase As String) As String
    Dim normalizedCase As String

    normalizedCase = UCase$(Trim$(nameCase))

    Select Case normalizedCase
        Case "", FIO_CASE_DATIVE, BuildUnicodeTextHelper(1044)
            NormalizeFioCaseCode = FIO_CASE_DATIVE
        Case FIO_CASE_GENITIVE, BuildUnicodeTextHelper(1056)
            NormalizeFioCaseCode = FIO_CASE_GENITIVE
        Case FIO_CASE_NOMINATIVE, "I", BuildUnicodeTextHelper(1048)
            NormalizeFioCaseCode = FIO_CASE_NOMINATIVE
        Case Else
            If IsLikelyBrokenFioText(normalizedCase) Then
                NormalizeFioCaseCode = FIO_CASE_DATIVE
            Else
                NormalizeFioCaseCode = normalizedCase
            End If
    End Select
End Function

Private Function IsLikelyBrokenFioText(ByVal text As String) As Boolean
    Dim normalizedText As String

    normalizedText = Trim$(text)

    If Len(normalizedText) = 0 Then Exit Function

    If InStr(normalizedText, ChrW$(65533)) > 0 Then
        IsLikelyBrokenFioText = True
        Exit Function
    End If

    If InStr(normalizedText, "?") > 0 Or InStr(normalizedText, "N") > 0 Or InStr(normalizedText, "?") > 0 Then
        IsLikelyBrokenFioText = True
    End If
End Function

Private Function BuildShortFioFallback(ByVal sourceText As String) As String
    Dim parts() As String
    Dim normalizedSource As String

    normalizedSource = Application.WorksheetFunction.Trim(CStr(sourceText))
    If Len(normalizedSource) = 0 Then Exit Function

    parts = Split(normalizedSource, " ")

    Select Case UBound(parts)
        Case 0
            BuildShortFioFallback = normalizedSource
        Case 1
            BuildShortFioFallback = parts(1) & " " & Left$(parts(0), 1) & "."
        Case Else
            BuildShortFioFallback = parts(0) & " " & Left$(parts(1), 1) & "." & Left$(parts(2), 1) & "."
    End Select
End Function

Private Function BuildInitialsDisplay(ByVal sourceText As String, ByVal nameCase As String) As String
    Dim surname As String
    Dim firstName As String
    Dim middleName As String
    Dim gender As Long
    Dim declinedSurname As String
    Dim initialsText As String
    Dim normalizedSource As String

    normalizedSource = CleanFioSource(sourceText)
    If Len(normalizedSource) = 0 Then Exit Function

    ParseFioParts normalizedSource, surname, firstName, middleName, gender
    declinedSurname = DeclineSurnameByCase(surname, gender, NormalizeFioCaseCode(nameCase))
    initialsText = BuildInitialsText(firstName, middleName)

    If Len(initialsText) > 0 And Len(declinedSurname) > 0 Then
        BuildInitialsDisplay = initialsText & " " & declinedSurname
    ElseIf Len(declinedSurname) > 0 Then
        BuildInitialsDisplay = declinedSurname
    ElseIf Len(initialsText) > 0 Then
        BuildInitialsDisplay = initialsText
    Else
        BuildInitialsDisplay = BuildShortFioFallback(normalizedSource)
    End If
End Function

Private Function ContainsLetterText(ByVal text As String) As Boolean
    Dim i As Long
    Dim codePoint As Long

    For i = 1 To Len(text)
        codePoint = AscW(Mid$(text, i, 1))
        If (codePoint >= 65 And codePoint <= 90) _
           Or (codePoint >= 97 And codePoint <= 122) _
           Or (codePoint >= 1040 And codePoint <= 1103) _
           Or codePoint = 1025 _
           Or codePoint = 1105 Then
            ContainsLetterText = True
            Exit Function
        End If
    Next i
End Function

Private Function IsUsefulShortFioText(ByVal text As String) As Boolean
    Dim normalizedText As String

    normalizedText = Trim$(text)

    If Len(normalizedText) < 4 Then Exit Function
    If InStr(normalizedText, ".") = 0 Then Exit Function
    If InStr(normalizedText, " ") = 0 Then Exit Function
    If Not ContainsLetterText(normalizedText) Then Exit Function

    IsUsefulShortFioText = True
End Function

Private Function FinalizeFioResult(ByVal resultText As String, ByVal sourceText As String, ByVal shortForm As Boolean) As String
    Dim normalizedSource As String

    normalizedSource = Application.WorksheetFunction.Trim(CStr(sourceText))

    If Len(Trim$(resultText)) = 0 _
       Or IsLikelyBrokenFioText(resultText) _
       Or (shortForm And Not IsUsefulShortFioText(resultText)) _
       Or ((Not shortForm) And Not ContainsLetterText(resultText)) Then
        If shortForm Then
            FinalizeFioResult = BuildShortFioFallback(normalizedSource)
        Else
            FinalizeFioResult = normalizedSource
        End If
    Else
        FinalizeFioResult = Trim$(resultText)
    End If
End Function

' ==========================================================
' FIO ENGINE
' ==========================================================

Private Function CleanFioSource(ByVal sourceText As String) As String
    Dim normalizedText As String

    normalizedText = CStr(sourceText)
    normalizedText = Replace(normalizedText, vbCr, " ")
    normalizedText = Replace(normalizedText, vbLf, " ")
    normalizedText = Replace(normalizedText, vbTab, " ")
    normalizedText = Application.WorksheetFunction.Trim(normalizedText)

    CleanFioSource = normalizedText
End Function

Private Function EndsWithText(ByVal sourceText As String, ByVal suffix As String) As Boolean
    If Len(sourceText) < Len(suffix) Then Exit Function
    EndsWithText = (StrComp(Right$(sourceText, Len(suffix)), suffix, vbTextCompare) = 0)
End Function

Private Function IsOneOf(ByVal sourceText As String, ParamArray values() As Variant) As Boolean
    Dim i As Long

    For i = LBound(values) To UBound(values)
        If StrComp(sourceText, CStr(values(i)), vbTextCompare) = 0 Then
            IsOneOf = True
            Exit Function
        End If
    Next i
End Function

Private Function RemoveLastChars(ByVal sourceText As String, ByVal charCount As Long) As String
    If charCount <= 0 Or Len(sourceText) <= charCount Then Exit Function
    RemoveLastChars = Left$(sourceText, Len(sourceText) - charCount)
End Function

Private Function GetLastLetterText(ByVal sourceText As String) As String
    If Len(sourceText) = 0 Then Exit Function
    GetLastLetterText = Right$(sourceText, 1)
End Function

Private Function HasSurnameAdjectiveEnding(ByVal word As String) As Boolean
    HasSurnameAdjectiveEnding = EndsWithText(word, Ru(1089, 1082, 1080, 1081)) _
        Or EndsWithText(word, Ru(1094, 1082, 1080, 1081)) _
        Or EndsWithText(word, Ru(1089, 1082, 1086, 1081)) _
        Or EndsWithText(word, Ru(1079, 1082, 1086, 1081)) _
        Or EndsWithText(word, Ru(1099, 1081)) _
        Or EndsWithText(word, Ru(1080, 1081)) _
        Or EndsWithText(word, Ru(1086, 1081))
End Function

Private Function IsLikelyIndeclinableSurname(ByVal word As String) As Boolean
    Dim lowerWord As String

    lowerWord = LCase$(word)

    If EndsWithText(lowerWord, Ru(1082, 1086)) _
       Or EndsWithText(lowerWord, Ru(1077, 1085, 1082, 1086)) _
       Or EndsWithText(lowerWord, Ru(1080, 1072)) _
       Or EndsWithText(lowerWord, Ru(1091, 1072)) _
       Or EndsWithText(lowerWord, Ru(1080, 1093)) _
       Or EndsWithText(lowerWord, Ru(1099, 1093)) _
       Or EndsWithText(lowerWord, Ru(1086)) _
       Or EndsWithText(lowerWord, Ru(1077)) _
       Or EndsWithText(lowerWord, Ru(1101)) _
       Or EndsWithText(lowerWord, Ru(1080)) _
       Or EndsWithText(lowerWord, Ru(1091)) _
       Or EndsWithText(lowerWord, Ru(1102)) Then
        IsLikelyIndeclinableSurname = True
    End If
End Function

Private Function RequiresIEnding(ByVal stemText As String) As Boolean
    Dim lastLetter As String

    lastLetter = LCase$(GetLastLetterText(stemText))
    RequiresIEnding = IsOneOf(lastLetter, Ru(1075), Ru(1082), Ru(1093), Ru(1078), Ru(1095), Ru(1096), Ru(1097), Ru(1094))
End Function

Private Function ApplyAEndingGenitive(ByVal word As String) As String
    Dim stemText As String
    Dim lastLetter As String

    lastLetter = LCase$(GetLastLetterText(word))
    stemText = RemoveLastChars(word, 1)

    Select Case lastLetter
        Case Ru(1072)
            If RequiresIEnding(stemText) Then
                ApplyAEndingGenitive = stemText & Ru(1080)
            Else
                ApplyAEndingGenitive = stemText & Ru(1099)
            End If
        Case Ru(1103)
            ApplyAEndingGenitive = stemText & Ru(1080)
        Case Else
            ApplyAEndingGenitive = word
    End Select
End Function

Private Function ApplyAEndingDative(ByVal word As String) As String
    Dim lastLetter As String

    lastLetter = LCase$(GetLastLetterText(word))
    If lastLetter = Ru(1072) Or lastLetter = Ru(1103) Then
        ApplyAEndingDative = RemoveLastChars(word, 1) & Ru(1077)
    Else
        ApplyAEndingDative = word
    End If
End Function

Private Function DetectPatronymicGender(ByVal word As String) As Long
    Dim lowerWord As String

    lowerWord = LCase$(word)

    If EndsWithText(lowerWord, Ru(1086, 1075, 1083, 1099)) Then
        DetectPatronymicGender = FIO_GENDER_MALE
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1082, 1099, 1079, 1099)) _
       Or EndsWithText(lowerWord, Ru(1086, 1074, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1077, 1074, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1105, 1074, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1080, 1095, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1080, 1085, 1080, 1095, 1085, 1072)) Then
        DetectPatronymicGender = FIO_GENDER_FEMALE
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1080, 1095)) Then
        DetectPatronymicGender = FIO_GENDER_MALE
    End If
End Function

Private Function IsLikelyPatronymic(ByVal word As String) As Boolean
    IsLikelyPatronymic = (DetectPatronymicGender(word) <> FIO_GENDER_UNKNOWN)
End Function

Private Function DetectNameGender(ByVal word As String) As Long
    Dim lowerWord As String
    Dim lastLetter As String

    lowerWord = LCase$(word)
    If Len(lowerWord) = 0 Then Exit Function

    Select Case lowerWord
        Case Ru(1080, 1083, 1100, 1103), Ru(1085, 1080, 1082, 1080, 1090, 1072), Ru(1082, 1091, 1079, 1100, 1084, 1072), Ru(1092, 1086, 1084, 1072), Ru(1083, 1091, 1082, 1072), Ru(1083, 1077, 1074, 1072), Ru(1089, 1072, 1074, 1074, 1072)
            DetectNameGender = FIO_GENDER_MALE
            Exit Function
        Case Ru(1083, 1102, 1073, 1086, 1074, 1100), Ru(1085, 1080, 1085, 1077, 1083, 1100)
            DetectNameGender = FIO_GENDER_FEMALE
            Exit Function
    End Select

    lastLetter = GetLastLetterText(lowerWord)

    If lastLetter = Ru(1072) Or lastLetter = Ru(1103) Then
        DetectNameGender = FIO_GENDER_FEMALE
        Exit Function
    End If

    If lastLetter = Ru(1081) Or lastLetter = Ru(1085) Or lastLetter = Ru(1084) Or lastLetter = Ru(1088) _
       Or lastLetter = Ru(1074) Or lastLetter = Ru(1075) Or lastLetter = Ru(1076) Or lastLetter = Ru(1073) _
       Or lastLetter = Ru(1087) Or lastLetter = Ru(1083) Or lastLetter = Ru(1089) Or lastLetter = Ru(1090) _
       Or lastLetter = Ru(1082) Or lastLetter = Ru(1093) Or lastLetter = Ru(1094) Then
        DetectNameGender = FIO_GENDER_MALE
        Exit Function
    End If

    If lastLetter = Ru(1100) Then
        If IsOneOf(lowerWord, Ru(1080, 1075, 1086, 1088, 1100)) Then
            DetectNameGender = FIO_GENDER_MALE
        End If
    End If
End Function

Private Function DetectSurnameGender(ByVal word As String) As Long
    Dim lowerWord As String

    lowerWord = LCase$(word)

    If EndsWithText(lowerWord, Ru(1086, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1077, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1105, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1080, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1099, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1072, 1103)) _
       Or EndsWithText(lowerWord, Ru(1103, 1103)) Then
        DetectSurnameGender = FIO_GENDER_FEMALE
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1086, 1074)) _
       Or EndsWithText(lowerWord, Ru(1077, 1074)) _
       Or EndsWithText(lowerWord, Ru(1105, 1074)) _
       Or EndsWithText(lowerWord, Ru(1080, 1085)) _
       Or EndsWithText(lowerWord, Ru(1099, 1085)) _
       Or HasSurnameAdjectiveEnding(lowerWord) Then
        DetectSurnameGender = FIO_GENDER_MALE
    End If
End Function

Private Function IsLikelyFirstName(ByVal word As String) As Boolean
    IsLikelyFirstName = (DetectNameGender(word) <> FIO_GENDER_UNKNOWN)
End Function

Private Function ResolveFioGender(ByVal firstName As String, ByVal middleName As String, ByVal surname As String) As Long
    ResolveFioGender = DetectPatronymicGender(middleName)
    If ResolveFioGender <> FIO_GENDER_UNKNOWN Then Exit Function

    ResolveFioGender = DetectNameGender(firstName)
    If ResolveFioGender <> FIO_GENDER_UNKNOWN Then Exit Function

    ResolveFioGender = DetectSurnameGender(surname)
End Function

Private Sub ParseFioParts(ByVal sourceText As String, ByRef surname As String, ByRef firstName As String, ByRef middleName As String, ByRef gender As Long)
    Dim normalizedText As String
    Dim parts() As String
    Dim partCount As Long

    normalizedText = CleanFioSource(sourceText)
    If Len(normalizedText) = 0 Then Exit Sub

    parts = Split(normalizedText, " ")
    partCount = UBound(parts) - LBound(parts) + 1

    Select Case partCount
        Case 1
            If IsLikelyFirstName(parts(0)) Then
                firstName = parts(0)
            Else
                surname = parts(0)
            End If
        Case 2
            If IsLikelyFirstName(parts(0)) And Not IsLikelyFirstName(parts(1)) Then
                firstName = parts(0)
                surname = parts(1)
            Else
                surname = parts(0)
                firstName = parts(1)
            End If
        Case Else
            If IsLikelyPatronymic(parts(2)) Then
                surname = parts(0)
                firstName = parts(1)
                middleName = parts(2)
            ElseIf IsLikelyPatronymic(parts(1)) Then
                firstName = parts(0)
                middleName = parts(1)
                surname = parts(2)
            Else
                surname = parts(0)
                firstName = parts(1)
                middleName = parts(2)
            End If
    End Select

    gender = ResolveFioGender(firstName, middleName, surname)
End Sub

Private Function ResolveIrregularFirstName(ByVal lowerWord As String, ByVal caseCode As String) As String
    If caseCode = FIO_CASE_NOMINATIVE Then Exit Function

    If StrComp(lowerWord, Ru(1083, 1077, 1074), vbTextCompare) = 0 Then
        If caseCode = FIO_CASE_DATIVE Then
            ResolveIrregularFirstName = Ru(1083, 1100, 1074, 1091)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            ResolveIrregularFirstName = Ru(1083, 1100, 1074, 1072)
        End If
        Exit Function
    End If

    If StrComp(lowerWord, Ru(1087, 1072, 1074, 1077, 1083), vbTextCompare) = 0 Then
        If caseCode = FIO_CASE_DATIVE Then
            ResolveIrregularFirstName = Ru(1087, 1072, 1074, 1083, 1091)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            ResolveIrregularFirstName = Ru(1087, 1072, 1074, 1083, 1072)
        End If
        Exit Function
    End If
End Function

Private Function DeclineNameByCase(ByVal word As String, ByVal gender As Long, ByVal caseCode As String) As String
    Dim lowerWord As String
    Dim irregularForm As String

    lowerWord = LCase$(word)
    If Len(lowerWord) = 0 Or caseCode = FIO_CASE_NOMINATIVE Then
        DeclineNameByCase = word
        Exit Function
    End If

    irregularForm = ResolveIrregularFirstName(lowerWord, caseCode)
    If Len(irregularForm) > 0 Then
        DeclineNameByCase = UCase$(Left$(word, 1)) & Mid$(irregularForm, 2)
        Exit Function
    End If

    If gender = FIO_GENDER_FEMALE Then
        If EndsWithText(lowerWord, Ru(1072)) Or EndsWithText(lowerWord, Ru(1103)) Then
            If caseCode = FIO_CASE_DATIVE Then
                DeclineNameByCase = ApplyAEndingDative(word)
            ElseIf caseCode = FIO_CASE_GENITIVE Then
                DeclineNameByCase = ApplyAEndingGenitive(word)
            Else
                DeclineNameByCase = word
            End If
            Exit Function
        End If

        If EndsWithText(lowerWord, Ru(1100)) And (caseCode = FIO_CASE_DATIVE Or caseCode = FIO_CASE_GENITIVE) Then
            DeclineNameByCase = RemoveLastChars(word, 1) & Ru(1080)
            Exit Function
        End If

        DeclineNameByCase = word
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1072)) Or EndsWithText(lowerWord, Ru(1103)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclineNameByCase = ApplyAEndingDative(word)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineNameByCase = ApplyAEndingGenitive(word)
        Else
            DeclineNameByCase = word
            End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1081)) Or EndsWithText(lowerWord, Ru(1100)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclineNameByCase = RemoveLastChars(word, 1) & Ru(1102)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineNameByCase = RemoveLastChars(word, 1) & Ru(1103)
        Else
            DeclineNameByCase = word
        End If
        Exit Function
    End If

    If Not IsLikelyFirstName(word) And Not ContainsLetterText(word) Then
        DeclineNameByCase = word
        Exit Function
    End If

    If caseCode = FIO_CASE_DATIVE Then
        DeclineNameByCase = word & Ru(1091)
    ElseIf caseCode = FIO_CASE_GENITIVE Then
        DeclineNameByCase = word & Ru(1072)
    Else
        DeclineNameByCase = word
    End If
End Function

Private Function DeclinePatronymicByCase(ByVal word As String, ByVal gender As Long, ByVal caseCode As String) As String
    Dim lowerWord As String

    lowerWord = LCase$(word)
    If Len(lowerWord) = 0 Or caseCode = FIO_CASE_NOMINATIVE Then
        DeclinePatronymicByCase = word
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1086, 1075, 1083, 1099)) Or EndsWithText(lowerWord, Ru(1082, 1099, 1079, 1099)) Then
        DeclinePatronymicByCase = word
        Exit Function
    End If

    If gender = FIO_GENDER_FEMALE Or EndsWithText(lowerWord, Ru(1085, 1072)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclinePatronymicByCase = ApplyAEndingDative(word)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclinePatronymicByCase = ApplyAEndingGenitive(word)
        Else
            DeclinePatronymicByCase = word
        End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1080, 1095)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclinePatronymicByCase = word & Ru(1091)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclinePatronymicByCase = word & Ru(1072)
        Else
            DeclinePatronymicByCase = word
        End If
        Exit Function
    End If

    DeclinePatronymicByCase = DeclineNameByCase(word, gender, caseCode)
End Function

Private Function DeclineFemaleSurnameByCase(ByVal word As String, ByVal caseCode As String) As String
    Dim lowerWord As String
    Dim stemText As String

    lowerWord = LCase$(word)

    If IsLikelyIndeclinableSurname(lowerWord) Or caseCode = FIO_CASE_NOMINATIVE Then
        DeclineFemaleSurnameByCase = word
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1086, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1077, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1105, 1074, 1072)) _
       Or EndsWithText(lowerWord, Ru(1080, 1085, 1072)) _
       Or EndsWithText(lowerWord, Ru(1099, 1085, 1072)) Then
        DeclineFemaleSurnameByCase = RemoveLastChars(word, 1) & Ru(1086, 1081)
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1072, 1103)) Then
        stemText = RemoveLastChars(word, 2)
        DeclineFemaleSurnameByCase = stemText & Ru(1086, 1081)
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1103, 1103)) Then
        stemText = RemoveLastChars(word, 2)
        DeclineFemaleSurnameByCase = stemText & Ru(1077, 1081)
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1072)) Or EndsWithText(lowerWord, Ru(1103)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclineFemaleSurnameByCase = ApplyAEndingDative(word)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineFemaleSurnameByCase = ApplyAEndingGenitive(word)
        Else
            DeclineFemaleSurnameByCase = word
            End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1100)) And (caseCode = FIO_CASE_DATIVE Or caseCode = FIO_CASE_GENITIVE) Then
        DeclineFemaleSurnameByCase = RemoveLastChars(word, 1) & Ru(1080)
        Exit Function
    End If

    DeclineFemaleSurnameByCase = word
End Function

Private Function DeclineMaleSurnameByCase(ByVal word As String, ByVal caseCode As String) As String
    Dim lowerWord As String
    Dim stemText As String

    lowerWord = LCase$(word)

    If Len(lowerWord) = 0 Or caseCode = FIO_CASE_NOMINATIVE Then
        DeclineMaleSurnameByCase = word
        Exit Function
    End If

    If IsLikelyIndeclinableSurname(lowerWord) Then
        DeclineMaleSurnameByCase = word
        Exit Function
    End If

    If HasSurnameAdjectiveEnding(lowerWord) Then
        stemText = RemoveLastChars(word, 2)
        If caseCode = FIO_CASE_DATIVE Then
            DeclineMaleSurnameByCase = stemText & Ru(1086, 1084, 1091)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineMaleSurnameByCase = stemText & Ru(1086, 1075, 1086)
        Else
            DeclineMaleSurnameByCase = word
        End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1072)) Or EndsWithText(lowerWord, Ru(1103)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclineMaleSurnameByCase = ApplyAEndingDative(word)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineMaleSurnameByCase = ApplyAEndingGenitive(word)
        Else
            DeclineMaleSurnameByCase = word
            End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1081)) Or EndsWithText(lowerWord, Ru(1100)) Then
        If caseCode = FIO_CASE_DATIVE Then
            DeclineMaleSurnameByCase = RemoveLastChars(word, 1) & Ru(1102)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineMaleSurnameByCase = RemoveLastChars(word, 1) & Ru(1103)
        Else
            DeclineMaleSurnameByCase = word
        End If
        Exit Function
    End If

    If EndsWithText(lowerWord, Ru(1077, 1094)) Then
        stemText = RemoveLastChars(word, 2)
        If caseCode = FIO_CASE_DATIVE Then
            DeclineMaleSurnameByCase = stemText & Ru(1094, 1091)
        ElseIf caseCode = FIO_CASE_GENITIVE Then
            DeclineMaleSurnameByCase = stemText & Ru(1094, 1072)
        Else
            DeclineMaleSurnameByCase = word
        End If
        Exit Function
    End If

    If caseCode = FIO_CASE_DATIVE Then
        DeclineMaleSurnameByCase = word & Ru(1091)
    ElseIf caseCode = FIO_CASE_GENITIVE Then
        DeclineMaleSurnameByCase = word & Ru(1072)
    Else
        DeclineMaleSurnameByCase = word
    End If
End Function

Private Function DeclineSurnameByCase(ByVal word As String, ByVal gender As Long, ByVal caseCode As String) As String
    If Len(Trim$(word)) = 0 Then Exit Function

    If gender = FIO_GENDER_FEMALE Then
        DeclineSurnameByCase = DeclineFemaleSurnameByCase(word, caseCode)
    Else
        DeclineSurnameByCase = DeclineMaleSurnameByCase(word, caseCode)
    End If
End Function

Private Function BuildInitialsText(ByVal firstName As String, ByVal middleName As String) As String
    If Len(firstName) > 0 Then
        BuildInitialsText = Left$(firstName, 1) & "."
    End If

    If Len(middleName) > 0 Then
        BuildInitialsText = BuildInitialsText & Left$(middleName, 1) & "."
    End If
End Function

Private Function JoinFioParts(ByVal surname As String, ByVal firstName As String, ByVal middleName As String) As String
    Dim resultText As String

    If Len(surname) > 0 Then resultText = surname
    If Len(firstName) > 0 Then
        If Len(resultText) > 0 Then resultText = resultText & " "
        resultText = resultText & firstName
    End If
    If Len(middleName) > 0 Then
        If Len(resultText) > 0 Then resultText = resultText & " "
        resultText = resultText & middleName
    End If

    JoinFioParts = resultText
End Function

Public Function fio(NameAsText As String, Optional NameCase As String = "D", Optional ShortForm As Boolean = False) As String
    Dim normalizedCase As String
    Dim normalizedSource As String
    Dim surname As String
    Dim firstName As String
    Dim middleName As String
    Dim gender As Long
    Dim declinedSurname As String
    Dim declinedFirstName As String
    Dim declinedMiddleName As String

    normalizedSource = CleanFioSource(NameAsText)
    If Len(normalizedSource) = 0 Then Exit Function

    normalizedCase = NormalizeFioCaseCode(NameCase)
    ParseFioParts normalizedSource, surname, firstName, middleName, gender

    declinedSurname = DeclineSurnameByCase(surname, gender, normalizedCase)
    declinedFirstName = DeclineNameByCase(firstName, gender, normalizedCase)
    declinedMiddleName = DeclinePatronymicByCase(middleName, gender, normalizedCase)

    If ShortForm Then
        fio = JoinFioParts(declinedSurname, BuildInitialsText(firstName, middleName), "")
    Else
        fio = JoinFioParts(declinedSurname, declinedFirstName, declinedMiddleName)
    End If

    fio = FinalizeFioResult(Trim$(fio), normalizedSource, ShortForm)
End Function

' ==========================================================
' REGEXP ENGINE
' ==========================================================

Public Function RegExpExtract(ByVal Text As String, ByVal Pattern As String, Optional ByVal item As Integer = 1, Optional ByVal MatchCase As Boolean = False, Optional ByVal delim As String = ", ") As String
    Dim objRegExp As Object, objMatches As Object, strResult As String, i As Integer
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    If objRegExp Is Nothing Then RegExpExtract = "": Exit Function
    With objRegExp
        .Global = True
        .IgnoreCase = Not MatchCase
        .Pattern = Pattern
    End With
    If objRegExp.Test(Text) Then
        Set objMatches = objRegExp.Execute(Text)
        If item = -1 Then
            For i = 0 To objMatches.count - 1
                If strResult = "" Then strResult = objMatches.item(i).value Else strResult = strResult & delim & objMatches.item(i).value
            Next i
            RegExpExtract = strResult
        ElseIf item > 0 Then
            If objMatches.count >= item Then RegExpExtract = objMatches.item(item - 1).value
        End If
    End If
    Set objRegExp = Nothing
End Function

Public Function RegExpMatch(ByVal Text As String, ByVal Pattern As String) As Boolean
    Dim objRegExp As Object
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    With objRegExp
        .Global = False
        .IgnoreCase = True
        .Pattern = Pattern
    End With
    RegExpMatch = objRegExp.Test(Text)
    Set objRegExp = Nothing
End Function

Public Function RegExpReplace(ByVal Text As String, ByVal Pattern As String, ByVal ReplaceWith As String) As String
    Dim objRegExp As Object
    On Error Resume Next
    Set objRegExp = CreateObject("VBScript.RegExp")
    If objRegExp Is Nothing Then RegExpReplace = Text: Exit Function
    With objRegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = Pattern
    End With
    RegExpReplace = objRegExp.Replace(Text, ReplaceWith)
    Set objRegExp = Nothing
End Function
