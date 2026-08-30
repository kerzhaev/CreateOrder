Attribute VB_Name = "mdlPersonnelOrderText"
Option Explicit

' Pure DTO and text builder shared by personnel preview and Word export.
' It receives already-read dictionaries and never touches Excel, Word, or sheets.

Public Function BuildPersonnelOrderTextModel(ByVal eventData As Object, ByVal beforeState As Object, ByVal afterState As Object) As Object
    Dim model As Object
    Dim transferLines As Collection
    Dim salaryLines As Collection
    Dim exclusionLines As Collection

    Set model = CreateObject("Scripting.Dictionary")
    model.CompareMode = vbTextCompare
    model.Add "event_type", TextValue(eventData, "event_type")
    model.Add "employee_id", TextValue(eventData, "employee_id")
    model.Add "fio", TextValue(beforeState, "fio")
    model.Add "personal_number", TextValue(beforeState, "personal_number")
    model.Add "employee_number_suffix", EmployeeNumberSuffix(TextValue(beforeState, "personal_number"))
    model.Add "rank_before", TextValue(beforeState, "rank")
    model.Add "rank_after", TextValue(afterState, "rank")
    model.Add "appointment_before", AppointmentText(beforeState)
    model.Add "appointment_after", AppointmentText(afterState)
    model.Add "vus_before_suffix", VusSuffix(beforeState)
    model.Add "vus_after_suffix", VusSuffix(afterState)
    model.Add "event_date", ValueOf(eventData, "event_date")
    model.Add "effective_date", ValueOf(eventData, "effective_date")
    model.Add "handover_date", EventDateOrFallback(eventData, "handover_date", "event_date")
    model.Add "acceptance_date", EventDateOrFallback(eventData, "acceptance_date", "effective_date")
    model.Add "duty_start_date", EventDateOrFallback(eventData, "duty_start_date", "effective_date")
    model.Add "order_reference", TextValue(eventData, "order_reference")
    model.Add "basis_text", TextValue(eventData, "basis_text")
    model.Add "destination_unit", TextValue(eventData, "destination_unit")
    model.Add "destination_location", TextValue(eventData, "destination_location")
    model.Add "destination_text", DestinationText(eventData)
    model.Add "material_assistance_status", TextValue(eventData, "material_assistance_status")
    model.Add "main_leave_status", TextValue(eventData, "main_leave_status")
    model.Add "additional_leave_status", TextValue(eventData, "additional_leave_status")
    model.Add "position_salary", TextValue(afterState, "position_salary")
    model.Add "tariff_rank", TextValue(afterState, "tariff_rank")
    model.Add "rank_salary", TextValue(afterState, "rank_salary")
    model.Add "salary_date", EventDateOrFallback(eventData, "duty_start_date", "effective_date")
    model.Add "section_heading", "§ 1"
    model.Add "basis_line", BasisLine(TextValue(eventData, "order_reference"), TextValue(eventData, "basis_text"))

    Set transferLines = BuildTransferLines(model)
    Set salaryLines = BuildSalaryLines(model)
    Set exclusionLines = BuildExclusionServiceLines(model)
    model.Add "transfer_lines", transferLines
    model.Add "salary_lines", salaryLines
    model.Add "exclusion_service_lines", exclusionLines
    model.Add "transfer_core_text", TransferCoreText(model)
    model.Add "exclusion_core_text", ExclusionCoreText(model)
    model.Add "exclusion_stop_text", "Прекратить с даты исключения выплату ранее установленных надбавок и повышающих коэффициентов."
    Set BuildPersonnelOrderTextModel = model
End Function

Private Function TransferCoreText(ByVal model As Object) As String
    TransferCoreText = TextValue(model, "rank_after") & " " & TextValue(model, "fio") & TextValue(model, "employee_number_suffix") & _
        ", ранее замещавшего " & TextValue(model, "appointment_before") & TextValue(model, "vus_before_suffix") & _
        ", назначенного приказом " & TextValue(model, "order_reference") & " на " & TextValue(model, "appointment_after") & TextValue(model, "vus_after_suffix") & "."
End Function

Private Function ExclusionCoreText(ByVal model As Object) As String
    ExclusionCoreText = TextValue(model, "rank_before") & " " & TextValue(model, "fio") & TextValue(model, "employee_number_suffix") & _
        ", замещавшего " & TextValue(model, "appointment_before") & TextValue(model, "vus_before_suffix") & _
        ", с " & FormatModelDate(ValueOf(model, "handover_date")) & " дела и должность сдал; с " & _
        FormatModelDate(ValueOf(model, "effective_date")) & " исключить из списков личного состава воинской части и всех видов обеспечения" & TextValue(model, "destination_text") & "."
End Function

Private Function BuildTransferLines(ByVal model As Object) As Collection
    Dim result As New Collection
    result.Add "С " & FormatModelDate(ValueOf(model, "handover_date")) & " дела и должность по прежней воинской должности сдал."
    result.Add "С " & FormatModelDate(ValueOf(model, "acceptance_date")) & " дела и должность по новой воинской должности принял."
    result.Add "К исполнению обязанностей по новой воинской должности приступил с " & FormatModelDate(ValueOf(model, "duty_start_date")) & "."
    Set BuildTransferLines = result
End Function

Private Function BuildSalaryLines(ByVal model As Object) As Collection
    Dim result As New Collection
    Dim salaryText As String
    Dim salaryDate As String
    salaryDate = FormatModelDate(ValueOf(model, "salary_date"))
    If TextValue(model, "position_salary") <> "" Then
        salaryText = "С " & salaryDate & " установить оклад по воинской должности в размере " & TextValue(model, "position_salary") & " руб. в месяц"
        If TextValue(model, "tariff_rank") <> "" Then salaryText = salaryText & " (" & TextValue(model, "tariff_rank") & " тарифный разряд)"
        result.Add salaryText & "."
    ElseIf TextValue(model, "tariff_rank") <> "" Then
        result.Add "С " & salaryDate & " установить оклад по " & TextValue(model, "tariff_rank") & " тарифному разряду."
    End If
    If TextValue(model, "rank_salary") <> "" Then
        result.Add "С " & salaryDate & " установить оклад по воинскому званию в размере " & TextValue(model, "rank_salary") & " руб. в месяц."
    End If
    Set BuildSalaryLines = result
End Function

Private Function BuildExclusionServiceLines(ByVal model As Object) As Collection
    Dim result As New Collection
    If TextValue(model, "material_assistance_status") <> "" Then result.Add "Материальная помощь за текущий год: " & TextValue(model, "material_assistance_status") & "."
    If TextValue(model, "main_leave_status") <> "" Then result.Add "Основной отпуск за текущий год: " & TextValue(model, "main_leave_status") & "."
    If TextValue(model, "additional_leave_status") <> "" Then result.Add "Дополнительный отпуск за текущий год: " & TextValue(model, "additional_leave_status") & "."
    Set BuildExclusionServiceLines = result
End Function

Private Function AppointmentText(ByVal stateData As Object) As String
    AppointmentText = "воинскую должность: " & TextValue(stateData, "position")
    If TextValue(stateData, "section") <> "" Then AppointmentText = AppointmentText & "; подразделение: " & TextValue(stateData, "section")
    If TextValue(stateData, "military_unit") <> "" Then AppointmentText = AppointmentText & "; воинская часть: " & TextValue(stateData, "military_unit")
End Function

Private Function VusSuffix(ByVal stateData As Object) As String
    If TextValue(stateData, "vus") <> "" Then VusSuffix = "; ВУС-" & TextValue(stateData, "vus")
End Function

Private Function EmployeeNumberSuffix(ByVal personalNumber As String) As String
    If Trim$(personalNumber) <> "" Then EmployeeNumberSuffix = ", личный номер " & Trim$(personalNumber)
End Function

Private Function DestinationText(ByVal eventData As Object) As String
    If TextValue(eventData, "destination_unit") <> "" Then DestinationText = ", полагать убывшим к новому месту службы в " & TextValue(eventData, "destination_unit")
    If TextValue(eventData, "destination_location") <> "" Then DestinationText = DestinationText & ", " & TextValue(eventData, "destination_location")
End Function

Private Function BasisLine(ByVal orderReference As String, ByVal basisText As String) As String
    BasisLine = "ОСНОВАНИЕ: " & orderReference
    If Trim$(basisText) <> "" Then BasisLine = BasisLine & "; " & basisText
End Function

Private Function EventDateOrFallback(ByVal eventData As Object, ByVal key As String, ByVal fallbackKey As String) As Variant
    If HasKey(eventData, key) Then
        If IsDate(eventData(key)) Then
            EventDateOrFallback = eventData(key)
            Exit Function
        End If
    End If
    EventDateOrFallback = ValueOf(eventData, fallbackKey)
End Function

Private Function FormatModelDate(ByVal rawValue As Variant) As String
    If IsDate(rawValue) Then FormatModelDate = Format$(CDate(rawValue), "dd.mm.yyyy") Else FormatModelDate = TextValueValue(rawValue)
End Function

Private Function ValueOf(ByVal source As Object, ByVal key As String) As Variant
    If source Is Nothing Then Exit Function
    On Error GoTo Missing
    If source.Exists(key) Then ValueOf = source(key)
Missing:
End Function

Private Function HasKey(ByVal source As Object, ByVal key As String) As Boolean
    If source Is Nothing Then Exit Function
    On Error GoTo Missing
    HasKey = source.Exists(key)
Missing:
End Function

Private Function TextValue(ByVal source As Object, ByVal key As String) As String
    TextValue = TextValueValue(ValueOf(source, key))
End Function

Private Function TextValueValue(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsNull(rawValue) Or IsEmpty(rawValue) Then Exit Function
    TextValueValue = Trim$(CStr(rawValue))
End Function
