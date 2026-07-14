Attribute VB_Name = "mdlEnrollmentWorkflow"
Option Explicit

Private Const STATUS_READY As Long = 0
Private Const STATUS_WARNING As Long = 1
Private Const STATUS_BLOCKED As Long = 2

Private Const YES_VALUE As String = "YES"
Private Const NO_VALUE As String = "NO"

Private Const DEFAULT_PREMIUM_PERCENT As String = "25"
Private Const DEFAULT_POSITION_ALLOWANCE_PERCENT As String = "100"
Private Const DEFAULT_SPECIAL_CONDITIONS_PERCENT As String = "50"
Private Const DEFAULT_TARIFF_PERCENT As String = "50"
Private Const DEFAULT_CONTRACT430_PERCENT As String = "60"
Private Const DEFAULT_PREFERENTIAL_COEFF As String = "1.5"
Private Const DEFAULT_EDV_AMOUNT As String = "400000"
Private Const DEFAULT_PER_DIEM_DAYS As String = "1"

Private enrollmentReferencesSynced As Boolean

Private Const BACKEND_COL_KEY As Long = 1
Private Const BACKEND_COL_LABEL As Long = 2
Private Const BACKEND_COL_VALUE As Long = 3
Private Const BACKEND_COL_DERIVED As Long = 4

Private Const BACKEND_HEADER_ROW As Long = 1
Private Const BACKEND_FIRST_DATA_ROW As Long = 2

Public Const COL_ENROLLMENT_ID As Long = 1
Public Const COL_ENROLLMENT_FIO As Long = 2
Public Const COL_ENROLLMENT_LICHNIY_NOMER As Long = 3
Public Const COL_ENROLLMENT_RANK As Long = 4
Public Const COL_ENROLLMENT_POSITION As Long = 5
Public Const COL_ENROLLMENT_SECTION As Long = 6
Public Const COL_ENROLLMENT_ORDER_DATE As Long = 7
Public Const COL_ENROLLMENT_ORDER_NUMBER As Long = 8
Public Const COL_ENROLLMENT_ACHIEVEMENT_PARAM As Long = 9
Public Const COL_ENROLLMENT_ACCEPT_DATE As Long = 10
Public Const COL_ENROLLMENT_ENROLL_DATE As Long = 11
Public Const COL_ENROLLMENT_MANUAL_START_DATE As Long = 12
Public Const COL_ENROLLMENT_REPORT_INFO As Long = 13
Public Const COL_ENROLLMENT_ASSIGNMENT_INFO As Long = 14
Public Const COL_ENROLLMENT_CLASS_PARAM As Long = 15
Public Const COL_ENROLLMENT_FIZO_PARAM As Long = 16
Public Const COL_ENROLLMENT_SECRECY_PARAM As Long = 17
Public Const COL_ENROLLMENT_STANDARD_TYPES As Long = 18
Public Const COL_ENROLLMENT_PAYMENT_BASIS As Long = 19
Public Const COL_ENROLLMENT_STATUS As Long = 20
Public Const COL_ENROLLMENT_COMMENT As Long = 21
Public Const COL_ENROLLMENT_ORDER_DRAFT_ID As Long = 22
Public Const COL_ENROLLMENT_WORD_READY As Long = 23
Public Const COL_ENROLLMENT_VALIDATION_SEVERITY As Long = 24
Public Const COL_ENROLLMENT_VALIDATION_ISSUES As Long = 25
Public Const COL_ENROLLMENT_SOURCE_MODE As Long = 26
Public Const COL_ENROLLMENT_LAST_DERIVED_AT As Long = 27
Public Const COL_ENROLLMENT_TABLE_NUMBER As Long = 28
Public Const COL_ENROLLMENT_SERVICE_CATEGORY As Long = 29
Public Const COL_ENROLLMENT_CONTRACT_KIND As Long = 30
Public Const COL_ENROLLMENT_CONTRACT_BASIS As Long = 31
Public Const COL_ENROLLMENT_VUS As Long = 32
Public Const COL_ENROLLMENT_MILITARY_UNIT As Long = 33
Public Const COL_ENROLLMENT_TARIFF_RANK As Long = 34
Public Const COL_ENROLLMENT_POSITION_SALARY As Long = 35
Public Const COL_ENROLLMENT_RANK_SALARY As Long = 36
Public Const COL_ENROLLMENT_ORDER_ISSUER As Long = 37
Public Const COL_ENROLLMENT_ARRIVAL_SOURCE As Long = 38
Public Const COL_ENROLLMENT_PRESCRIPTION_NUMBER As Long = 39
Public Const COL_ENROLLMENT_PRESCRIPTION_DATE As Long = 40
Public Const COL_ENROLLMENT_REPORT_NUMBER As Long = 41
Public Const COL_ENROLLMENT_REPORT_DATE As Long = 42
Public Const COL_ENROLLMENT_DUTY_START_DATE As Long = 43
Public Const COL_ENROLLMENT_STANDARD_START_DATE As Long = 44
Public Const COL_ENROLLMENT_PREFERENTIAL_START_DATE As Long = 45
Public Const COL_ENROLLMENT_BASIS_SECTION1 As Long = 46
Public Const COL_ENROLLMENT_BASIS_SECTION2 As Long = 47
Public Const COL_ENROLLMENT_BIRTH_DATE As Long = 48
Public Const COL_ENROLLMENT_BIRTH_PLACE As Long = 49
Public Const COL_ENROLLMENT_CITIZENSHIP As Long = 50
Public Const COL_ENROLLMENT_INN As Long = 51
Public Const COL_ENROLLMENT_SNILS As Long = 52
Public Const COL_ENROLLMENT_PASSPORT_SERIES As Long = 53
Public Const COL_ENROLLMENT_PASSPORT_NUMBER As Long = 54
Public Const COL_ENROLLMENT_PASSPORT_ISSUER As Long = 55
Public Const COL_ENROLLMENT_PASSPORT_ISSUE_DATE As Long = 56
Public Const COL_ENROLLMENT_PASSPORT_CODE As Long = 57
Public Const COL_ENROLLMENT_BANK_ACCOUNT As Long = 58
Public Const COL_ENROLLMENT_BANK_NAME As Long = 59
Public Const COL_ENROLLMENT_REQUISITES_NOTE As Long = 60
Public Const COL_ENROLLMENT_PREFERENTIAL_ENABLED As Long = 61
Public Const COL_ENROLLMENT_PREFERENTIAL_COEFF As Long = 62
Public Const COL_ENROLLMENT_PREFERENTIAL_BASIS As Long = 63
Public Const COL_ENROLLMENT_PREMIUM_ENABLED As Long = 64
Public Const COL_ENROLLMENT_PREMIUM_PERCENT As Long = 65
Public Const COL_ENROLLMENT_PREMIUM_START As Long = 66
Public Const COL_ENROLLMENT_PREMIUM_END As Long = 67
Public Const COL_ENROLLMENT_PREMIUM_BASIS As Long = 68
Public Const COL_ENROLLMENT_LIFT_ENABLED As Long = 69
Public Const COL_ENROLLMENT_LIFT_AMOUNT As Long = 70
Public Const COL_ENROLLMENT_LIFT_DATE As Long = 71
Public Const COL_ENROLLMENT_LIFT_BASIS As Long = 72
Public Const COL_ENROLLMENT_PER_DIEM_ENABLED As Long = 73
Public Const COL_ENROLLMENT_PER_DIEM_DAYS As Long = 74
Public Const COL_ENROLLMENT_PER_DIEM_AMOUNT As Long = 75
Public Const COL_ENROLLMENT_PER_DIEM_DATE As Long = 76
Public Const COL_ENROLLMENT_PER_DIEM_BASIS As Long = 77
Public Const COL_ENROLLMENT_EDV_ENABLED As Long = 78
Public Const COL_ENROLLMENT_EDV_AMOUNT As Long = 79
Public Const COL_ENROLLMENT_EDV_DATE As Long = 80
Public Const COL_ENROLLMENT_EDV_BASIS As Long = 81
Public Const COL_ENROLLMENT_CLASS_ENABLED As Long = 82
Public Const COL_ENROLLMENT_CLASS_PERCENT As Long = 83
Public Const COL_ENROLLMENT_CLASS_BASIS As Long = 84
Public Const COL_ENROLLMENT_FIZO_ENABLED As Long = 85
Public Const COL_ENROLLMENT_FIZO_PERCENT As Long = 86
Public Const COL_ENROLLMENT_FIZO_BASIS As Long = 87
Public Const COL_ENROLLMENT_SECRECY_ENABLED As Long = 88
Public Const COL_ENROLLMENT_SECRECY_PERCENT As Long = 89
Public Const COL_ENROLLMENT_SECRECY_BASIS As Long = 90
Public Const COL_ENROLLMENT_ACHIEVEMENT_ENABLED As Long = 91
Public Const COL_ENROLLMENT_ACHIEVEMENT_AMOUNT As Long = 92
Public Const COL_ENROLLMENT_ACHIEVEMENT_BASIS As Long = 93
Public Const COL_ENROLLMENT_STD_DUTY_ENABLED As Long = 94
Public Const COL_ENROLLMENT_STD_DUTY_PERCENT As Long = 95
Public Const COL_ENROLLMENT_STD_DUTY_DATE As Long = 96
Public Const COL_ENROLLMENT_STD_DUTY_BASIS As Long = 97
Public Const COL_ENROLLMENT_STD_SPECIAL_ENABLED As Long = 98
Public Const COL_ENROLLMENT_STD_SPECIAL_PERCENT As Long = 99
Public Const COL_ENROLLMENT_STD_SPECIAL_DATE As Long = 100
Public Const COL_ENROLLMENT_STD_SPECIAL_BASIS As Long = 101
Public Const COL_ENROLLMENT_STD_TARIFF_ENABLED As Long = 102
Public Const COL_ENROLLMENT_STD_TARIFF_PERCENT As Long = 103
Public Const COL_ENROLLMENT_STD_TARIFF_DATE As Long = 104
Public Const COL_ENROLLMENT_STD_TARIFF_BASIS As Long = 105
Public Const COL_ENROLLMENT_STD_CONTRACT430_ENABLED As Long = 106
Public Const COL_ENROLLMENT_STD_CONTRACT430_PERCENT As Long = 107
Public Const COL_ENROLLMENT_STD_CONTRACT430_DATE As Long = 108
Public Const COL_ENROLLMENT_STD_CONTRACT430_BASIS As Long = 109
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_NAME As Long = 110
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_PARAM As Long = 111
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_AMOUNT As Long = 112
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_START As Long = 113
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_BASIS As Long = 114
Public Const COL_ENROLLMENT_EXTRA_MONTHLY1_ENABLED As Long = 115
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_NAME As Long = 116
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_PARAM As Long = 117
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_AMOUNT As Long = 118
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_START As Long = 119
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_BASIS As Long = 120
Public Const COL_ENROLLMENT_EXTRA_MONTHLY2_ENABLED As Long = 121
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_NAME As Long = 122
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_PARAM As Long = 123
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_AMOUNT As Long = 124
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_START As Long = 125
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_BASIS As Long = 126
Public Const COL_ENROLLMENT_EXTRA_MONTHLY3_ENABLED As Long = 127
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_NAME As Long = 128
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_PARAM As Long = 129
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_AMOUNT As Long = 130
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_START As Long = 131
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_BASIS As Long = 132
Public Const COL_ENROLLMENT_EXTRA_MONTHLY4_ENABLED As Long = 133
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME1_NAME As Long = 134
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME1_AMOUNT As Long = 135
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME1_DATE As Long = 136
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME1_BASIS As Long = 137
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME1_ENABLED As Long = 138
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME2_NAME As Long = 139
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME2_AMOUNT As Long = 140
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME2_DATE As Long = 141
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME2_BASIS As Long = 142
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME2_ENABLED As Long = 143
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME3_NAME As Long = 144
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME3_AMOUNT As Long = 145
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME3_DATE As Long = 146
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME3_BASIS As Long = 147
Public Const COL_ENROLLMENT_EXTRA_ONE_TIME3_ENABLED As Long = 148
Public Const COL_ENROLLMENT_ACHIEVEMENT_AWARD_DATE As Long = 149
Public Const COL_ENROLLMENT_ACHIEVEMENT_DOCUMENT_REFERENCE As Long = 150

Public Const ENROLLMENT_REFERENCE_SHEET As String = "EnrollmentReferenceData"

Public Sub EnsureEnrollmentInfrastructure()
    Dim wsJournal As Worksheet
    Dim wsBackend As Worksheet

    Set wsJournal = EnsureWorksheet(mdlReferenceData.SHEET_ENROLLMENT)
    Set wsBackend = EnsureWorksheet(mdlReferenceData.SHEET_ENROLLMENT_FORM)

    EnsureEnrollmentSheetStructure wsJournal
    EnsureEnrollmentFormSheetStructure wsBackend
    EnsureEnrollmentSettings
    EnsureEnrollmentReferenceData
    mdlEnrollmentOrderExport.EnsureEnrollmentTemplateAvailable

    On Error Resume Next
    wsBackend.Visible = xlSheetVeryHidden
    On Error GoTo 0
    HideLegacyEnrollmentFormSheets
End Sub

Private Sub HideLegacyEnrollmentFormSheets()
    Dim legacyName As String
    Dim ws As Worksheet

    legacyName = mdlHelper.Ru(1052, 1072, 1089, 1090, 1077, 1088, 95, 1047, 1072, 1095, 1080, 1089, 1083, 1077, 1085, 1080, 1103)
    If StrComp(legacyName, mdlReferenceData.SHEET_ENROLLMENT_FORM, vbTextCompare) = 0 Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(legacyName)
    If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Sub EnsureEnrollmentSheetStructure(ByVal ws As Worksheet)
    SetEnrollmentHeader ws, COL_ENROLLMENT_ID, "enrollment.header.id", "ID зачисления", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_FIO, "enrollment.header.fio", "ФИО", 32
    SetEnrollmentHeader ws, COL_ENROLLMENT_LICHNIY_NOMER, "enrollment.header.personal_number", "Личный номер", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_RANK, "enrollment.header.rank", "Воинское звание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_POSITION, "enrollment.header.position", "Штатная должность", 34
    SetEnrollmentHeader ws, COL_ENROLLMENT_SECTION, "enrollment.header.section", "Подразделение / раздел персонала", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_ORDER_DATE, "enrollment.header.order_date", "Дата приказа", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_ORDER_NUMBER, "enrollment.header.order_number", "Номер приказа", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_PARAM, "enrollment.header.achievement_param", "Особое достижение / медаль", 22
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACCEPT_DATE, "enrollment.header.accept_date", "Дата принятия дел и должности", 20
    SetEnrollmentHeader ws, COL_ENROLLMENT_ENROLL_DATE, "enrollment.header.enroll_date", "Дата зачисления", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_MANUAL_START_DATE, "enrollment.header.manual_start_date", "Ручная дата старта", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_REPORT_INFO, "enrollment.header.report_info", "Рапорт / регистрация", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_ASSIGNMENT_INFO, "enrollment.header.assignment_info", "Предписание / основание", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_CLASS_PARAM, "enrollment.header.class_param", "Классность", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_FIZO_PARAM, "enrollment.header.fizo_param", "ФИЗО", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_SECRECY_PARAM, "enrollment.header.secrecy_param", "Секретность", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_STANDARD_TYPES, "enrollment.header.standard_types", "Стандартные надбавки", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_PAYMENT_BASIS, "enrollment.header.payment_basis", "Основание проекта приказа", 36
    SetEnrollmentHeader ws, COL_ENROLLMENT_STATUS, "enrollment.header.status", "Статус", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_COMMENT, "enrollment.header.comment", "Комментарий", 40
    SetEnrollmentHeader ws, COL_ENROLLMENT_ORDER_DRAFT_ID, "enrollment.header.order_draft_id", "ID проекта приказа", 20
    SetEnrollmentHeader ws, COL_ENROLLMENT_WORD_READY, "enrollment.header.word_ready", "Готовность Word", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_VALIDATION_SEVERITY, "enrollment.header.validation_severity", "Критичность проверки", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_VALIDATION_ISSUES, "enrollment.header.validation_issues", "Замечания проверки", 40
    SetEnrollmentHeader ws, COL_ENROLLMENT_SOURCE_MODE, "enrollment.header.source_mode", "Источник данных", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_LAST_DERIVED_AT, "enrollment.header.last_derived_at", "Последний пересчёт", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_TABLE_NUMBER, "enrollment.header.table_number", "Табельный номер", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_SERVICE_CATEGORY, "enrollment.header.service_category", "Категория службы", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_CONTRACT_KIND, "enrollment.header.contract_kind", "Признак контракта", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_CONTRACT_BASIS, "enrollment.header.contract_basis", "Основание контракта", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_VUS, "enrollment.header.vus", "ВУС", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_MILITARY_UNIT, "enrollment.header.military_unit", "Воинская часть", 20
    SetEnrollmentHeader ws, COL_ENROLLMENT_TARIFF_RANK, "enrollment.header.tariff_rank", "Тарифный разряд", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_POSITION_SALARY, "enrollment.header.position_salary", "Оклад по должности", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_RANK_SALARY, "enrollment.header.rank_salary", "Оклад по званию", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_ORDER_ISSUER, "enrollment.header.order_issuer", "Кем издан приказ", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_ARRIVAL_SOURCE, "enrollment.header.arrival_source", "Пункт отбора / источник прибытия", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_PRESCRIPTION_NUMBER, "enrollment.header.prescription_number", "Номер предписания", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_PRESCRIPTION_DATE, "enrollment.header.prescription_date", "Дата предписания", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_REPORT_NUMBER, "enrollment.header.report_number", "Номер рапорта", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_REPORT_DATE, "enrollment.header.report_date", "Дата рапорта", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_DUTY_START_DATE, "enrollment.header.duty_start_date", "Дата вступления в исполнение", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_STANDARD_START_DATE, "enrollment.header.standard_start_date", "Дата старта стандартных выплат", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREFERENTIAL_START_DATE, "enrollment.header.preferential_start_date", "Дата льготной выслуги", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_BASIS_SECTION1, "enrollment.header.basis_section1", "Основание для §1", 38
    SetEnrollmentHeader ws, COL_ENROLLMENT_BASIS_SECTION2, "enrollment.header.basis_section2", "Основание для §2", 38
    SetEnrollmentHeader ws, COL_ENROLLMENT_BIRTH_DATE, "enrollment.header.birth_date", "Дата рождения", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_BIRTH_PLACE, "enrollment.header.birth_place", "Место рождения", 22
    SetEnrollmentHeader ws, COL_ENROLLMENT_CITIZENSHIP, "enrollment.header.citizenship", "Гражданство", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_INN, "enrollment.header.inn", "ИНН", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_SNILS, "enrollment.header.snils", "СНИЛС", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_PASSPORT_SERIES, "enrollment.header.passport_series", "Серия паспорта", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PASSPORT_NUMBER, "enrollment.header.passport_number", "Номер паспорта", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PASSPORT_ISSUER, "enrollment.header.passport_issuer", "Кем выдан", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_PASSPORT_ISSUE_DATE, "enrollment.header.passport_issue_date", "Дата выдачи", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PASSPORT_CODE, "enrollment.header.passport_code", "Код подразделения", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_BANK_ACCOUNT, "enrollment.header.bank_account", "Лицевой / банковский счёт", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_BANK_NAME, "enrollment.header.bank_name", "Банк", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_REQUISITES_NOTE, "enrollment.header.requisites_note", "Примечание по реквизитам", 26
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREFERENTIAL_ENABLED, "enrollment.header.preferential_enabled", "Льготная выслуга вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREFERENTIAL_COEFF, "enrollment.header.preferential_coeff", "Коэффициент льготы", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREFERENTIAL_BASIS, "enrollment.header.preferential_basis", "Основание льготной выслуги", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREMIUM_ENABLED, "enrollment.header.premium_enabled", "Премия вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREMIUM_PERCENT, "enrollment.header.premium_percent", "Премия %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREMIUM_START, "enrollment.header.premium_start", "Начало премии", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREMIUM_END, "enrollment.header.premium_end", "Окончание премии", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PREMIUM_BASIS, "enrollment.header.premium_basis", "Основание премии", 26
    SetEnrollmentHeader ws, COL_ENROLLMENT_LIFT_ENABLED, "enrollment.header.lift_enabled", "Подъёмное вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_LIFT_AMOUNT, "enrollment.header.lift_amount", "Подъёмное", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_LIFT_DATE, "enrollment.header.lift_date", "Дата подъёмного", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_LIFT_BASIS, "enrollment.header.lift_basis", "Основание подъёмного", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_PER_DIEM_ENABLED, "enrollment.header.per_diem_enabled", "Суточные вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_PER_DIEM_DAYS, "enrollment.header.per_diem_days", "Суточные дней", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_PER_DIEM_AMOUNT, "enrollment.header.per_diem_amount", "Размер суточных", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_PER_DIEM_DATE, "enrollment.header.per_diem_date", "Дата суточных", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_PER_DIEM_BASIS, "enrollment.header.per_diem_basis", "Основание суточных", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_EDV_ENABLED, "enrollment.header.edv_enabled", "ЕДВ 400000 вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_EDV_AMOUNT, "enrollment.header.edv_amount", "Сумма ЕДВ", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_EDV_DATE, "enrollment.header.edv_date", "Дата ЕДВ", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_EDV_BASIS, "enrollment.header.edv_basis", "Основание ЕДВ", 26
    SetEnrollmentHeader ws, COL_ENROLLMENT_CLASS_ENABLED, "enrollment.header.class_enabled", "Классность вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_CLASS_PERCENT, "enrollment.header.class_percent", "Классность %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_CLASS_BASIS, "enrollment.header.class_basis", "Основание классности", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_FIZO_ENABLED, "enrollment.header.fizo_enabled", "ФИЗО вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_FIZO_PERCENT, "enrollment.header.fizo_percent", "ФИЗО %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_FIZO_BASIS, "enrollment.header.fizo_basis", "Основание ФИЗО", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_SECRECY_ENABLED, "enrollment.header.secrecy_enabled", "Секретность вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_SECRECY_PERCENT, "enrollment.header.secrecy_percent", "Секретность %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_SECRECY_BASIS, "enrollment.header.secrecy_basis", "Основание секретности", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_ENABLED, "enrollment.header.achievement_enabled", "Особые достижения вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_AMOUNT, "enrollment.header.achievement_amount", "Особое достижение %/сумма", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_BASIS, "enrollment.header.achievement_basis", "Основание особых достижений", 26
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_AWARD_DATE, "common.date", "Award order date", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_ACHIEVEMENT_DOCUMENT_REFERENCE, "enrollment.header.order_number", "Award order reference", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_DUTY_ENABLED, "enrollment.header.std_duty_enabled", "Надбавка по должности вкл", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_DUTY_PERCENT, "enrollment.header.std_duty_percent", "Надбавка по должности %", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_DUTY_DATE, "enrollment.header.std_duty_date", "Дата надбавки по должности", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_DUTY_BASIS, "enrollment.header.std_duty_basis", "Основание надбавки по должности", 28
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_SPECIAL_ENABLED, "enrollment.header.std_special_enabled", "Особые условия вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_SPECIAL_PERCENT, "enrollment.header.std_special_percent", "Особые условия %", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_SPECIAL_DATE, "enrollment.header.std_special_date", "Дата особых условий", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_SPECIAL_BASIS, "enrollment.header.std_special_basis", "Основание особых условий", 26
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_TARIFF_ENABLED, "enrollment.header.std_tariff_enabled", "1-4 тариф вкл", 14
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_TARIFF_PERCENT, "enrollment.header.std_tariff_percent", "1-4 тариф %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_TARIFF_DATE, "enrollment.header.std_tariff_date", "Дата 1-4 тарифа", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_TARIFF_BASIS, "enrollment.header.std_tariff_basis", "Основание 1-4 тарифа", 24
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_CONTRACT430_ENABLED, "enrollment.header.std_contract430_enabled", "430 приказ вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_CONTRACT430_PERCENT, "enrollment.header.std_contract430_percent", "430 приказ %", 12
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_CONTRACT430_DATE, "enrollment.header.std_contract430_date", "Дата 430 приказа", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_STD_CONTRACT430_BASIS, "enrollment.header.std_contract430_basis", "Основание 430 приказа", 24

    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_NAME, "enrollment.header.extra_monthly1_name", "Доп. ежемесячная 1: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_PARAM, "enrollment.header.extra_monthly1_param", "Доп. ежемесячная 1: параметр", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_AMOUNT, "enrollment.header.extra_monthly1_amount", "Доп. ежемесячная 1: размер", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_START, "enrollment.header.extra_monthly1_start", "Доп. ежемесячная 1: дата начала", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_BASIS, "enrollment.header.extra_monthly1_basis", "Доп. ежемесячная 1: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY1_ENABLED, "enrollment.header.extra_monthly1_enabled", "Доп. ежемесячная 1: вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_NAME, "enrollment.header.extra_monthly2_name", "Доп. ежемесячная 2: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_PARAM, "enrollment.header.extra_monthly2_param", "Доп. ежемесячная 2: параметр", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_AMOUNT, "enrollment.header.extra_monthly2_amount", "Доп. ежемесячная 2: размер", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_START, "enrollment.header.extra_monthly2_start", "Доп. ежемесячная 2: дата начала", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_BASIS, "enrollment.header.extra_monthly2_basis", "Доп. ежемесячная 2: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY2_ENABLED, "enrollment.header.extra_monthly2_enabled", "Доп. ежемесячная 2: вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_NAME, "enrollment.header.extra_monthly3_name", "Доп. ежемесячная 3: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_PARAM, "enrollment.header.extra_monthly3_param", "Доп. ежемесячная 3: параметр", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_AMOUNT, "enrollment.header.extra_monthly3_amount", "Доп. ежемесячная 3: размер", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_START, "enrollment.header.extra_monthly3_start", "Доп. ежемесячная 3: дата начала", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_BASIS, "enrollment.header.extra_monthly3_basis", "Доп. ежемесячная 3: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY3_ENABLED, "enrollment.header.extra_monthly3_enabled", "Доп. ежемесячная 3: вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_NAME, "enrollment.header.extra_monthly4_name", "Доп. ежемесячная 4: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_PARAM, "enrollment.header.extra_monthly4_param", "Доп. ежемесячная 4: параметр", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_AMOUNT, "enrollment.header.extra_monthly4_amount", "Доп. ежемесячная 4: размер", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_START, "enrollment.header.extra_monthly4_start", "Доп. ежемесячная 4: дата начала", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_BASIS, "enrollment.header.extra_monthly4_basis", "Доп. ежемесячная 4: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_MONTHLY4_ENABLED, "enrollment.header.extra_monthly4_enabled", "Доп. ежемесячная 4: вкл", 16

    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME1_NAME, "enrollment.header.extra_one_time1_name", "Доп. разовая 1: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME1_AMOUNT, "enrollment.header.extra_one_time1_amount", "Доп. разовая 1: сумма", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME1_DATE, "enrollment.header.extra_one_time1_date", "Доп. разовая 1: дата", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME1_BASIS, "enrollment.header.extra_one_time1_basis", "Доп. разовая 1: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME1_ENABLED, "enrollment.header.extra_one_time1_enabled", "Доп. разовая 1: вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME2_NAME, "enrollment.header.extra_one_time2_name", "Доп. разовая 2: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME2_AMOUNT, "enrollment.header.extra_one_time2_amount", "Доп. разовая 2: сумма", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME2_DATE, "enrollment.header.extra_one_time2_date", "Доп. разовая 2: дата", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME2_BASIS, "enrollment.header.extra_one_time2_basis", "Доп. разовая 2: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME2_ENABLED, "enrollment.header.extra_one_time2_enabled", "Доп. разовая 2: вкл", 16
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME3_NAME, "enrollment.header.extra_one_time3_name", "Доп. разовая 3: наименование", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME3_AMOUNT, "enrollment.header.extra_one_time3_amount", "Доп. разовая 3: сумма", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME3_DATE, "enrollment.header.extra_one_time3_date", "Доп. разовая 3: дата", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME3_BASIS, "enrollment.header.extra_one_time3_basis", "Доп. разовая 3: основание", 18
    SetEnrollmentHeader ws, COL_ENROLLMENT_EXTRA_ONE_TIME3_ENABLED, "enrollment.header.extra_one_time3_enabled", "Доп. разовая 3: вкл", 16

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, COL_ENROLLMENT_EXTRA_ONE_TIME3_ENABLED))
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Rows(1).RowHeight = 42
    ApplyEnrollmentTextColumnFormats ws
End Sub

Public Sub EnsureEnrollmentFormSheetStructure(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Cells.UnMerge
    On Error GoTo 0

    If Not IsBackendHeaderMatch(ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_KEY).Value, "Ключ", "Key") _
        Or Not IsBackendHeaderMatch(ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_LABEL).Value, "Поле", "Label") _
        Or Not IsBackendHeaderMatch(ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_VALUE).Value, "Значение", "Value") _
        Or Not IsBackendHeaderMatch(ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_DERIVED).Value, "Производное", "Derived") Then
        ws.Cells.Clear
    End If

    ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_KEY).Value = "Ключ"
    ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_LABEL).Value = "Поле"
    ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_VALUE).Value = "Значение"
    ws.Cells(BACKEND_HEADER_ROW, BACKEND_COL_DERIVED).Value = "Производное"

    EnsureBackendField ws, "current_row", "Текущая строка"
    EnsureBackendField ws, "enrollment_id", "ID зачисления"
    EnsureBackendField ws, "order_draft_id", "ID проекта приказа"
    EnsureBackendField ws, "personnel_event_id", "Personnel Event ID"
    EnsureBackendField ws, "source_mode", "Источник данных"
    EnsureBackendField ws, "fio", "ФИО"
    EnsureBackendField ws, "personal_number", "Личный номер"
    EnsureBackendField ws, "table_number", "Табельный номер"
    EnsureBackendField ws, "rank", "Воинское звание"
    EnsureBackendField ws, "service_category", "Категория службы"
    EnsureBackendField ws, "contract_kind", "Признак контракта"
    EnsureBackendField ws, "contract_basis", "Основание контракта"
    EnsureBackendField ws, "vus", "ВУС"
    EnsureBackendField ws, "position", "Штатная должность"
    EnsureBackendField ws, "section", "Подразделение / раздел персонала"
    EnsureBackendField ws, "military_unit", "Воинская часть"
    EnsureBackendField ws, "tariff_rank", "Тарифный разряд"
    EnsureBackendField ws, "position_salary", "Оклад по должности"
    EnsureBackendField ws, "rank_salary", "Оклад по званию"
    EnsureBackendField ws, "order_date", "Дата приказа"
    EnsureBackendField ws, "order_number", "Номер приказа"
    EnsureBackendField ws, "order_issuer", "Кем издан приказ"
    EnsureBackendField ws, "arrival_source", "Источник прибытия"
    EnsureBackendField ws, "prescription_number", "Номер предписания"
    EnsureBackendField ws, "prescription_date", "Дата предписания"
    EnsureBackendField ws, "report_number", "Номер рапорта"
    EnsureBackendField ws, "report_date", "Дата рапорта"
    EnsureBackendField ws, "report_info", "Рапорт / регистрация"
    EnsureBackendField ws, "assignment_info", "Предписание / основание"
    EnsureBackendField ws, "accept_date", "Дата принятия дел и должности"
    EnsureBackendField ws, "enroll_date", "Дата зачисления"
    EnsureBackendField ws, "duty_start_date", "Дата вступления в исполнение"
    EnsureBackendField ws, "manual_start_date", "Ручная дата старта"
    EnsureBackendField ws, "standard_start_date", "Дата старта стандартных выплат"
    EnsureBackendField ws, "preferential_start_date", "Дата льготной выслуги"
    EnsureBackendField ws, "basis_section1", "Основание §1"
    EnsureBackendField ws, "basis_section2", "Основание §2"
    EnsureBackendField ws, "birth_date", "Дата рождения"
    EnsureBackendField ws, "birth_place", "Место рождения"
    EnsureBackendField ws, "citizenship", "Гражданство"
    EnsureBackendField ws, "inn", "ИНН"
    EnsureBackendField ws, "snils", "СНИЛС"
    EnsureBackendField ws, "passport_series", "Серия паспорта"
    EnsureBackendField ws, "passport_number", "Номер паспорта"
    EnsureBackendField ws, "passport_issuer", "Кем выдан"
    EnsureBackendField ws, "passport_issue_date", "Дата выдачи паспорта"
    EnsureBackendField ws, "passport_code", "Код подразделения"
    EnsureBackendField ws, "bank_account", "Лицевой счёт / банк"
    EnsureBackendField ws, "bank_name", "Банк"
    EnsureBackendField ws, "requisites_note", "Примечание по реквизитам"
    EnsureBackendField ws, "preferential_enabled", "Льготная выслуга вкл"
    EnsureBackendField ws, "preferential_coeff", "Коэффициент льготной выслуги"
    EnsureBackendField ws, "preferential_basis", "Основание льготной выслуги"
    EnsureBackendField ws, "premium_enabled", "Премия вкл"
    EnsureBackendField ws, "premium_percent", "Премия %"
    EnsureBackendField ws, "premium_start", "Начало премии"
    EnsureBackendField ws, "premium_end", "Окончание премии"
    EnsureBackendField ws, "premium_basis", "Основание премии"
    EnsureBackendField ws, "lift_enabled", "Подъёмное вкл"
    EnsureBackendField ws, "lift_amount", "Подъёмное"
    EnsureBackendField ws, "lift_date", "Дата подъёмного"
    EnsureBackendField ws, "lift_basis", "Основание подъёмного"
    EnsureBackendField ws, "per_diem_enabled", "Суточные вкл"
    EnsureBackendField ws, "per_diem_days", "Суточные дни"
    EnsureBackendField ws, "per_diem_amount", "Размер суточных"
    EnsureBackendField ws, "per_diem_date", "Дата суточных"
    EnsureBackendField ws, "per_diem_basis", "Основание суточных"
    EnsureBackendField ws, "edv_enabled", "ЕДВ 400000 вкл"
    EnsureBackendField ws, "edv_amount", "Сумма ЕДВ"
    EnsureBackendField ws, "edv_date", "Дата ЕДВ"
    EnsureBackendField ws, "edv_basis", "Основание ЕДВ"
    EnsureBackendField ws, "class_param", "Классность"
    EnsureBackendField ws, "class_enabled", "Классность вкл"
    EnsureBackendField ws, "class_percent", "Классность %"
    EnsureBackendField ws, "class_basis", "Основание классности"
    EnsureBackendField ws, "fizo_param", "ФИЗО"
    EnsureBackendField ws, "fizo_enabled", "ФИЗО вкл"
    EnsureBackendField ws, "fizo_percent", "ФИЗО %"
    EnsureBackendField ws, "fizo_basis", "Основание ФИЗО"
    EnsureBackendField ws, "secrecy_param", "Секретность"
    EnsureBackendField ws, "secrecy_enabled", "Секретность вкл"
    EnsureBackendField ws, "secrecy_percent", "Секретность %"
    EnsureBackendField ws, "secrecy_basis", "Основание секретности"
    EnsureBackendField ws, "achievement_param", "Особое достижение / медаль"
    EnsureBackendField ws, "achievement_enabled", "Особое достижение вкл"
    EnsureBackendField ws, "achievement_amount", "Особое достижение % / сумма"
    EnsureBackendField ws, "achievement_basis", "Основание особых достижений"
    EnsureBackendField ws, "achievement_award_date", "Award order date"
    EnsureBackendField ws, "achievement_document_reference", "Award order reference"
    EnsureBackendField ws, "std_duty_enabled", "Надбавка по должности вкл"
    EnsureBackendField ws, "std_duty_percent", "Надбавка по должности %"
    EnsureBackendField ws, "std_duty_date", "Дата надбавки по должности"
    EnsureBackendField ws, "std_duty_basis", "Основание надбавки по должности"
    EnsureBackendField ws, "std_special_enabled", "Особые условия вкл"
    EnsureBackendField ws, "std_special_percent", "Особые условия %"
    EnsureBackendField ws, "std_special_date", "Дата особых условий"
    EnsureBackendField ws, "std_special_basis", "Основание особых условий"
    EnsureBackendField ws, "std_tariff_enabled", "1-4 тариф вкл"
    EnsureBackendField ws, "std_tariff_percent", "1-4 тариф %"
    EnsureBackendField ws, "std_tariff_date", "Дата 1-4 тарифа"
    EnsureBackendField ws, "std_tariff_basis", "Основание 1-4 тарифа"
    EnsureBackendField ws, "std_contract430_enabled", "430 приказ вкл"
    EnsureBackendField ws, "std_contract430_percent", "430 приказ %"
    EnsureBackendField ws, "std_contract430_date", "Дата 430 приказа"
    EnsureBackendField ws, "std_contract430_basis", "Основание 430 приказа"
    EnsureExtraBackendFields ws
    EnsureBackendField ws, "preview_status", "Предпросмотр: статус"
    EnsureBackendField ws, "preview_word_ready", "Предпросмотр: готовность Word"
    EnsureBackendField ws, "preview_issues", "Предпросмотр: замечания"
    EnsureBackendField ws, "preview_standard", "Предпросмотр: стандартные выплаты"
    EnsureBackendField ws, "preview_personal", "Предпросмотр: именные выплаты"
    EnsureBackendField ws, "preview_section1", "Предпросмотр: текст §1"
    EnsureBackendField ws, "preview_section2", "Предпросмотр: текст §2"

    ws.Columns("A:D").Font.Name = "Times New Roman"
    ws.Columns("A:D").Font.Size = 10
    ws.Columns(BACKEND_COL_KEY).ColumnWidth = 24
    ws.Columns(BACKEND_COL_LABEL).ColumnWidth = 28
    ws.Columns(BACKEND_COL_VALUE).ColumnWidth = 34
    ws.Columns(BACKEND_COL_DERIVED).ColumnWidth = 42
    ws.Columns(BACKEND_COL_VALUE).NumberFormat = "@"
    ws.Columns(BACKEND_COL_DERIVED).NumberFormat = "@"
End Sub

Private Function IsBackendHeaderMatch(ByVal rawValue As Variant, ByVal localizedValue As String, ByVal legacyValue As String) As Boolean
    Dim normalizedValue As String

    normalizedValue = Trim$(CStr(rawValue))
    IsBackendHeaderMatch = (StrComp(normalizedValue, localizedValue, vbTextCompare) = 0 Or _
        StrComp(normalizedValue, legacyValue, vbTextCompare) = 0)
End Function


Public Sub EnsureEnrollmentReferenceData()
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ENROLLMENT_REFERENCE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ENROLLMENT_REFERENCE_SHEET
    End If
    If SafeText(ws.Cells(1, 1).Value) = "" Then
        ws.Cells(1, 1).Value = "ReferenceType"
        ws.Cells(1, 2).Value = "Code"
        ws.Cells(1, 3).Value = "DisplayName"
        ws.Cells(1, 4).Value = "Amount"
        ws.Cells(1, 5).Value = "Active"
        ws.Cells(1, 6).Value = "Notes"
        AddEnrollmentReference ws, "SERVICE_CATEGORY", "CONTRACT", "Контракт", "", "YES", ""
        AddEnrollmentReference ws, "SERVICE_CATEGORY", "MOBILIZED", "Мобилизация", "", "YES", ""
        AddEnrollmentReference ws, "SERVICE_CATEGORY", "CONSCRIPT", "Призыв", "", "YES", ""
        For i = 1 To 50
            AddEnrollmentReference ws, "TARIFF_RANK", CStr(i), CStr(i) & " тарифный разряд", "", "YES", "Укажите оклад по должности для этого разряда."
        Next i
        AddEnrollmentReference ws, "RANK", "РЯДОВОЙ", "рядовой", "", "YES", "Укажите оклад по званию."
        AddEnrollmentReference ws, "RANK", "ЕФРЕЙТОР", "ефрейтор", "", "YES", "Укажите оклад по званию."
        AddEnrollmentReference ws, "RANK", "СЕРЖАНТ", "сержант", "", "YES", "Укажите оклад по званию."
        AddEnrollmentReference ws, "CLASS", "THIRD", "3 класс", "", "YES", "Заполните подтвержденный процент."
        AddEnrollmentReference ws, "CLASS", "SECOND", "2 класс", "", "YES", "Заполните подтвержденный процент."
        AddEnrollmentReference ws, "CLASS", "FIRST", "1 класс", "", "YES", "Заполните подтвержденный процент."
        AddEnrollmentReference ws, "SECRECY", "LEVEL_3", "3 степень", "", "YES", "Заполните подтвержденный процент."
        AddEnrollmentReference ws, "SECRECY", "LEVEL_2", "2 степень", "", "YES", "Заполните подтвержденный процент."
        AddEnrollmentReference ws, "SECRECY", "LEVEL_1", "1 степень", "", "YES", "Заполните подтвержденный процент."
        ws.Rows(1).Font.Bold = True
        FormatEnrollmentReferenceSheet ws
    End If
    EnsureEnrollmentReference ws, "FIZO", "2 уровень", "15"
    EnsureEnrollmentReferenceWithCode ws, "ACHIEVEMENT", "COMBAT_DISTINCTION", MedalDisplayName("COMBAT_DISTINCTION"), "30"
    EnsureEnrollmentReferenceWithCode ws, "ACHIEVEMENT", "DEMINING", MedalDisplayName("DEMINING"), "20"
    EnsureEnrollmentReferenceWithCode ws, "ACHIEVEMENT", "MILITARY_VALOR_I", MedalDisplayName("MILITARY_VALOR_I"), "20"
    EnsureEnrollmentReferenceWithCode ws, "ACHIEVEMENT", "MILITARY_VALOR_II", MedalDisplayName("MILITARY_VALOR_II"), "10"
    FormatEnrollmentReferenceSheet ws
        If Not enrollmentReferencesSynced Then
        SyncEnrollmentReferencesFromStaff ws
        enrollmentReferencesSynced = True
    End If
End Sub

Private Sub SyncEnrollmentReferencesFromStaff(ByVal wsReferences As Worksheet)
    Dim wsStaff As Worksheet
    Dim rowNum As Long
    Dim lastRow As Long
    Dim refRow As Long
    Dim vusColumn As Long
    Dim knownValues As Object

    On Error Resume Next
    mdlHelper.EnsureStaffColumnsInitialized
    Set wsStaff = mdlHelper.GetStaffWorksheet()
    On Error GoTo 0
    If wsStaff Is Nothing Then Exit Sub

    Set knownValues = CreateObject("Scripting.Dictionary")
    knownValues.CompareMode = vbTextCompare
    For refRow = 2 To wsReferences.Cells(wsReferences.Rows.Count, 1).End(xlUp).Row
        If SafeText(wsReferences.Cells(refRow, 1).Value) <> "" And SafeText(wsReferences.Cells(refRow, 3).Value) <> "" Then
            knownValues(UCase$(SafeText(wsReferences.Cells(refRow, 1).Value)) & "|" & UCase$(SafeText(wsReferences.Cells(refRow, 3).Value))) = True
        End If
    Next refRow

    vusColumn = FindEnrollmentStaffColumn(wsStaff, mdlHelper.Ru(1042, 1059, 1057))
    lastRow = wsStaff.Cells(wsStaff.Rows.Count, mdlHelper.colLichniyNomer_Global).End(xlUp).Row
    For rowNum = 2 To lastRow
        AddStaffReferenceIfNew wsReferences, knownValues, "RANK", SafeText(wsStaff.Cells(rowNum, mdlHelper.colZvanie_Global).Value)
        AddStaffReferenceIfNew wsReferences, knownValues, "POSITION", SafeText(wsStaff.Cells(rowNum, mdlHelper.colDolzhnost_Global).Value)
        AddStaffReferenceIfNew wsReferences, knownValues, "SECTION", SafeText(wsStaff.Cells(rowNum, mdlHelper.colVoinskayaChast_Global).Value)
        AddStaffReferenceIfNew wsReferences, knownValues, "MILITARY_UNIT", SafeText(wsStaff.Cells(rowNum, mdlHelper.colVoinskayaChast_Global).Value)
        If vusColumn > 0 Then AddStaffReferenceIfNew wsReferences, knownValues, "VUS", SafeText(wsStaff.Cells(rowNum, vusColumn).Value)
    Next rowNum
    FormatEnrollmentReferenceSheet wsReferences
End Sub

Private Sub FormatEnrollmentReferenceSheet(ByVal ws As Worksheet)
    ws.Rows(1).Font.Bold = True
    ws.Columns(1).ColumnWidth = 18
    ws.Columns(2).ColumnWidth = 20
    ws.Columns(3).ColumnWidth = 48
    ws.Columns(4).ColumnWidth = 14
    ws.Columns(5).ColumnWidth = 10
    ws.Columns(6).ColumnWidth = 60
    ws.Columns(6).WrapText = True
End Sub
Private Function FindEnrollmentStaffColumn(ByVal wsStaff As Worksheet, ByVal headerText As String) As Long
    Dim columnNum As Long
    Dim lastColumn As Long

    lastColumn = wsStaff.Cells(1, wsStaff.Columns.Count).End(xlToLeft).Column
    For columnNum = 1 To lastColumn
        If StrComp(SafeText(wsStaff.Cells(1, columnNum).Value), headerText, vbTextCompare) = 0 Then
            FindEnrollmentStaffColumn = columnNum
            Exit Function
        End If
    Next columnNum
End Function

Private Sub AddStaffReferenceIfNew(ByVal wsReferences As Worksheet, ByVal knownValues As Object, ByVal referenceType As String, ByVal displayName As String)
    Dim key As String

    If displayName = "" Then Exit Sub
    key = UCase$(referenceType) & "|" & UCase$(displayName)
    If knownValues.Exists(key) Then Exit Sub
    AddEnrollmentReference wsReferences, referenceType, displayName, displayName, "", "YES", "Импортировано из листа Штат; оператор может изменить или отключить строку."
    knownValues(key) = True
End Sub

Private Sub EnsureEnrollmentReference(ByVal ws As Worksheet, ByVal referenceType As String, ByVal displayName As String, ByVal amountValue As String)
    Dim rowNum As Long

    If displayName = "" Then Exit Sub
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(SafeText(ws.Cells(rowNum, 1).Value)) = UCase$(referenceType) _
            And StrComp(SafeText(ws.Cells(rowNum, 3).Value), displayName, vbTextCompare) = 0 Then Exit Sub
    Next rowNum
    AddEnrollmentReference ws, referenceType, displayName, displayName, amountValue, "YES", "Импортировано из листа Штат; оператор может изменить или отключить строку."
End Sub
Private Function MedalDisplayName(ByVal medalCode As String) As String
    Select Case UCase$(medalCode)
        Case "COMBAT_DISTINCTION"
            MedalDisplayName = mdlHelper.Ru(1052, 1077, 1076, 1072, 1083, 1100, 58, 32, 1073, 1086, 1077, 1074, 1099, 1077, 32, 1086, 1090, 1083, 1080, 1095, 1080, 1103)
        Case "DEMINING"
            MedalDisplayName = mdlHelper.Ru(1052, 1077, 1076, 1072, 1083, 1100, 58, 32, 1088, 1072, 1079, 1084, 1080, 1085, 1080, 1088, 1086, 1074, 1072, 1085, 1080, 1077)
        Case "MILITARY_VALOR_I"
            MedalDisplayName = mdlHelper.Ru(1052, 1077, 1076, 1072, 1083, 1100, 58, 32, 1074, 1086, 1080, 1085, 1089, 1082, 1072, 1103, 32, 1076, 1086, 1073, 1083, 1077, 1089, 1090, 1100, 32, 73)
        Case "MILITARY_VALOR_II"
            MedalDisplayName = mdlHelper.Ru(1052, 1077, 1076, 1072, 1083, 1100, 58, 32, 1074, 1086, 1080, 1085, 1089, 1082, 1072, 1103, 32, 1076, 1086, 1073, 1083, 1077, 1089, 1090, 1100, 32, 73, 73)
    End Select
End Function

Private Sub EnsureEnrollmentReferenceWithCode(ByVal ws As Worksheet, ByVal referenceType As String, ByVal code As String, ByVal displayName As String, ByVal amountValue As String)
    Dim rowNum As Long

    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(SafeText(ws.Cells(rowNum, 1).Value)) = UCase$(referenceType) _
            And UCase$(SafeText(ws.Cells(rowNum, 2).Value)) = UCase$(code) Then Exit Sub
    Next rowNum
    AddEnrollmentReference ws, referenceType, code, displayName, amountValue, "YES", "Reference value; operator may amend it after confirming the supporting document."
End Sub

Private Sub AddEnrollmentReference(ByVal ws As Worksheet, ByVal referenceType As String, ByVal code As String, ByVal displayName As String, ByVal amountValue As String, ByVal activeValue As String, ByVal noteText As String)
    Dim rowNum As Long
    rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(rowNum, 1).Value = referenceType
    ws.Cells(rowNum, 2).Value = code
    ws.Cells(rowNum, 3).Value = displayName
    ws.Cells(rowNum, 4).Value = amountValue
    ws.Cells(rowNum, 5).Value = activeValue
    ws.Cells(rowNum, 6).Value = noteText
End Sub

Public Function GetEnrollmentReferenceValues(ByVal referenceType As String) As Collection
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim values As New Collection

    EnsureEnrollmentReferenceData
    Set ws = ThisWorkbook.Worksheets(ENROLLMENT_REFERENCE_SHEET)
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(SafeText(ws.Cells(rowNum, 1).Value)) = UCase$(referenceType) Then
            If UCase$(SafeText(ws.Cells(rowNum, 5).Value)) <> "NO" Then values.Add SafeText(ws.Cells(rowNum, 3).Value)
        End If
    Next rowNum
    Set GetEnrollmentReferenceValues = values
End Function

Public Function GetEnrollmentReferenceAmount(ByVal referenceType As String, ByVal displayName As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long

    EnsureEnrollmentReferenceData
    Set ws = ThisWorkbook.Worksheets(ENROLLMENT_REFERENCE_SHEET)
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(SafeText(ws.Cells(rowNum, 1).Value)) = UCase$(referenceType) _
            And StrComp(SafeText(ws.Cells(rowNum, 3).Value), SafeText(displayName), vbTextCompare) = 0 Then
            GetEnrollmentReferenceAmount = SafeText(ws.Cells(rowNum, 4).Value)
            Exit Function
        End If
    Next rowNum
End Function

Public Function GetEnrollmentReferenceCode(ByVal referenceType As String, ByVal displayName As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long

    EnsureEnrollmentReferenceData
    Set ws = ThisWorkbook.Worksheets(ENROLLMENT_REFERENCE_SHEET)
    For rowNum = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(SafeText(ws.Cells(rowNum, 1).Value)) = UCase$(referenceType) _
            And StrComp(SafeText(ws.Cells(rowNum, 3).Value), SafeText(displayName), vbTextCompare) = 0 Then
            GetEnrollmentReferenceCode = SafeText(ws.Cells(rowNum, 2).Value)
            Exit Function
        End If
    Next rowNum
End Function

Public Sub ApplyEnrollmentReferenceValues(ByVal record As Object)
    Dim amountValue As String

    amountValue = GetEnrollmentReferenceAmount("RANK", SafeText(record("rank")))
    If amountValue <> "" Then record("rank_salary") = amountValue
    amountValue = GetEnrollmentReferenceAmount("TARIFF_RANK", SafeText(record("tariff_rank")) & " тарифный разряд")
    If amountValue <> "" Then record("position_salary") = amountValue
    amountValue = GetEnrollmentReferenceAmount("CLASS", SafeText(record("class_param")))
    If amountValue <> "" Then record("class_percent") = amountValue
    amountValue = GetEnrollmentReferenceAmount("FIZO", SafeText(record("fizo_param")))
    If amountValue <> "" Then record("fizo_percent") = amountValue
    amountValue = GetEnrollmentReferenceAmount("SECRECY", SafeText(record("secrecy_param")))
    If amountValue <> "" Then record("secrecy_percent") = amountValue
    amountValue = GetEnrollmentReferenceAmount("ACHIEVEMENT", SafeText(record("achievement_param")))
    If amountValue <> "" Then record("achievement_amount") = amountValue
End Sub
Public Sub OpenEnrollmentForm()
    EnsureEnrollmentInfrastructure
    ' The normal ribbon command always starts a new card. Existing cards are opened explicitly.
    ClearEnrollmentForm
    ' Apply configured standard-payment defaults before the operator sees the blank card.
    RefreshEnrollmentForm
    frmEnrollmentWizard.Show
End Sub

Public Sub OpenSelectedEnrollmentRowInForm()
    Call LoadSelectedEnrollmentRowToBackend
    frmEnrollmentWizard.Show
End Sub

Public Function LoadSelectedEnrollmentRowToBackend() As Long
    Dim rowNum As Long

    EnsureEnrollmentInfrastructure
    rowNum = ResolveActiveEnrollmentRow()
    LoadEnrollmentRowToBackend rowNum
    LoadSelectedEnrollmentRowToBackend = rowNum
End Function

Public Sub ClearEnrollmentForm()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT_FORM)
    On Error Resume Next
    ws.Cells.UnMerge
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, BACKEND_COL_KEY).End(xlUp).Row
    If lastRow >= BACKEND_FIRST_DATA_ROW Then
        For rowNum = BACKEND_FIRST_DATA_ROW To lastRow
            ClearBackendCellSafe ws.Cells(rowNum, BACKEND_COL_VALUE)
            ClearBackendCellSafe ws.Cells(rowNum, BACKEND_COL_DERIVED)
        Next rowNum
    End If

    SetBackendValue "source_mode", "manual"
    SyncEnrollmentWizardIfOpen
End Sub

Public Sub RefreshEnrollmentForm()
    Dim record As Object
    Dim evaluation As Object

    Set record = GetBackendRecord()
    NormalizeEnrollmentRecord record
    Set evaluation = EvaluateEnrollmentRecord(record)
    SetBackendRecord record
    WritePreviewToBackend record, evaluation
    SetBackendValue "personnel_event_id", mdlPersonnelEvents.EnsureEnrollmentPersonnelEvent(record)
    SyncEnrollmentWizardIfOpen
End Sub

Public Function SaveEnrollmentFormToSheet(Optional ByVal createPayments As Boolean = False) As Long
    Dim ws As Worksheet
    Dim rowNum As Long
    Dim record As Object
    Dim evaluation As Object

    EnsureEnrollmentInfrastructure
    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)

    Set record = GetBackendRecord()
    NormalizeEnrollmentRecord record
    Set evaluation = EvaluateEnrollmentRecord(record)

    rowNum = ResolveTargetRow(ws, record)
    WriteRecordToSheet ws, rowNum, record, evaluation
    Set record = GetResolvedEnrollmentRecordByRow(rowNum)

    SaveEnrollmentFormToSheet = rowNum

    SetBackendRecord record

    SetBackendValue "current_row", CStr(rowNum)
    SetBackendValue "enrollment_id", CStr(record("enrollment_id"))
    SetBackendValue "order_draft_id", CStr(record("order_draft_id"))
    WritePreviewToBackend record, evaluation
    SetBackendValue "personnel_event_id", mdlPersonnelEvents.EnsureEnrollmentPersonnelEvent(record)

    If createPayments Then
        Call GeneratePaymentsFromEnrollmentRowDirect(rowNum)
    End If

    SyncEnrollmentWizardIfOpen
End Function

Public Function SaveEnrollmentFormAndGeneratePayments() As Long
    Dim targetRow As Long

    targetRow = SaveEnrollmentFormToSheet(False)
    SaveEnrollmentFormAndGeneratePayments = GeneratePaymentsFromEnrollmentRowDirect(targetRow)
End Function

Public Function SaveEnrollmentFormAndContinuePackage() As String
    Dim targetRow As Long

    targetRow = SaveEnrollmentFormToSheet(False)
    SaveEnrollmentFormAndContinuePackage = PrepareNextEnrollmentInPackage(targetRow)
End Function

Public Function PrepareNextEnrollmentInPackage(Optional ByVal sourceRow As Long = 0) As String
    Dim sourceRecord As Object
    Dim nextRecord As Object

    EnsureEnrollmentInfrastructure

    If sourceRow >= 2 Then
        Set sourceRecord = GetResolvedEnrollmentRecordByRow(sourceRow)
    Else
        Set sourceRecord = GetResolvedBackendRecord()
    End If

    If SafeText(sourceRecord("order_draft_id")) = "" Then
        Err.Raise vbObjectError + 1765, "PrepareNextEnrollmentInPackage", ET("enrollment.form.error.continue_row_required", "Сначала сохраните текущую карточку зачисления, чтобы продолжить пакет.")
    End If

    Set nextRecord = BuildNextPackageRecord(sourceRecord)
    SetBackendRecord nextRecord, True
    SyncEnrollmentWizardIfOpen

    PrepareNextEnrollmentInPackage = SafeText(nextRecord("order_draft_id"))
End Function

Public Sub CloseEnrollmentWizardIfOpen()
    Dim frm As Object

    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmEnrollmentWizard" Then
            Unload frm
            Exit For
        End If
    Next frm
End Sub

Public Sub ValidateEnrollmentSheet(Optional ByVal isSilent As Boolean = False)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim issuesCount As Long

    On Error GoTo ErrorHandler

    EnsureEnrollmentInfrastructure
    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    lastRow = GetEnrollmentLastRow(ws)

    For rowNum = 2 To lastRow
        If RowHasEnrollmentInput(ws, rowNum) Then
            EvaluateAndApplyRow ws, rowNum
            If Trim$(CStr(ws.Cells(rowNum, COL_ENROLLMENT_VALIDATION_ISSUES).Value)) <> "" Then
                issuesCount = issuesCount + 1
            End If
        End If
    Next rowNum

    If Not isSilent Then
        MsgBox tf("enrollment.message.validation_done", "Проверка листа зачисления завершена. Строк с замечаниями: {count}", "{count}", issuesCount), vbInformation, ET("enrollment.caption.main", "Зачисление")
    End If
    Exit Sub

ErrorHandler:
    If Not isSilent Then
        MsgBox tf("enrollment.message.validation_error", "Ошибка при проверке листа зачисления: {error}", "{error}", Err.Description), vbCritical, ET("enrollment.caption.main", "Зачисление")
    End If
End Sub

Public Sub RefreshEnrollmentSuggestions()
    ValidateEnrollmentSheet True
    MsgBox ET("enrollment.message.suggestions_done", "Предложения по зачислению обновлены."), vbInformation, ET("enrollment.caption.main", "Зачисление")
End Sub

Public Sub RefreshEnrollmentRowDirect(ByVal sheetName As String, ByVal rowNum As Long)
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(sheetName)
    If rowNum < 2 Then Exit Sub
    EvaluateAndApplyRow ws, rowNum
End Sub

Public Sub GeneratePaymentsFromEnrollmentSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNum As Long
    Dim createdCount As Long

    EnsureEnrollmentInfrastructure
    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    lastRow = GetEnrollmentLastRow(ws)

    For rowNum = 2 To lastRow
        If RowHasEnrollmentInput(ws, rowNum) Then
            createdCount = createdCount + GeneratePaymentsFromEnrollmentRowDirect(rowNum)
        End If
    Next rowNum

    MsgBox tf("enrollment.message.transfer_done", "Подготовка выплат завершена. Добавлено строк: {count}", "{count}", createdCount), vbInformation, ET("enrollment.caption.main", "Зачисление")
End Sub

Public Function GeneratePaymentsFromEnrollmentRowDirect(ByVal rowNum As Long) As Long
    Dim wsEnrollment As Worksheet
    Dim wsPayments As Worksheet
    Dim record As Object
    Dim evaluation As Object
    Dim packageId As String

    EnsureEnrollmentInfrastructure
    Set wsEnrollment = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    Set wsPayments = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_PAYMENTS_NO_PERIODS)

    Set record = GetEnrollmentRecordFromWorksheet(wsEnrollment, rowNum)
    NormalizeEnrollmentRecord record
    Set evaluation = EvaluateEnrollmentRecord(record)
    If CLng(evaluation("severity")) = STATUS_BLOCKED Then Exit Function

    packageId = SafeText(record("order_draft_id"))
    If packageId = "" Then packageId = SafeText(record("enrollment_id"))

    If IsEnabledRecord(record, "class_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("class")("label")), SafeText(record("class_param")), SafeText(record("class_percent")) & "%", SafeText(record("basis_section1"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "fizo_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("fizo")("label")), SafeText(record("fizo_param")), SafeText(record("fizo_percent")) & "%", SafeText(record("basis_section1"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "secrecy_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("secrecy")("label")), SafeText(record("secrecy_param")), SafeText(record("secrecy_percent")) & "%", SafeText(record("basis_section1"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "achievement_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("achievement")("label")), SafeText(record("achievement_param")), SafeText(record("achievement_amount")) & "%", SafeText(record("achievement_basis"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "std_duty_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("std_duty")("label")), "", SafeText(record("std_duty_percent")) & "%", SafeText(record("std_duty_basis"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "std_special_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("std_special")("label")), "", SafeText(record("std_special_percent")) & "%", SafeText(record("std_special_basis"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "std_tariff_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("std_tariff")("label")), "", SafeText(record("std_tariff_percent")) & "%", SafeText(record("std_tariff_basis"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    If IsEnabledRecord(record, "std_contract430_enabled") Then
        AppendPaymentRowFromEnrollment wsPayments, record, packageId, CStr(GetEnrollmentPaymentDefinition("std_contract430")("label")), "", SafeText(record("std_contract430_percent")) & "%", SafeText(record("std_contract430_basis"))
        GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + 1
    End If
    GeneratePaymentsFromEnrollmentRowDirect = GeneratePaymentsFromEnrollmentRowDirect + AppendExtraPaymentRows(wsPayments, record, packageId)
End Function

Public Function GetBackendRecord() As Object
    Dim ws As Worksheet
    Dim record As Object
    Dim lastRow As Long
    Dim rowNum As Long
    Dim fieldKey As String

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT_FORM)
    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare

    lastRow = ws.Cells(ws.Rows.Count, BACKEND_COL_KEY).End(xlUp).Row
    For rowNum = BACKEND_FIRST_DATA_ROW To lastRow
        fieldKey = Trim$(CStr(ws.Cells(rowNum, BACKEND_COL_KEY).Value))
        If fieldKey <> "" Then
            record(fieldKey) = ws.Cells(rowNum, BACKEND_COL_VALUE).Value
        End If
    Next rowNum

    Set GetBackendRecord = record
End Function

Public Function GetResolvedBackendRecord() As Object
    Dim record As Object
    Set record = GetBackendRecord()
    NormalizeEnrollmentRecord record
    Set GetResolvedBackendRecord = record
End Function

Public Function GetEnrollmentRecordFromWorksheet(ByVal ws As Worksheet, ByVal rowNum As Long) As Object
    Dim record As Object

    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare

    record("current_row") = rowNum
    record("enrollment_id") = ws.Cells(rowNum, COL_ENROLLMENT_ID).Value
    record("fio") = ws.Cells(rowNum, COL_ENROLLMENT_FIO).Value
    record("personal_number") = ws.Cells(rowNum, COL_ENROLLMENT_LICHNIY_NOMER).Value
    record("rank") = ws.Cells(rowNum, COL_ENROLLMENT_RANK).Value
    record("position") = ws.Cells(rowNum, COL_ENROLLMENT_POSITION).Value
    record("section") = ws.Cells(rowNum, COL_ENROLLMENT_SECTION).Value
    record("order_date") = ws.Cells(rowNum, COL_ENROLLMENT_ORDER_DATE).Value
    record("order_number") = ws.Cells(rowNum, COL_ENROLLMENT_ORDER_NUMBER).Value
    record("achievement_param") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_PARAM).Value
    record("accept_date") = ws.Cells(rowNum, COL_ENROLLMENT_ACCEPT_DATE).Value
    record("enroll_date") = ws.Cells(rowNum, COL_ENROLLMENT_ENROLL_DATE).Value
    record("manual_start_date") = ws.Cells(rowNum, COL_ENROLLMENT_MANUAL_START_DATE).Value
    record("report_info") = ws.Cells(rowNum, COL_ENROLLMENT_REPORT_INFO).Value
    record("assignment_info") = ws.Cells(rowNum, COL_ENROLLMENT_ASSIGNMENT_INFO).Value
    record("class_param") = ws.Cells(rowNum, COL_ENROLLMENT_CLASS_PARAM).Value
    record("fizo_param") = ws.Cells(rowNum, COL_ENROLLMENT_FIZO_PARAM).Value
    record("secrecy_param") = ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_PARAM).Value
    record("standard_types") = ws.Cells(rowNum, COL_ENROLLMENT_STANDARD_TYPES).Value
    record("payment_basis") = ws.Cells(rowNum, COL_ENROLLMENT_PAYMENT_BASIS).Value
    record("comment") = ws.Cells(rowNum, COL_ENROLLMENT_COMMENT).Value
    record("order_draft_id") = ws.Cells(rowNum, COL_ENROLLMENT_ORDER_DRAFT_ID).Value
    record("word_ready") = ws.Cells(rowNum, COL_ENROLLMENT_WORD_READY).Value
    record("validation_severity") = ws.Cells(rowNum, COL_ENROLLMENT_VALIDATION_SEVERITY).Value
    record("validation_issues") = ws.Cells(rowNum, COL_ENROLLMENT_VALIDATION_ISSUES).Value
    record("source_mode") = ws.Cells(rowNum, COL_ENROLLMENT_SOURCE_MODE).Value
    record("last_derived_at") = ws.Cells(rowNum, COL_ENROLLMENT_LAST_DERIVED_AT).Value
    record("table_number") = ws.Cells(rowNum, COL_ENROLLMENT_TABLE_NUMBER).Value
    record("service_category") = ws.Cells(rowNum, COL_ENROLLMENT_SERVICE_CATEGORY).Value
    record("contract_kind") = ws.Cells(rowNum, COL_ENROLLMENT_CONTRACT_KIND).Value
    record("contract_basis") = ws.Cells(rowNum, COL_ENROLLMENT_CONTRACT_BASIS).Value
    record("vus") = ws.Cells(rowNum, COL_ENROLLMENT_VUS).Value
    record("military_unit") = ws.Cells(rowNum, COL_ENROLLMENT_MILITARY_UNIT).Value
    record("tariff_rank") = ws.Cells(rowNum, COL_ENROLLMENT_TARIFF_RANK).Value
    record("position_salary") = ws.Cells(rowNum, COL_ENROLLMENT_POSITION_SALARY).Value
    record("rank_salary") = ws.Cells(rowNum, COL_ENROLLMENT_RANK_SALARY).Value
    record("order_issuer") = ws.Cells(rowNum, COL_ENROLLMENT_ORDER_ISSUER).Value
    record("arrival_source") = ws.Cells(rowNum, COL_ENROLLMENT_ARRIVAL_SOURCE).Value
    record("prescription_number") = ws.Cells(rowNum, COL_ENROLLMENT_PRESCRIPTION_NUMBER).Value
    record("prescription_date") = ws.Cells(rowNum, COL_ENROLLMENT_PRESCRIPTION_DATE).Value
    record("report_number") = ws.Cells(rowNum, COL_ENROLLMENT_REPORT_NUMBER).Value
    record("report_date") = ws.Cells(rowNum, COL_ENROLLMENT_REPORT_DATE).Value
    record("duty_start_date") = ws.Cells(rowNum, COL_ENROLLMENT_DUTY_START_DATE).Value
    record("standard_start_date") = ws.Cells(rowNum, COL_ENROLLMENT_STANDARD_START_DATE).Value
    record("preferential_start_date") = ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_START_DATE).Value
    record("basis_section1") = ws.Cells(rowNum, COL_ENROLLMENT_BASIS_SECTION1).Value
    record("basis_section2") = ws.Cells(rowNum, COL_ENROLLMENT_BASIS_SECTION2).Value
    record("birth_date") = ws.Cells(rowNum, COL_ENROLLMENT_BIRTH_DATE).Value
    record("birth_place") = ws.Cells(rowNum, COL_ENROLLMENT_BIRTH_PLACE).Value
    record("citizenship") = ws.Cells(rowNum, COL_ENROLLMENT_CITIZENSHIP).Value
    record("inn") = ws.Cells(rowNum, COL_ENROLLMENT_INN).Value
    record("snils") = ws.Cells(rowNum, COL_ENROLLMENT_SNILS).Value
    record("passport_series") = ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_SERIES).Value
    record("passport_number") = ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_NUMBER).Value
    record("passport_issuer") = ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_ISSUER).Value
    record("passport_issue_date") = ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_ISSUE_DATE).Value
    record("passport_code") = ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_CODE).Value
    record("bank_account") = ws.Cells(rowNum, COL_ENROLLMENT_BANK_ACCOUNT).Value
    record("bank_name") = ws.Cells(rowNum, COL_ENROLLMENT_BANK_NAME).Value
    record("requisites_note") = ws.Cells(rowNum, COL_ENROLLMENT_REQUISITES_NOTE).Value
    record("preferential_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_ENABLED).Value
    record("preferential_coeff") = ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_COEFF).Value
    record("preferential_basis") = ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_BASIS).Value
    record("premium_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_ENABLED).Value
    record("premium_percent") = ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_PERCENT).Value
    record("premium_start") = ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_START).Value
    record("premium_end") = ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_END).Value
    record("premium_basis") = ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_BASIS).Value
    record("lift_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_LIFT_ENABLED).Value
    record("lift_amount") = ws.Cells(rowNum, COL_ENROLLMENT_LIFT_AMOUNT).Value
    record("lift_date") = ws.Cells(rowNum, COL_ENROLLMENT_LIFT_DATE).Value
    record("lift_basis") = ws.Cells(rowNum, COL_ENROLLMENT_LIFT_BASIS).Value
    record("per_diem_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_ENABLED).Value
    record("per_diem_days") = ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_DAYS).Value
    record("per_diem_amount") = ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_AMOUNT).Value
    record("per_diem_date") = ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_DATE).Value
    record("per_diem_basis") = ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_BASIS).Value
    record("edv_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_EDV_ENABLED).Value
    record("edv_amount") = ws.Cells(rowNum, COL_ENROLLMENT_EDV_AMOUNT).Value
    record("edv_date") = ws.Cells(rowNum, COL_ENROLLMENT_EDV_DATE).Value
    record("edv_basis") = ws.Cells(rowNum, COL_ENROLLMENT_EDV_BASIS).Value
    record("class_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_CLASS_ENABLED).Value
    record("class_percent") = ws.Cells(rowNum, COL_ENROLLMENT_CLASS_PERCENT).Value
    record("class_basis") = ws.Cells(rowNum, COL_ENROLLMENT_CLASS_BASIS).Value
    record("fizo_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_FIZO_ENABLED).Value
    record("fizo_percent") = ws.Cells(rowNum, COL_ENROLLMENT_FIZO_PERCENT).Value
    record("fizo_basis") = ws.Cells(rowNum, COL_ENROLLMENT_FIZO_BASIS).Value
    record("secrecy_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_ENABLED).Value
    record("secrecy_percent") = ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_PERCENT).Value
    record("secrecy_basis") = ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_BASIS).Value
    record("achievement_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_ENABLED).Value
    record("achievement_amount") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_AMOUNT).Value
    record("achievement_basis") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_BASIS).Value
    record("achievement_award_date") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_AWARD_DATE).Value
    record("achievement_document_reference") = ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_DOCUMENT_REFERENCE).Value
    record("std_duty_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_ENABLED).Value
    record("std_duty_percent") = ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_PERCENT).Value
    record("std_duty_date") = ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_DATE).Value
    record("std_duty_basis") = ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_BASIS).Value
    record("std_special_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_ENABLED).Value
    record("std_special_percent") = ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_PERCENT).Value
    record("std_special_date") = ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_DATE).Value
    record("std_special_basis") = ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_BASIS).Value
    record("std_tariff_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_ENABLED).Value
    record("std_tariff_percent") = ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_PERCENT).Value
    record("std_tariff_date") = ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_DATE).Value
    record("std_tariff_basis") = ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_BASIS).Value
    record("std_contract430_enabled") = ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_ENABLED).Value
    record("std_contract430_percent") = ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_PERCENT).Value
    record("std_contract430_date") = ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_DATE).Value
    record("std_contract430_basis") = ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_BASIS).Value
    LoadExtraPaymentFieldsFromWorksheet ws, rowNum, record

    Set GetEnrollmentRecordFromWorksheet = record
End Function

Public Function GetResolvedEnrollmentRecordByRow(ByVal rowNum As Long) As Object
    Dim record As Object

    Set record = GetEnrollmentRecordFromWorksheet(ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT), rowNum)
    NormalizeEnrollmentRecord record
    Set GetResolvedEnrollmentRecordByRow = record
End Function

Public Sub LoadEnrollmentRowToBackend(ByVal rowNum As Long)
    Dim ws As Worksheet
    Dim record As Object

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT)
    If rowNum < 2 Then Exit Sub
    If Not RowHasEnrollmentInput(ws, rowNum) Then Exit Sub

    Set record = GetEnrollmentRecordFromWorksheet(ws, rowNum)
    SetBackendRecord record, True
    SyncEnrollmentWizardIfOpen
End Sub

Public Function ResolveActiveEnrollmentRow() As Long
    If ActiveSheet Is Nothing Then
        Err.Raise vbObjectError + 1760, "ResolveActiveEnrollmentRow", ET("enrollment.message.active_row_required", "Выберите заполненную строку на листе 'Зачисление'.")
    End If

    If ActiveSheet.Name <> mdlReferenceData.SHEET_ENROLLMENT Then
        Err.Raise vbObjectError + 1761, "ResolveActiveEnrollmentRow", ET("enrollment.message.active_row_required", "Выберите заполненную строку на листе 'Зачисление'.")
    End If

    If ActiveCell Is Nothing Then
        Err.Raise vbObjectError + 1762, "ResolveActiveEnrollmentRow", ET("enrollment.message.active_row_required", "Выберите заполненную строку на листе 'Зачисление'.")
    End If

    If ActiveCell.Row < 2 Then
        Err.Raise vbObjectError + 1763, "ResolveActiveEnrollmentRow", ET("enrollment.message.active_row_required", "Выберите заполненную строку на листе 'Зачисление'.")
    End If

    If Not RowHasEnrollmentInput(ActiveSheet, ActiveCell.Row) Then
        Err.Raise vbObjectError + 1764, "ResolveActiveEnrollmentRow", ET("enrollment.message.active_row_required", "Выберите заполненную строку на листе 'Зачисление'.")
    End If

    ResolveActiveEnrollmentRow = ActiveCell.Row
End Function

Public Sub SetBackendValue(ByVal fieldKey As String, ByVal fieldValue As Variant)
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT_FORM)
    rowNum = FindBackendFieldRow(ws, fieldKey)
    If rowNum = 0 Then Exit Sub
    ws.Cells(rowNum, BACKEND_COL_VALUE).NumberFormat = "@"
    ws.Cells(rowNum, BACKEND_COL_VALUE).Value2 = CStr(SafeText(fieldValue))
End Sub

Public Function GetBackendValue(ByVal fieldKey As String) As Variant
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT_FORM)
    rowNum = FindBackendFieldRow(ws, fieldKey)
    If rowNum = 0 Then Exit Function
    GetBackendValue = ws.Cells(rowNum, BACKEND_COL_VALUE).Value
End Function

Public Function GetBackendDerivedValue(ByVal fieldKey As String) As Variant
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = ThisWorkbook.Worksheets(mdlReferenceData.SHEET_ENROLLMENT_FORM)
    rowNum = FindBackendFieldRow(ws, fieldKey)
    If rowNum = 0 Then Exit Function
    GetBackendDerivedValue = ws.Cells(rowNum, BACKEND_COL_DERIVED).Value
End Function

Public Sub NormalizeEnrollmentRecord(ByVal record As Object)
    Dim staffData As Object
    Dim contractDate As Date
    Dim enrollDate As Date
    Dim standardStart As Date
    Dim premiumWasAutoEnabled As Boolean
    Dim sourceModeText As String
    Dim explicitManualMode As Boolean

    EnsureRecordKeys record

    sourceModeText = LCase$(SafeText(record("source_mode")))
    explicitManualMode = (sourceModeText = "manual")
    If sourceModeText = "" Then
        record("source_mode") = "manual"
    End If

    If Not explicitManualMode Then
        Set staffData = FindStaffDataForRecord(record)
        If Not staffData Is Nothing Then
            If staffData.Count > 0 Then
                record("source_mode") = "staff"
                If SafeText(record("fio")) = "" And staffData.Exists("Лицо") Then record("fio") = staffData("Лицо")
                If SafeText(record("personal_number")) = "" And staffData.Exists("Личный номер") Then record("personal_number") = staffData("Личный номер")
                If SafeText(record("table_number")) = "" And staffData.Exists("Табельный номер") Then record("table_number") = staffData("Табельный номер")
                If SafeText(record("rank")) = "" And staffData.Exists("Воинское звание") Then record("rank") = staffData("Воинское звание")
                If SafeText(record("service_category")) = "" And staffData.Exists("Группа сотрудников") Then record("service_category") = staffData("Группа сотрудников")
                If SafeText(record("contract_kind")) = "" And staffData.Exists("Вид контракта") Then record("contract_kind") = staffData("Вид контракта")
                If SafeText(record("contract_basis")) = "" And staffData.Exists("Тип контракта") Then record("contract_basis") = staffData("Тип контракта")
                If SafeText(record("vus")) = "" And staffData.Exists("ВУС") Then record("vus") = staffData("ВУС")
                If SafeText(record("position")) = "" And staffData.Exists("Штатная должность") Then record("position") = staffData("Штатная должность")
                If SafeText(record("section")) = "" And staffData.Exists("Часть") Then record("section") = staffData("Часть")
                If SafeText(record("military_unit")) = "" And staffData.Exists("Часть") Then record("military_unit") = staffData("Часть")
                If SafeText(record("tariff_rank")) = "" And staffData.Exists("Тарифный разряд") Then record("tariff_rank") = staffData("Тарифный разряд")
                If SafeText(record("bank_account")) = "" And staffData.Exists("Номер счета в банке") Then record("bank_account") = staffData("Номер счета в банке")
                If SafeText(record("citizenship")) = "" And staffData.Exists("Гражданство") Then record("citizenship") = staffData("Гражданство")
                If SafeText(record("birth_date")) = "" And staffData.Exists("Дата рождения") Then
                    record("birth_date") = FormatStaffDateValue(staffData("Дата рождения"))
                End If
            End If
        End If
    End If

    ApplyEnrollmentReferenceValues record
    If SafeText(record("military_unit")) = "" Then record("military_unit") = SafeText(record("section"))
    If SafeText(record("order_draft_id")) = "" Then record("order_draft_id") = BuildEnrollmentOrderDraftId()
    If SafeText(record("enrollment_id")) = "" Then record("enrollment_id") = BuildEnrollmentId()

    enrollDate = mdlHelper.ParseDateSafe(record("enroll_date"))
    If SafeText(record("duty_start_date")) = "" And enrollDate > 0 Then record("duty_start_date") = Format$(enrollDate, "dd.mm.yyyy")
    If SafeText(record("standard_start_date")) = "" And enrollDate > 0 Then record("standard_start_date") = Format$(enrollDate, "dd.mm.yyyy")
    If SafeText(record("preferential_start_date")) = "" And enrollDate > 0 Then record("preferential_start_date") = Format$(enrollDate, "dd.mm.yyyy")

    If SafeText(record("preferential_enabled")) = "" Then record("preferential_enabled") = YES_VALUE
    If SafeText(record("preferential_coeff")) = "" Then record("preferential_coeff") = DEFAULT_PREFERENTIAL_COEFF
    If SafeText(record("preferential_basis")) = "" Then record("preferential_basis") = ET("enrollment.default.preferential_basis", "Льготная выслуга 1,5 по зачислению")

    premiumWasAutoEnabled = (SafeText(record("premium_enabled")) = "")
    If premiumWasAutoEnabled Then record("premium_enabled") = YES_VALUE
    If SafeText(record("premium_percent")) = "" Then record("premium_percent") = DEFAULT_PREMIUM_PERCENT
    If premiumWasAutoEnabled And SafeText(record("premium_end")) = "" Then record("premium_end") = "31.12." & CStr(IIf(enrollDate > 0, Year(enrollDate), Year(Date)))
    If SafeText(record("premium_basis")) = "" Then record("premium_basis") = ET("enrollment.default.premium_basis", "Премия за добросовестное и эффективное исполнение обязанностей")

    If SafeText(record("lift_enabled")) = "" Then record("lift_enabled") = YES_VALUE
    If SafeText(record("lift_amount")) = "" Then record("lift_amount") = DefaultLiftAmount()
    If SafeText(record("lift_basis")) = "" Then record("lift_basis") = ET("enrollment.default.lift_basis", "В связи с переездом к новому месту службы")

    If SafeText(record("per_diem_enabled")) = "" Then record("per_diem_enabled") = YES_VALUE
    If SafeText(record("per_diem_days")) = "" Then record("per_diem_days") = DEFAULT_PER_DIEM_DAYS
    If SafeText(record("per_diem_amount")) = "" Then record("per_diem_amount") = DefaultPerDiemAmount()
    If SafeText(record("per_diem_basis")) = "" Then record("per_diem_basis") = ET("enrollment.default.per_diem_basis", "В связи с переездом к новому месту службы")

    If SafeText(record("edv_amount")) = "" Then record("edv_amount") = DEFAULT_EDV_AMOUNT

    standardStart = mdlHelper.ParseDateSafe(record("standard_start_date"))
    If standardStart = 0 Then standardStart = enrollDate

    If SafeText(record("std_duty_enabled")) = "" Then record("std_duty_enabled") = YES_VALUE
    If SafeText(record("std_duty_percent")) = "" Then record("std_duty_percent") = DEFAULT_POSITION_ALLOWANCE_PERCENT
    If SafeText(record("std_duty_basis")) = "" Then record("std_duty_basis") = ET("enrollment.default.std_duty_basis", "Ежемесячная надбавка к денежному довольствию по занимаемой должности")

    If SafeText(record("std_special_enabled")) = "" Then record("std_special_enabled") = YES_VALUE
    If SafeText(record("std_special_percent")) = "" Then record("std_special_percent") = DEFAULT_SPECIAL_CONDITIONS_PERCENT
    If SafeText(record("std_special_basis")) = "" Then record("std_special_basis") = ET("enrollment.default.std_special_basis", "Надбавка за особые условия военной службы")

    If SafeText(record("std_tariff_enabled")) = "" Then
        record("std_tariff_enabled") = IIf(IsTariffOneToFour(record("tariff_rank")), YES_VALUE, NO_VALUE)
    End If
    If SafeText(record("std_tariff_percent")) = "" Then record("std_tariff_percent") = DEFAULT_TARIFF_PERCENT
    If SafeText(record("std_tariff_basis")) = "" Then record("std_tariff_basis") = ET("enrollment.default.std_tariff_basis", "Надбавка по должностям 1-4 тарифных разрядов")

    contractDate = ExtractContractDate(record)
    If SafeText(record("std_contract430_enabled")) = "" Then
        record("std_contract430_enabled") = IIf(IsContract430Candidate(record, contractDate), YES_VALUE, NO_VALUE)
    End If
    If SafeText(record("std_contract430_percent")) = "" Then record("std_contract430_percent") = DEFAULT_CONTRACT430_PERCENT
    If SafeText(record("std_contract430_basis")) = "" Then record("std_contract430_basis") = ET("enrollment.default.std_contract430_basis", "Надбавка за особые достижения в службе по 430 ДСП")

    If SafeText(record("class_param")) <> "" Then
        record("class_enabled") = YES_VALUE
        If SafeText(record("class_percent")) = "" Then record("class_percent") = CStr(ResolveClassPercent(SafeText(record("class_param"))))
        If SafeText(record("class_basis")) = "" Then record("class_basis") = ET("enrollment.default.class_basis", "Надбавка за классную квалификацию")
    ElseIf SafeText(record("class_enabled")) = "" Then
        record("class_enabled") = NO_VALUE
    End If

    If SafeText(record("fizo_param")) <> "" Then
        record("fizo_enabled") = YES_VALUE
        If SafeText(record("fizo_percent")) = "" Then record("fizo_percent") = CStr(ResolveFizoPercent(SafeText(record("fizo_param"))))
        If SafeText(record("fizo_basis")) = "" Then record("fizo_basis") = ET("enrollment.default.fizo_basis", "Надбавка за уровень физической подготовленности")
    ElseIf SafeText(record("fizo_enabled")) = "" Then
        record("fizo_enabled") = NO_VALUE
    End If

    If SafeText(record("secrecy_param")) <> "" Then
        record("secrecy_enabled") = YES_VALUE
        If SafeText(record("secrecy_percent")) = "" Then record("secrecy_percent") = CStr(ResolveSecrecyPercent(SafeText(record("secrecy_param"))))
        If SafeText(record("secrecy_basis")) = "" Then record("secrecy_basis") = ET("enrollment.default.secrecy_basis", "Надбавка за работу со сведениями, составляющими государственную тайну")
    ElseIf SafeText(record("secrecy_enabled")) = "" Then
        record("secrecy_enabled") = NO_VALUE
    End If

    If SafeText(record("achievement_param")) <> "" Then
        record("achievement_enabled") = YES_VALUE
        If SafeText(record("achievement_amount")) = "" Then record("achievement_amount") = ResolveAchievementAmountText(SafeText(record("achievement_param")))
        If SafeText(record("achievement_basis")) = "" Then record("achievement_basis") = ET("enrollment.default.achievement_basis", "Надбавка за особые достижения в службе / медаль")
    ElseIf SafeText(record("achievement_enabled")) = "" Then
        record("achievement_enabled") = NO_VALUE
    End If

    ApplyConfiguredPaymentStartDates record, enrollDate, standardStart
    NormalizeExtraPaymentFields record, standardStart, enrollDate

    If SafeText(record("basis_section1")) = "" Then record("basis_section1") = BuildSection1Basis(record)
    If SafeText(record("basis_section2")) = "" Then record("basis_section2") = BuildSection2Basis(record)

    record("standard_types") = JoinEnabledStandardTypes(record)
    record("payment_basis") = SafeText(record("basis_section1"))
    record("last_derived_at") = Format$(Now, "dd.mm.yyyy hh:nn:ss")
End Sub

Private Function FormatStaffDateValue(ByVal rawValue As Variant) As String
    On Error GoTo FallbackValue

    If IsEmpty(rawValue) Then Exit Function

    If IsNumeric(rawValue) Then
        FormatStaffDateValue = Format$(CDate(CDbl(rawValue)), "dd.mm.yyyy")
        Exit Function
    End If

    If IsDate(CStr(rawValue)) Then
        FormatStaffDateValue = Format$(CDate(CStr(rawValue)), "dd.mm.yyyy")
        Exit Function
    End If

FallbackValue:
    FormatStaffDateValue = SafeText(rawValue)
End Function

Public Function EvaluateEnrollmentRecord(ByVal record As Object) As Object
    Dim result As Object
    Dim issues As String
    Dim severity As Long
    Dim preview As Object

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    issues = ""
    severity = STATUS_READY

    AppendIssueIfBlank issues, severity, record("fio"), ET("enrollment.issue.fio_missing", "Не заполнено ФИО."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("personal_number"), ET("enrollment.issue.personal_number_missing", "Не заполнен личный номер."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("rank"), ET("enrollment.issue.rank_missing", "Не заполнено воинское звание."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("position"), ET("enrollment.issue.position_missing", "Не заполнена штатная должность."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("position_salary"), ET("enrollment.issue.position_salary_missing", "Не заполнен оклад по воинской должности."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("section"), ET("enrollment.issue.section_missing", "Не заполнено подразделение / раздел персонала."), STATUS_WARNING
    AppendIssueIfBlank issues, severity, record("order_number"), ET("enrollment.issue.order_number_missing", "Не заполнен номер приказа."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("order_date"), ET("enrollment.issue.order_date_missing", "Не заполнена дата приказа."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("order_issuer"), ET("enrollment.issue.order_issuer_missing", "Не заполнено поле 'кем издан приказ'."), STATUS_WARNING
    AppendIssueIfBlank issues, severity, record("accept_date"), ET("enrollment.issue.accept_date_missing", "Не заполнена дата принятия дел и должности."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("enroll_date"), ET("enrollment.issue.enroll_date_missing", "Не заполнена дата зачисления."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("duty_start_date"), ET("enrollment.issue.duty_start_missing", "Не заполнена дата вступления в исполнение обязанностей."), STATUS_BLOCKED
    ValidateDateField record, "order_date", "Дата приказа", issues, severity
    ValidateDateField record, "prescription_date", "Дата предписания", issues, severity
    ValidateDateField record, "report_date", "Дата рапорта", issues, severity
    ValidateDateField record, "accept_date", "Дата принятия дел и должности", issues, severity
    ValidateDateField record, "enroll_date", "Дата зачисления", issues, severity
    ValidateDateField record, "duty_start_date", "Дата вступления в обязанности", issues, severity
    ValidateDateField record, "manual_start_date", "Ручная дата старта", issues, severity
    ValidateDateField record, "standard_start_date", "Старт стандартных выплат", issues, severity
    ValidateDateField record, "preferential_start_date", "Старт льготной выслуги", issues, severity
    ValidateDateField record, "premium_start", "Начало премии", issues, severity
    ValidateDateField record, "premium_end", "Окончание премии", issues, severity
    ValidateDateField record, "lift_date", "Дата подъёмного пособия", issues, severity
    ValidateDateField record, "per_diem_date", "Дата суточных", issues, severity
    ValidateDateField record, "edv_date", "Дата ЕДВ", issues, severity
    ValidateDateField record, "birth_date", "Дата рождения", issues, severity
    ValidateDateField record, "passport_issue_date", "Дата выдачи паспорта", issues, severity
    ValidateStaffConsistency record, issues, severity

    If IsAllBlank(record("prescription_number"), record("prescription_date"), record("assignment_info")) Then
        AppendIssue issues, ET("enrollment.issue.assignment_missing", "Не заполнены сведения о предписании / основании прибытия."), severity, STATUS_WARNING
    End If
    If IsAllBlank(record("report_number"), record("report_date"), record("report_info")) Then
        AppendIssue issues, ET("enrollment.issue.report_missing", "Не заполнены сведения о рапорте / регистрации."), severity, STATUS_WARNING
    End If

    ValidateConfiguredPaymentDefinitions record, issues, severity
    ValidatePersonalPaymentAmounts record, issues, severity
    ValidateExtraPaymentFields record, issues, severity

    AppendIssueIfBlank issues, severity, record("passport_series"), ET("enrollment.issue.passport_series_missing", "Не заполнена серия паспорта."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("passport_number"), ET("enrollment.issue.passport_number_missing", "Не заполнен номер паспорта."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("passport_issuer"), ET("enrollment.issue.passport_issuer_missing", "Не заполнено поле 'кем выдан паспорт'."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("passport_issue_date"), ET("enrollment.issue.passport_date_missing", "Не заполнена дата выдачи паспорта."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("inn"), ET("enrollment.issue.inn_missing", "Не заполнен ИНН."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("snils"), ET("enrollment.issue.snils_missing", "Не заполнен СНИЛС."), STATUS_BLOCKED
    ValidateExactDigits record, "inn", "ИНН", 10, 12, issues, severity
    ValidateExactDigits record, "snils", "СНИЛС", 11, 11, issues, severity, True
    ValidateExactDigits record, "passport_series", "Серия паспорта", 4, 4, issues, severity
    ValidateExactDigits record, "passport_number", "Номер паспорта", 6, 6, issues, severity
    ValidatePassportCode record, issues, severity
    AppendIssueIfBlank issues, severity, record("bank_account"), ET("enrollment.issue.bank_account_missing", "Не заполнен лицевой / банковский счёт."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("bank_name"), ET("enrollment.issue.bank_name_missing", "Не заполнено наименование банка."), STATUS_BLOCKED
    AppendIssueIfBlank issues, severity, record("basis_section1"), ET("enrollment.issue.basis_section1_missing", "Не заполнен общий блок оснований для §1."), STATUS_BLOCKED

    If SafeText(record("order_draft_id")) = "" Then
        AppendIssue issues, ET("enrollment.issue.order_draft_warning", "OrderDraftId пуст. Для пакетного экспорта нужно заполнить идентификатор проекта."), severity, STATUS_WARNING
    End If

    If SafeText(record("arrival_source")) = "" Then
        AppendIssue issues, ET("enrollment.issue.arrival_source_missing", "Не заполнен пункт отбора / источник прибытия."), severity, STATUS_WARNING
    End If

    Set preview = mdlEnrollmentOrderExport.BuildPreviewPayload(record)

    result("severity") = severity
    result("status_text") = SeverityToStatusText(severity)
    result("issues") = issues
    result("word_ready") = IIf(severity = STATUS_BLOCKED, NO_VALUE, YES_VALUE)
    result("preview_standard") = preview("standard_text")
    result("preview_personal") = preview("personal_text")
    result("preview_section1") = preview("section1_text")
    result("preview_section2") = preview("section2_text")

    Set EvaluateEnrollmentRecord = result
End Function

Private Sub EvaluateAndApplyRow(ByVal ws As Worksheet, ByVal rowNum As Long)
    Dim record As Object
    Dim evaluation As Object

    Set record = GetEnrollmentRecordFromWorksheet(ws, rowNum)
    NormalizeEnrollmentRecord record
    Set evaluation = EvaluateEnrollmentRecord(record)
    WriteRecordToSheet ws, rowNum, record, evaluation
End Sub

Private Sub WriteRecordToSheet(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal record As Object, ByVal evaluation As Object)
    If SafeText(record("enrollment_id")) = "" Then record("enrollment_id") = BuildEnrollmentId()
    If SafeText(record("order_draft_id")) = "" Then record("order_draft_id") = BuildEnrollmentOrderDraftId()

    ws.Cells(rowNum, COL_ENROLLMENT_ID).Value = record("enrollment_id")
    ws.Cells(rowNum, COL_ENROLLMENT_FIO).Value = record("fio")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_LICHNIY_NOMER, record("personal_number")
    ws.Cells(rowNum, COL_ENROLLMENT_RANK).Value = record("rank")
    ws.Cells(rowNum, COL_ENROLLMENT_POSITION).Value = record("position")
    ws.Cells(rowNum, COL_ENROLLMENT_SECTION).Value = record("section")
    ws.Cells(rowNum, COL_ENROLLMENT_ORDER_DATE).Value = record("order_date")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_ORDER_NUMBER, record("order_number")
    ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_PARAM).Value = record("achievement_param")
    ws.Cells(rowNum, COL_ENROLLMENT_ACCEPT_DATE).Value = record("accept_date")
    ws.Cells(rowNum, COL_ENROLLMENT_ENROLL_DATE).Value = record("enroll_date")
    ws.Cells(rowNum, COL_ENROLLMENT_MANUAL_START_DATE).Value = record("manual_start_date")
    ws.Cells(rowNum, COL_ENROLLMENT_REPORT_INFO).Value = record("report_info")
    ws.Cells(rowNum, COL_ENROLLMENT_ASSIGNMENT_INFO).Value = record("assignment_info")
    ws.Cells(rowNum, COL_ENROLLMENT_CLASS_PARAM).Value = record("class_param")
    ws.Cells(rowNum, COL_ENROLLMENT_FIZO_PARAM).Value = record("fizo_param")
    ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_PARAM).Value = record("secrecy_param")
    ws.Cells(rowNum, COL_ENROLLMENT_STANDARD_TYPES).Value = record("standard_types")
    ws.Cells(rowNum, COL_ENROLLMENT_PAYMENT_BASIS).Value = record("payment_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_STATUS).Value = evaluation("status_text")
    ws.Cells(rowNum, COL_ENROLLMENT_COMMENT).Value = evaluation("preview_section1")
    ws.Cells(rowNum, COL_ENROLLMENT_ORDER_DRAFT_ID).Value = record("order_draft_id")
    ws.Cells(rowNum, COL_ENROLLMENT_WORD_READY).Value = evaluation("word_ready")
    ws.Cells(rowNum, COL_ENROLLMENT_VALIDATION_SEVERITY).Value = evaluation("severity")
    ws.Cells(rowNum, COL_ENROLLMENT_VALIDATION_ISSUES).Value = evaluation("issues")
    ws.Cells(rowNum, COL_ENROLLMENT_SOURCE_MODE).Value = record("source_mode")
    ws.Cells(rowNum, COL_ENROLLMENT_LAST_DERIVED_AT).Value = record("last_derived_at")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_TABLE_NUMBER, record("table_number")
    ws.Cells(rowNum, COL_ENROLLMENT_SERVICE_CATEGORY).Value = record("service_category")
    ws.Cells(rowNum, COL_ENROLLMENT_CONTRACT_KIND).Value = record("contract_kind")
    ws.Cells(rowNum, COL_ENROLLMENT_CONTRACT_BASIS).Value = record("contract_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_VUS).Value = record("vus")
    ws.Cells(rowNum, COL_ENROLLMENT_MILITARY_UNIT).Value = record("military_unit")
    ws.Cells(rowNum, COL_ENROLLMENT_TARIFF_RANK).Value = record("tariff_rank")
    ws.Cells(rowNum, COL_ENROLLMENT_POSITION_SALARY).Value = record("position_salary")
    ws.Cells(rowNum, COL_ENROLLMENT_RANK_SALARY).Value = record("rank_salary")
    ws.Cells(rowNum, COL_ENROLLMENT_ORDER_ISSUER).Value = record("order_issuer")
    ws.Cells(rowNum, COL_ENROLLMENT_ARRIVAL_SOURCE).Value = record("arrival_source")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_PRESCRIPTION_NUMBER, record("prescription_number")
    ws.Cells(rowNum, COL_ENROLLMENT_PRESCRIPTION_DATE).Value = record("prescription_date")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_REPORT_NUMBER, record("report_number")
    ws.Cells(rowNum, COL_ENROLLMENT_REPORT_DATE).Value = record("report_date")
    ws.Cells(rowNum, COL_ENROLLMENT_DUTY_START_DATE).Value = record("duty_start_date")
    ws.Cells(rowNum, COL_ENROLLMENT_STANDARD_START_DATE).Value = record("standard_start_date")
    ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_START_DATE).Value = record("preferential_start_date")
    ws.Cells(rowNum, COL_ENROLLMENT_BASIS_SECTION1).Value = record("basis_section1")
    ws.Cells(rowNum, COL_ENROLLMENT_BASIS_SECTION2).Value = record("basis_section2")
    ws.Cells(rowNum, COL_ENROLLMENT_BIRTH_DATE).Value = record("birth_date")
    ws.Cells(rowNum, COL_ENROLLMENT_BIRTH_PLACE).Value = record("birth_place")
    ws.Cells(rowNum, COL_ENROLLMENT_CITIZENSHIP).Value = record("citizenship")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_INN, record("inn")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_SNILS, record("snils")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_PASSPORT_SERIES, record("passport_series")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_PASSPORT_NUMBER, record("passport_number")
    ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_ISSUER).Value = record("passport_issuer")
    ws.Cells(rowNum, COL_ENROLLMENT_PASSPORT_ISSUE_DATE).Value = record("passport_issue_date")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_PASSPORT_CODE, record("passport_code")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_BANK_ACCOUNT, record("bank_account")
    ws.Cells(rowNum, COL_ENROLLMENT_BANK_NAME).Value = record("bank_name")
    ws.Cells(rowNum, COL_ENROLLMENT_REQUISITES_NOTE).Value = record("requisites_note")
    ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_ENABLED).Value = record("preferential_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_COEFF).Value = record("preferential_coeff")
    ws.Cells(rowNum, COL_ENROLLMENT_PREFERENTIAL_BASIS).Value = record("preferential_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_ENABLED).Value = record("premium_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_PERCENT).Value = record("premium_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_START).Value = record("premium_start")
    ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_END).Value = record("premium_end")
    ws.Cells(rowNum, COL_ENROLLMENT_PREMIUM_BASIS).Value = record("premium_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_LIFT_ENABLED).Value = record("lift_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_LIFT_AMOUNT).Value = record("lift_amount")
    ws.Cells(rowNum, COL_ENROLLMENT_LIFT_DATE).Value = record("lift_date")
    ws.Cells(rowNum, COL_ENROLLMENT_LIFT_BASIS).Value = record("lift_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_ENABLED).Value = record("per_diem_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_DAYS).Value = record("per_diem_days")
    ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_AMOUNT).Value = record("per_diem_amount")
    ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_DATE).Value = record("per_diem_date")
    ws.Cells(rowNum, COL_ENROLLMENT_PER_DIEM_BASIS).Value = record("per_diem_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_EDV_ENABLED).Value = record("edv_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_EDV_AMOUNT).Value = record("edv_amount")
    ws.Cells(rowNum, COL_ENROLLMENT_EDV_DATE).Value = record("edv_date")
    ws.Cells(rowNum, COL_ENROLLMENT_EDV_BASIS).Value = record("edv_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_CLASS_ENABLED).Value = record("class_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_CLASS_PERCENT).Value = record("class_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_CLASS_BASIS).Value = record("class_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_FIZO_ENABLED).Value = record("fizo_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_FIZO_PERCENT).Value = record("fizo_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_FIZO_BASIS).Value = record("fizo_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_ENABLED).Value = record("secrecy_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_PERCENT).Value = record("secrecy_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_SECRECY_BASIS).Value = record("secrecy_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_ENABLED).Value = record("achievement_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_AMOUNT).Value = record("achievement_amount")
    ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_BASIS).Value = record("achievement_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_ACHIEVEMENT_AWARD_DATE).Value = record("achievement_award_date")
    WriteTextCell ws, rowNum, COL_ENROLLMENT_ACHIEVEMENT_DOCUMENT_REFERENCE, record("achievement_document_reference")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_ENABLED).Value = record("std_duty_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_PERCENT).Value = record("std_duty_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_DATE).Value = record("std_duty_date")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_DUTY_BASIS).Value = record("std_duty_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_ENABLED).Value = record("std_special_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_PERCENT).Value = record("std_special_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_DATE).Value = record("std_special_date")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_SPECIAL_BASIS).Value = record("std_special_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_ENABLED).Value = record("std_tariff_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_PERCENT).Value = record("std_tariff_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_DATE).Value = record("std_tariff_date")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_TARIFF_BASIS).Value = record("std_tariff_basis")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_ENABLED).Value = record("std_contract430_enabled")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_PERCENT).Value = record("std_contract430_percent")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_DATE).Value = record("std_contract430_date")
    ws.Cells(rowNum, COL_ENROLLMENT_STD_CONTRACT430_BASIS).Value = record("std_contract430_basis")
    WriteExtraPaymentFieldsToWorksheet ws, rowNum, record

    ApplyRowStatusFormatting ws, rowNum, CLng(evaluation("severity"))
End Sub

Private Sub ApplyRowStatusFormatting(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal severity As Long)
    Select Case severity
        Case STATUS_BLOCKED
            ws.Cells(rowNum, COL_ENROLLMENT_STATUS).Interior.Color = RGB(255, 199, 206)
        Case STATUS_WARNING
            ws.Cells(rowNum, COL_ENROLLMENT_STATUS).Interior.Color = RGB(255, 235, 156)
        Case Else
            ws.Cells(rowNum, COL_ENROLLMENT_STATUS).Interior.Color = RGB(198, 239, 206)
    End Select
End Sub

Private Sub WritePreviewToBackend(ByVal record As Object, ByVal evaluation As Object)
    SetBackendValue "preview_status", evaluation("status_text")
    SetBackendValue "preview_word_ready", evaluation("word_ready")
    SetBackendValue "preview_issues", evaluation("issues")
    SetBackendValue "preview_standard", evaluation("preview_standard")
    SetBackendValue "preview_personal", evaluation("preview_personal")
    SetBackendValue "preview_section1", evaluation("preview_section1")
    SetBackendValue "preview_section2", evaluation("preview_section2")
    SetBackendValue "basis_section1", record("basis_section1")
    SetBackendValue "basis_section2", record("basis_section2")
End Sub

Private Sub AppendPaymentRowFromEnrollment(ByVal wsPayments As Worksheet, ByVal record As Object, ByVal packageId As String, ByVal paymentType As String, ByVal parameterValue As String, ByVal amountText As String, ByVal basisText As String)
    Dim nextRow As Long

    nextRow = wsPayments.Cells(wsPayments.Rows.Count, mdlPaymentValidation.COL_LICHNIY_NOMER).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_NUMBER).Value = nextRow - 1
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_FIO).Value = record("fio")
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_LICHNIY_NOMER).Value = record("personal_number")
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_PAYMENT_TYPE).Value = paymentType
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_AMOUNT).Value = amountText
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_FOUNDATION).Value = basisText
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_PACKAGE_ID).Value = packageId
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_PACKAGE_MODE).Value = "LIST"
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_PARAMETER).Value = parameterValue
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_SHARED_FOUNDATION).Value = basisText
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_GROUP_EXPORT).Value = YES_VALUE
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_NOTE).Value = ET("enrollment.payment.note", "Сформировано из листа зачисления")
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_STATUS).Value = "READY"
    wsPayments.Cells(nextRow, mdlPaymentValidation.COL_SOURCE_ENROLLMENT_ID).Value = record("enrollment_id")
End Sub

Private Function ResolveTargetRow(ByVal ws As Worksheet, ByVal record As Object) As Long
    Dim currentRow As Long

    If IsNumeric(record("current_row")) Then
        currentRow = CLng(record("current_row"))
        If currentRow >= 2 Then
            ResolveTargetRow = currentRow
            Exit Function
        End If
    End If

    ResolveTargetRow = GetEnrollmentLastRow(ws) + 1
    If ResolveTargetRow < 2 Then ResolveTargetRow = 2
End Function

Private Function ShouldLoadActiveEnrollmentRow() As Boolean
    On Error GoTo SafeExit

    If ActiveWorkbook Is Nothing Then Exit Function
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Function
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
    If StrComp(ActiveSheet.Name, mdlReferenceData.SHEET_ENROLLMENT, vbTextCompare) <> 0 Then Exit Function
    If ActiveCell Is Nothing Then Exit Function
    If ActiveCell.Row < 2 Then Exit Function

    ShouldLoadActiveEnrollmentRow = RowHasEnrollmentInput(ActiveSheet, ActiveCell.Row)

SafeExit:
End Function

Private Sub SetBackendRecord(ByVal record As Object, Optional ByVal clearExisting As Boolean = False)
    Dim fieldKey As Variant

    If clearExisting Then
        ClearEnrollmentForm
    End If

    For Each fieldKey In record.Keys
        SetBackendValue CStr(fieldKey), record(fieldKey)
    Next fieldKey
End Sub

Private Function BuildNextPackageRecord(ByVal sourceRecord As Object) As Object
    Dim nextRecord As Object

    Set nextRecord = CreateObject("Scripting.Dictionary")
    nextRecord.CompareMode = vbTextCompare

    CopyRecordValues sourceRecord, nextRecord
    EnsureRecordKeys nextRecord

    ClearRecordFields nextRecord, Array( _
        "current_row", "enrollment_id", "fio", "personal_number", "table_number", "rank", _
        "service_category", "contract_kind", "contract_basis", "vus", "position", "section", _
        "military_unit", "tariff_rank", "position_salary", "rank_salary", "prescription_number", _
        "prescription_date", "report_number", "report_date", "report_info", "assignment_info", _
        "birth_date", "birth_place", "citizenship", "inn", "snils", "passport_series", _
        "passport_number", "passport_issuer", "passport_issue_date", "passport_code", _
        "bank_account", "bank_name", "requisites_note", "class_param", "class_enabled", _
        "class_percent", "class_basis", "fizo_param", "fizo_enabled", "fizo_percent", _
        "fizo_basis", "secrecy_param", "secrecy_enabled", "secrecy_percent", "secrecy_basis", _
        "achievement_param", "achievement_enabled", "achievement_amount", "achievement_basis", "achievement_award_date", "achievement_document_reference", _
        "basis_section1", "basis_section2", "standard_types", "payment_basis", "comment", _
        "word_ready", "validation_severity", "validation_issues", "last_derived_at")

    ClearExtraPaymentFields nextRecord

    nextRecord("source_mode") = "manual"

    Set BuildNextPackageRecord = nextRecord
End Function

Private Sub CopyRecordValues(ByVal sourceRecord As Object, ByVal targetRecord As Object)
    Dim fieldKey As Variant

    For Each fieldKey In sourceRecord.Keys
        targetRecord(CStr(fieldKey)) = sourceRecord(fieldKey)
    Next fieldKey
End Sub

Private Sub ClearRecordFields(ByVal record As Object, ByVal fieldNames As Variant)
    Dim i As Long

    For i = LBound(fieldNames) To UBound(fieldNames)
        record(CStr(fieldNames(i))) = vbNullString
    Next i
End Sub

Private Sub ClearExtraPaymentFields(ByVal record As Object)
    Dim i As Long

    For i = 1 To 4
        record(ExtraMonthlyKey(i, "name")) = vbNullString
        record(ExtraMonthlyKey(i, "param")) = vbNullString
        record(ExtraMonthlyKey(i, "amount")) = vbNullString
        record(ExtraMonthlyKey(i, "start")) = vbNullString
        record(ExtraMonthlyKey(i, "basis")) = vbNullString
        record(ExtraMonthlyKey(i, "enabled")) = vbNullString
    Next i

    For i = 1 To 3
        record(ExtraOneTimeKey(i, "name")) = vbNullString
        record(ExtraOneTimeKey(i, "amount")) = vbNullString
        record(ExtraOneTimeKey(i, "date")) = vbNullString
        record(ExtraOneTimeKey(i, "basis")) = vbNullString
        record(ExtraOneTimeKey(i, "enabled")) = vbNullString
    Next i
End Sub

Private Function GetEnrollmentLastRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, COL_ENROLLMENT_FIO).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, COL_ENROLLMENT_LICHNIY_NOMER).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, COL_ENROLLMENT_LICHNIY_NOMER).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, COL_ENROLLMENT_ORDER_NUMBER).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, COL_ENROLLMENT_ORDER_NUMBER).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, COL_ENROLLMENT_ORDER_DRAFT_ID).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, COL_ENROLLMENT_ORDER_DRAFT_ID).End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    GetEnrollmentLastRow = lastRow
End Function

Private Function RowHasEnrollmentInput(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    RowHasEnrollmentInput = _
        SafeText(ws.Cells(rowNum, COL_ENROLLMENT_FIO).Value) <> "" Or _
        SafeText(ws.Cells(rowNum, COL_ENROLLMENT_LICHNIY_NOMER).Value) <> "" Or _
        SafeText(ws.Cells(rowNum, COL_ENROLLMENT_ORDER_NUMBER).Value) <> "" Or _
        SafeText(ws.Cells(rowNum, COL_ENROLLMENT_ORDER_DRAFT_ID).Value) <> ""
End Function

Private Function EnsureWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Sub EnsureBackendField(ByVal ws As Worksheet, ByVal fieldKey As String, ByVal fieldLabel As String)
    Dim rowNum As Long
    rowNum = FindBackendFieldRow(ws, fieldKey)
    If rowNum = 0 Then
        rowNum = ws.Cells(ws.Rows.Count, BACKEND_COL_KEY).End(xlUp).Row + 1
        If rowNum < BACKEND_FIRST_DATA_ROW Then rowNum = BACKEND_FIRST_DATA_ROW
        ws.Cells(rowNum, BACKEND_COL_KEY).Value = fieldKey
    End If
    ws.Cells(rowNum, BACKEND_COL_LABEL).Value = fieldLabel
End Sub

Private Function FindBackendFieldRow(ByVal ws As Worksheet, ByVal fieldKey As String) As Long
    Dim lastRow As Long
    Dim rowNum As Long

    lastRow = ws.Cells(ws.Rows.Count, BACKEND_COL_KEY).End(xlUp).Row
    For rowNum = BACKEND_FIRST_DATA_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowNum, BACKEND_COL_KEY).Value)), fieldKey, vbTextCompare) = 0 Then
            FindBackendFieldRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Sub SetEnrollmentHeader(ByVal ws As Worksheet, ByVal columnNumber As Long, ByVal key As String, ByVal fallback As String, ByVal columnWidth As Double)
    ws.Cells(1, columnNumber).Value = ET(key, fallback)
    ws.Columns(columnNumber).ColumnWidth = columnWidth
End Sub

Private Sub WriteTextCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal columnNumber As Long, ByVal valueText As Variant)
    ws.Cells(rowNum, columnNumber).NumberFormat = "@"
    ws.Cells(rowNum, columnNumber).Value2 = CStr(SafeText(valueText))
End Sub

Private Sub ApplyEnrollmentTextColumnFormats(ByVal ws As Worksheet)
    Dim textColumns As Variant
    Dim i As Long

    textColumns = Array( _
        COL_ENROLLMENT_LICHNIY_NOMER, COL_ENROLLMENT_TABLE_NUMBER, _
        COL_ENROLLMENT_ORDER_NUMBER, COL_ENROLLMENT_PRESCRIPTION_NUMBER, COL_ENROLLMENT_REPORT_NUMBER, _
        COL_ENROLLMENT_INN, COL_ENROLLMENT_SNILS, COL_ENROLLMENT_PASSPORT_SERIES, _
        COL_ENROLLMENT_PASSPORT_NUMBER, COL_ENROLLMENT_PASSPORT_CODE, COL_ENROLLMENT_BANK_ACCOUNT)

    For i = LBound(textColumns) To UBound(textColumns)
        ws.Columns(CLng(textColumns(i))).NumberFormat = "@"
    Next i
End Sub

Private Sub SyncEnrollmentWizardIfOpen()
    Dim frm As Object
    For Each frm In VBA.UserForms
        If TypeName(frm) = "frmEnrollmentWizard" Then
            frmEnrollmentWizard.ReloadFromBackend
            Exit For
        End If
    Next frm
End Sub

Private Function BuildEnrollmentId() As String
    BuildEnrollmentId = "ENR-" & Format$(Now, "yyyymmdd-hhnnss")
End Function

Private Function BuildEnrollmentOrderDraftId() As String
    BuildEnrollmentOrderDraftId = "ORD-" & Format$(Now, "yyyymmdd-hhnnss")
End Function

Private Function FindStaffDataForRecord(ByVal record As Object) As Object
    Dim numberValue As String
    Dim tableValue As String

    numberValue = SafeText(record("personal_number"))
    tableValue = SafeText(record("table_number"))

    If numberValue <> "" Then
        Set FindStaffDataForRecord = mdlHelper.FindEmployeeByAnyNumber(numberValue)
        If Not FindStaffDataForRecord Is Nothing Then
            If FindStaffDataForRecord.Count > 0 Then Exit Function
        End If
    End If

    If tableValue <> "" Then
        Set FindStaffDataForRecord = mdlHelper.FindEmployeeByAnyNumber(tableValue)
        Exit Function
    End If

    Set FindStaffDataForRecord = Nothing
End Function

Private Sub ValidateStaffConsistency(ByVal record As Object, ByRef issues As String, ByRef severity As Long)
    Dim staffData As Object

    Set staffData = FindStaffDataForRecord(record)
    If staffData Is Nothing Then Exit Sub
    If staffData.Count = 0 Then Exit Sub

    AppendStaffMismatchIssue record, staffData, issues, severity, "fio", mdlHelper.Ru(1051, 1080, 1094, 1086), ET("enrollment.field.fio", "FIO")
    AppendStaffMismatchIssue record, staffData, issues, severity, "rank", mdlHelper.Ru(1042, 1086, 1080, 1085, 1089, 1082, 1086, 1077, 32, 1079, 1074, 1072, 1085, 1080, 1077), ET("enrollment.field.rank", "Rank")
    AppendStaffMismatchIssue record, staffData, issues, severity, "position", mdlHelper.Ru(1064, 1090, 1072, 1090, 1085, 1072, 1103, 32, 1076, 1086, 1083, 1078, 1085, 1086, 1089, 1090, 1100), ET("enrollment.field.position", "Position")
    AppendStaffMismatchIssue record, staffData, issues, severity, "section", mdlHelper.Ru(1063, 1072, 1089, 1090, 1100), ET("enrollment.field.section", "Personnel section")
End Sub

Private Sub AppendStaffMismatchIssue(ByVal record As Object, ByVal staffData As Object, ByRef issues As String, ByRef severity As Long, ByVal recordFieldKey As String, ByVal staffFieldKey As String, ByVal fieldLabel As String)
    Dim recordValue As String
    Dim staffValue As String

    If Not staffData.Exists(staffFieldKey) Then Exit Sub
    recordValue = SafeText(record(recordFieldKey))
    staffValue = SafeText(staffData(staffFieldKey))
    If recordValue = "" Or staffValue = "" Then Exit Sub
    If NormalizeStaffCompareText(recordValue) = NormalizeStaffCompareText(staffValue) Then Exit Sub

    AppendIssue issues, tf("enrollment.issue.staff_mismatch", "Card data differs from Staff: {field}: '{value}' <> '{staff}'.", "{field}", fieldLabel, "{value}", recordValue, "{staff}", staffValue), severity, STATUS_WARNING
End Sub

Private Function NormalizeStaffCompareText(ByVal valueText As Variant) As String
    Dim resultText As String

    resultText = LCase$(SafeText(valueText))
    resultText = Replace$(resultText, ChrW$(160), " ")
    resultText = Replace$(resultText, vbCr, " ")
    resultText = Replace$(resultText, vbLf, " ")
    resultText = Replace$(resultText, vbTab, " ")
    Do While InStr(1, resultText, "  ", vbBinaryCompare) > 0
        resultText = Replace$(resultText, "  ", " ")
    Loop
    NormalizeStaffCompareText = Trim$(resultText)
End Function

Private Sub EnsureExtraBackendFields(ByVal ws As Worksheet)
    Dim i As Long

    For i = 1 To 4
        EnsureBackendField ws, ExtraMonthlyKey(i, "name"), "Доп. ежемесячная " & CStr(i) & ": наименование"
        EnsureBackendField ws, ExtraMonthlyKey(i, "param"), "Доп. ежемесячная " & CStr(i) & ": параметр"
        EnsureBackendField ws, ExtraMonthlyKey(i, "amount"), "Доп. ежемесячная " & CStr(i) & ": размер"
        EnsureBackendField ws, ExtraMonthlyKey(i, "start"), "Доп. ежемесячная " & CStr(i) & ": дата начала"
        EnsureBackendField ws, ExtraMonthlyKey(i, "basis"), "Доп. ежемесячная " & CStr(i) & ": основание"
        EnsureBackendField ws, ExtraMonthlyKey(i, "enabled"), "Доп. ежемесячная " & CStr(i) & ": вкл"
    Next i

    For i = 1 To 3
        EnsureBackendField ws, ExtraOneTimeKey(i, "name"), "Доп. разовая " & CStr(i) & ": наименование"
        EnsureBackendField ws, ExtraOneTimeKey(i, "amount"), "Доп. разовая " & CStr(i) & ": сумма"
        EnsureBackendField ws, ExtraOneTimeKey(i, "date"), "Доп. разовая " & CStr(i) & ": дата"
        EnsureBackendField ws, ExtraOneTimeKey(i, "basis"), "Доп. разовая " & CStr(i) & ": основание"
        EnsureBackendField ws, ExtraOneTimeKey(i, "enabled"), "Доп. разовая " & CStr(i) & ": вкл"
    Next i
End Sub

Private Sub NormalizeExtraPaymentFields(ByVal record As Object, ByVal standardStart As Date, ByVal enrollDate As Date)
    Dim i As Long
    Dim enabledKey As String
    Dim nameKey As String
    Dim amountKey As String
    Dim basisKey As String
    Dim dateKey As String

    For i = 1 To 4
        enabledKey = ExtraMonthlyKey(i, "enabled")
        nameKey = ExtraMonthlyKey(i, "name")
        amountKey = ExtraMonthlyKey(i, "amount")
        dateKey = ExtraMonthlyKey(i, "start")
        basisKey = ExtraMonthlyKey(i, "basis")

        If SafeText(record(enabledKey)) = "" Then
            If SafeText(record(nameKey)) <> "" Or SafeText(record(amountKey)) <> "" Or SafeText(record(ExtraMonthlyKey(i, "param"))) <> "" Or SafeText(record(basisKey)) <> "" Then
                record(enabledKey) = YES_VALUE
            Else
                record(enabledKey) = NO_VALUE
            End If
        End If

        If NormalizeYesNo(record(enabledKey)) = YES_VALUE Then
            If SafeText(record(dateKey)) = "" And standardStart > 0 Then record(dateKey) = Format$(standardStart, "dd.mm.yyyy")
            If SafeText(record(basisKey)) = "" Then record(basisKey) = SafeText(record("basis_section1"))
        End If
    Next i

    For i = 1 To 3
        enabledKey = ExtraOneTimeKey(i, "enabled")
        nameKey = ExtraOneTimeKey(i, "name")
        amountKey = ExtraOneTimeKey(i, "amount")
        dateKey = ExtraOneTimeKey(i, "date")
        basisKey = ExtraOneTimeKey(i, "basis")

        If SafeText(record(enabledKey)) = "" Then
            If SafeText(record(nameKey)) <> "" Or SafeText(record(amountKey)) <> "" Or SafeText(record(basisKey)) <> "" Then
                record(enabledKey) = YES_VALUE
            Else
                record(enabledKey) = NO_VALUE
            End If
        End If

        If NormalizeYesNo(record(enabledKey)) = YES_VALUE Then
            If SafeText(record(dateKey)) = "" And enrollDate > 0 Then record(dateKey) = Format$(enrollDate, "dd.mm.yyyy")
            If SafeText(record(basisKey)) = "" Then record(basisKey) = SafeText(record("basis_section1"))
        End If
    Next i
End Sub


Private Sub ValidateDateField(ByVal record As Object, ByVal fieldKey As String, ByVal fieldLabel As String, ByRef issues As String, ByRef severity As Long)
    If SafeText(record(fieldKey)) <> "" And Not IsDate(record(fieldKey)) Then
        AppendIssue issues, fieldLabel & ": укажите корректную дату в формате дд.мм.гггг.", severity, STATUS_BLOCKED
    End If
End Sub

Private Sub ValidateExactDigits(ByVal record As Object, ByVal fieldKey As String, ByVal fieldLabel As String, ByVal minLength As Long, ByVal maxLength As Long, ByRef issues As String, ByRef severity As Long, Optional ByVal allowSeparators As Boolean = False)
    Dim normalized As String
    Dim comparable As String

    normalized = DigitsOnly(SafeText(record(fieldKey)))
    comparable = SafeText(record(fieldKey))
    If allowSeparators Then comparable = Replace$(Replace$(comparable, "-", ""), " ", "")
    If SafeText(record(fieldKey)) <> "" And (normalized <> comparable Or Len(normalized) < minLength Or Len(normalized) > maxLength) Then
        If minLength = maxLength Then
            AppendIssue issues, fieldLabel & ": требуется ровно " & CStr(minLength) & " цифр.", severity, STATUS_BLOCKED
        Else
            AppendIssue issues, fieldLabel & ": требуется " & CStr(minLength) & " или " & CStr(maxLength) & " цифр.", severity, STATUS_BLOCKED
        End If
    End If
End Sub

Private Sub ValidatePassportCode(ByVal record As Object, ByRef issues As String, ByRef severity As Long)
    Dim normalized As String

    If SafeText(record("passport_code")) = "" Then Exit Sub
    normalized = Replace$(SafeText(record("passport_code")), "-", "")
    If Len(normalized) <> 6 Or DigitsOnly(normalized) <> normalized Then
        AppendIssue issues, "Код подразделения: укажите шесть цифр, например 123-456.", severity, STATUS_BLOCKED
    End If
End Sub

Private Function DigitsOnly(ByVal rawValue As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(rawValue)
        ch = Mid$(rawValue, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next i
End Function
Private Sub ValidateExtraPaymentFields(ByVal record As Object, ByRef issues As String, ByRef severity As Long)
    Dim i As Long

    For i = 1 To 4
        If NormalizeYesNo(record(ExtraMonthlyKey(i, "enabled"))) = YES_VALUE Then
            AppendIssueIfBlank issues, severity, record(ExtraMonthlyKey(i, "name")), tf("enrollment.issue.extra_monthly_name_missing", "Включена дополнительная ежемесячная выплата #{index}, но не заполнено наименование.", "{index}", i), STATUS_BLOCKED
            AppendIssueIfBlank issues, severity, record(ExtraMonthlyKey(i, "amount")), tf("enrollment.issue.extra_monthly_amount_missing", "Включена дополнительная ежемесячная выплата #{index}, но не заполнен размер.", "{index}", i), STATUS_BLOCKED
            AppendIssueIfBlank issues, severity, record(ExtraMonthlyKey(i, "start")), tf("enrollment.issue.extra_monthly_start_missing", "Для дополнительной ежемесячной выплаты #{index} не заполнена дата начала.", "{index}", i), STATUS_WARNING
            AppendIssueIfBlank issues, severity, record(ExtraMonthlyKey(i, "basis")), tf("enrollment.issue.extra_monthly_basis_missing", "Для дополнительной ежемесячной выплаты #{index} не заполнено основание.", "{index}", i), STATUS_WARNING
        End If
    Next i

    For i = 1 To 3
        If NormalizeYesNo(record(ExtraOneTimeKey(i, "enabled"))) = YES_VALUE Then
            AppendIssueIfBlank issues, severity, record(ExtraOneTimeKey(i, "name")), tf("enrollment.issue.extra_onetime_name_missing", "Включена дополнительная разовая выплата #{index}, но не заполнено наименование.", "{index}", i), STATUS_BLOCKED
            AppendIssueIfBlank issues, severity, record(ExtraOneTimeKey(i, "amount")), tf("enrollment.issue.extra_onetime_amount_missing", "Для дополнительной разовой выплаты #{index} не заполнена сумма.", "{index}", i), STATUS_BLOCKED
            AppendIssueIfBlank issues, severity, record(ExtraOneTimeKey(i, "date")), tf("enrollment.issue.extra_onetime_date_missing", "Для дополнительной разовой выплаты #{index} не заполнена дата.", "{index}", i), STATUS_BLOCKED
            AppendIssueIfBlank issues, severity, record(ExtraOneTimeKey(i, "basis")), tf("enrollment.issue.extra_onetime_basis_missing", "Для дополнительной разовой выплаты #{index} не заполнено основание.", "{index}", i), STATUS_WARNING
        End If
    Next i
End Sub

Private Sub LoadExtraPaymentFieldsFromWorksheet(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal record As Object)
    Dim i As Long

    For i = 1 To 4
        record(ExtraMonthlyKey(i, "name")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "name")).Value
        record(ExtraMonthlyKey(i, "param")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "param")).Value
        record(ExtraMonthlyKey(i, "amount")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "amount")).Value
        record(ExtraMonthlyKey(i, "start")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "start")).Value
        record(ExtraMonthlyKey(i, "basis")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "basis")).Value
        record(ExtraMonthlyKey(i, "enabled")) = ws.Cells(rowNum, ExtraMonthlyColumn(i, "enabled")).Value
    Next i

    For i = 1 To 3
        record(ExtraOneTimeKey(i, "name")) = ws.Cells(rowNum, ExtraOneTimeColumn(i, "name")).Value
        record(ExtraOneTimeKey(i, "amount")) = ws.Cells(rowNum, ExtraOneTimeColumn(i, "amount")).Value
        record(ExtraOneTimeKey(i, "date")) = ws.Cells(rowNum, ExtraOneTimeColumn(i, "date")).Value
        record(ExtraOneTimeKey(i, "basis")) = ws.Cells(rowNum, ExtraOneTimeColumn(i, "basis")).Value
        record(ExtraOneTimeKey(i, "enabled")) = ws.Cells(rowNum, ExtraOneTimeColumn(i, "enabled")).Value
    Next i
End Sub

Private Sub WriteExtraPaymentFieldsToWorksheet(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal record As Object)
    Dim i As Long

    For i = 1 To 4
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "name")).Value = record(ExtraMonthlyKey(i, "name"))
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "param")).Value = record(ExtraMonthlyKey(i, "param"))
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "amount")).Value = record(ExtraMonthlyKey(i, "amount"))
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "start")).Value = record(ExtraMonthlyKey(i, "start"))
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "basis")).Value = record(ExtraMonthlyKey(i, "basis"))
        ws.Cells(rowNum, ExtraMonthlyColumn(i, "enabled")).Value = record(ExtraMonthlyKey(i, "enabled"))
    Next i

    For i = 1 To 3
        ws.Cells(rowNum, ExtraOneTimeColumn(i, "name")).Value = record(ExtraOneTimeKey(i, "name"))
        ws.Cells(rowNum, ExtraOneTimeColumn(i, "amount")).Value = record(ExtraOneTimeKey(i, "amount"))
        ws.Cells(rowNum, ExtraOneTimeColumn(i, "date")).Value = record(ExtraOneTimeKey(i, "date"))
        ws.Cells(rowNum, ExtraOneTimeColumn(i, "basis")).Value = record(ExtraOneTimeKey(i, "basis"))
        ws.Cells(rowNum, ExtraOneTimeColumn(i, "enabled")).Value = record(ExtraOneTimeKey(i, "enabled"))
    Next i
End Sub

Private Function AppendExtraPaymentRows(ByVal wsPayments As Worksheet, ByVal record As Object, ByVal packageId As String) As Long
    Dim i As Long
    Dim paymentType As String

    For i = 1 To 4
        If NormalizeYesNo(record(ExtraMonthlyKey(i, "enabled"))) = YES_VALUE Then
            paymentType = SafeText(record(ExtraMonthlyKey(i, "name")))
            If paymentType = "" Then paymentType = tf("payments.type.extra_monthly", "Дополнительная выплата #{index}", "{index}", i)
            AppendPaymentRowFromEnrollment wsPayments, record, packageId, paymentType, SafeText(record(ExtraMonthlyKey(i, "param"))), SafeText(record(ExtraMonthlyKey(i, "amount"))), SafeText(record(ExtraMonthlyKey(i, "basis")))
            AppendExtraPaymentRows = AppendExtraPaymentRows + 1
        End If
    Next i

    For i = 1 To 3
        If NormalizeYesNo(record(ExtraOneTimeKey(i, "enabled"))) = YES_VALUE Then
            paymentType = SafeText(record(ExtraOneTimeKey(i, "name")))
            If paymentType = "" Then paymentType = tf("payments.type.extra_onetime", "Дополнительная разовая выплата #{index}", "{index}", i)
            AppendPaymentRowFromEnrollment wsPayments, record, packageId, paymentType, "", SafeText(record(ExtraOneTimeKey(i, "amount"))), SafeText(record(ExtraOneTimeKey(i, "basis")))
            AppendExtraPaymentRows = AppendExtraPaymentRows + 1
        End If
    Next i
End Function

Private Sub EnsureExtraRecordKeys(ByVal record As Object)
    Dim i As Long

    For i = 1 To 4
        EnsureRecordKey record, ExtraMonthlyKey(i, "name")
        EnsureRecordKey record, ExtraMonthlyKey(i, "param")
        EnsureRecordKey record, ExtraMonthlyKey(i, "amount")
        EnsureRecordKey record, ExtraMonthlyKey(i, "start")
        EnsureRecordKey record, ExtraMonthlyKey(i, "basis")
        EnsureRecordKey record, ExtraMonthlyKey(i, "enabled")
    Next i

    For i = 1 To 3
        EnsureRecordKey record, ExtraOneTimeKey(i, "name")
        EnsureRecordKey record, ExtraOneTimeKey(i, "amount")
        EnsureRecordKey record, ExtraOneTimeKey(i, "date")
        EnsureRecordKey record, ExtraOneTimeKey(i, "basis")
        EnsureRecordKey record, ExtraOneTimeKey(i, "enabled")
    Next i
End Sub

Private Function ExtraMonthlyKey(ByVal index As Long, ByVal fieldName As String) As String
    ExtraMonthlyKey = "extra_monthly" & CStr(index) & "_" & fieldName
End Function

Private Function ExtraOneTimeKey(ByVal index As Long, ByVal fieldName As String) As String
    ExtraOneTimeKey = "extra_one_time" & CStr(index) & "_" & fieldName
End Function

Private Function ExtraMonthlyColumn(ByVal index As Long, ByVal fieldName As String) As Long
    Select Case index
        Case 1
            Select Case LCase$(fieldName)
                Case "name": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_NAME
                Case "param": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_PARAM
                Case "amount": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_AMOUNT
                Case "start": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_START
                Case "basis": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_BASIS
                Case "enabled": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY1_ENABLED
            End Select
        Case 2
            Select Case LCase$(fieldName)
                Case "name": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_NAME
                Case "param": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_PARAM
                Case "amount": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_AMOUNT
                Case "start": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_START
                Case "basis": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_BASIS
                Case "enabled": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY2_ENABLED
            End Select
        Case 3
            Select Case LCase$(fieldName)
                Case "name": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_NAME
                Case "param": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_PARAM
                Case "amount": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_AMOUNT
                Case "start": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_START
                Case "basis": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_BASIS
                Case "enabled": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY3_ENABLED
            End Select
        Case 4
            Select Case LCase$(fieldName)
                Case "name": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_NAME
                Case "param": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_PARAM
                Case "amount": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_AMOUNT
                Case "start": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_START
                Case "basis": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_BASIS
                Case "enabled": ExtraMonthlyColumn = COL_ENROLLMENT_EXTRA_MONTHLY4_ENABLED
            End Select
    End Select
End Function

Private Function ExtraOneTimeColumn(ByVal index As Long, ByVal fieldName As String) As Long
    Select Case index
        Case 1
            Select Case LCase$(fieldName)
                Case "name": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME1_NAME
                Case "amount": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME1_AMOUNT
                Case "date": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME1_DATE
                Case "basis": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME1_BASIS
                Case "enabled": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME1_ENABLED
            End Select
        Case 2
            Select Case LCase$(fieldName)
                Case "name": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME2_NAME
                Case "amount": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME2_AMOUNT
                Case "date": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME2_DATE
                Case "basis": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME2_BASIS
                Case "enabled": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME2_ENABLED
            End Select
        Case 3
            Select Case LCase$(fieldName)
                Case "name": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME3_NAME
                Case "amount": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME3_AMOUNT
                Case "date": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME3_DATE
                Case "basis": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME3_BASIS
                Case "enabled": ExtraOneTimeColumn = COL_ENROLLMENT_EXTRA_ONE_TIME3_ENABLED
            End Select
    End Select
End Function

Private Sub EnsureRecordKeys(ByVal record As Object)
    EnsureRecordKey record, "current_row"
    EnsureRecordKey record, "enrollment_id"
    EnsureRecordKey record, "order_draft_id"
    EnsureRecordKey record, "source_mode"
    EnsureRecordKey record, "fio"
    EnsureRecordKey record, "personal_number"
    EnsureRecordKey record, "table_number"
    EnsureRecordKey record, "rank"
    EnsureRecordKey record, "service_category"
    EnsureRecordKey record, "contract_kind"
    EnsureRecordKey record, "contract_basis"
    EnsureRecordKey record, "vus"
    EnsureRecordKey record, "position"
    EnsureRecordKey record, "section"
    EnsureRecordKey record, "military_unit"
    EnsureRecordKey record, "tariff_rank"
    EnsureRecordKey record, "position_salary"
    EnsureRecordKey record, "rank_salary"
    EnsureRecordKey record, "order_date"
    EnsureRecordKey record, "order_number"
    EnsureRecordKey record, "order_issuer"
    EnsureRecordKey record, "arrival_source"
    EnsureRecordKey record, "prescription_number"
    EnsureRecordKey record, "prescription_date"
    EnsureRecordKey record, "report_number"
    EnsureRecordKey record, "report_date"
    EnsureRecordKey record, "report_info"
    EnsureRecordKey record, "assignment_info"
    EnsureRecordKey record, "accept_date"
    EnsureRecordKey record, "enroll_date"
    EnsureRecordKey record, "duty_start_date"
    EnsureRecordKey record, "manual_start_date"
    EnsureRecordKey record, "standard_start_date"
    EnsureRecordKey record, "preferential_start_date"
    EnsureRecordKey record, "basis_section1"
    EnsureRecordKey record, "basis_section2"
    EnsureRecordKey record, "birth_date"
    EnsureRecordKey record, "birth_place"
    EnsureRecordKey record, "citizenship"
    EnsureRecordKey record, "inn"
    EnsureRecordKey record, "snils"
    EnsureRecordKey record, "passport_series"
    EnsureRecordKey record, "passport_number"
    EnsureRecordKey record, "passport_issuer"
    EnsureRecordKey record, "passport_issue_date"
    EnsureRecordKey record, "passport_code"
    EnsureRecordKey record, "bank_account"
    EnsureRecordKey record, "bank_name"
    EnsureRecordKey record, "requisites_note"
    EnsureRecordKey record, "preferential_enabled"
    EnsureRecordKey record, "preferential_coeff"
    EnsureRecordKey record, "preferential_basis"
    EnsureRecordKey record, "premium_enabled"
    EnsureRecordKey record, "premium_percent"
    EnsureRecordKey record, "premium_start"
    EnsureRecordKey record, "premium_end"
    EnsureRecordKey record, "premium_basis"
    EnsureRecordKey record, "lift_enabled"
    EnsureRecordKey record, "lift_amount"
    EnsureRecordKey record, "lift_date"
    EnsureRecordKey record, "lift_basis"
    EnsureRecordKey record, "per_diem_enabled"
    EnsureRecordKey record, "per_diem_days"
    EnsureRecordKey record, "per_diem_amount"
    EnsureRecordKey record, "per_diem_date"
    EnsureRecordKey record, "per_diem_basis"
    EnsureRecordKey record, "edv_enabled"
    EnsureRecordKey record, "edv_amount"
    EnsureRecordKey record, "edv_date"
    EnsureRecordKey record, "edv_basis"
    EnsureRecordKey record, "class_param"
    EnsureRecordKey record, "class_enabled"
    EnsureRecordKey record, "class_percent"
    EnsureRecordKey record, "class_basis"
    EnsureRecordKey record, "fizo_param"
    EnsureRecordKey record, "fizo_enabled"
    EnsureRecordKey record, "fizo_percent"
    EnsureRecordKey record, "fizo_basis"
    EnsureRecordKey record, "secrecy_param"
    EnsureRecordKey record, "secrecy_enabled"
    EnsureRecordKey record, "secrecy_percent"
    EnsureRecordKey record, "secrecy_basis"
    EnsureRecordKey record, "achievement_param"
    EnsureRecordKey record, "achievement_enabled"
    EnsureRecordKey record, "achievement_amount"
    EnsureRecordKey record, "achievement_basis"
    EnsureRecordKey record, "achievement_award_date"
    EnsureRecordKey record, "achievement_document_reference"
    EnsureRecordKey record, "std_duty_enabled"
    EnsureRecordKey record, "std_duty_percent"
    EnsureRecordKey record, "std_duty_date"
    EnsureRecordKey record, "std_duty_basis"
    EnsureRecordKey record, "std_special_enabled"
    EnsureRecordKey record, "std_special_percent"
    EnsureRecordKey record, "std_special_date"
    EnsureRecordKey record, "std_special_basis"
    EnsureRecordKey record, "std_tariff_enabled"
    EnsureRecordKey record, "std_tariff_percent"
    EnsureRecordKey record, "std_tariff_date"
    EnsureRecordKey record, "std_tariff_basis"
    EnsureRecordKey record, "std_contract430_enabled"
    EnsureRecordKey record, "std_contract430_percent"
    EnsureRecordKey record, "std_contract430_date"
    EnsureRecordKey record, "std_contract430_basis"
    EnsureExtraRecordKeys record
    EnsureRecordKey record, "standard_types"
    EnsureRecordKey record, "payment_basis"
    EnsureRecordKey record, "last_derived_at"
End Sub

Private Sub EnsureRecordKey(ByVal record As Object, ByVal fieldKey As String)
    If Not record.Exists(fieldKey) Then record(fieldKey) = vbNullString
End Sub

Private Function IsEnabledRecord(ByVal record As Object, ByVal fieldKey As String) As Boolean
    IsEnabledRecord = NormalizeYesNo(record(fieldKey)) = YES_VALUE
End Function

Private Function NormalizeYesNo(ByVal valueText As Variant) As String
    Dim normalized As String

    normalized = UCase$(Trim$(CStr(valueText)))
    Select Case normalized
        Case "1", "TRUE", "YES", "ДА", "Y"
            NormalizeYesNo = YES_VALUE
        Case Else
            NormalizeYesNo = NO_VALUE
    End Select
End Function

Private Function SeverityToStatusText(ByVal severity As Long) As String
    Select Case severity
        Case STATUS_BLOCKED
            SeverityToStatusText = ET("enrollment.status.blocked", "Блокировано")
        Case STATUS_WARNING
            SeverityToStatusText = ET("enrollment.status.warning", "Предупреждение")
        Case Else
            SeverityToStatusText = ET("enrollment.status.ready", "Готово")
    End Select
End Function

Private Sub AppendIssueIfBlank(ByRef issues As String, ByRef severity As Long, ByVal valueText As Variant, ByVal issueText As String, ByVal issueSeverity As Long)
    If SafeText(valueText) = "" Then
        AppendIssue issues, issueText, severity, issueSeverity
    End If
End Sub

Private Sub AppendIssue(ByRef issues As String, ByVal issueText As String, ByRef severity As Long, ByVal issueSeverity As Long)
    If issueText = "" Then Exit Sub
    If issues <> "" Then issues = issues & vbCrLf
    issues = issues & "- " & issueText
    If issueSeverity > severity Then severity = issueSeverity
End Sub

Private Sub ValidatePersonalPaymentAmounts(ByVal record As Object, ByRef issues As String, ByRef severity As Long)
    If IsEnabledRecord(record, "class_enabled") And IsBlankOrZeroAmount(record("class_percent")) Then
        AppendIssue issues, ET("enrollment.issue.class_percent_missing", "Включена классность, но не определён размер надбавки."), severity, STATUS_BLOCKED
    End If

    If IsEnabledRecord(record, "fizo_enabled") And IsBlankOrZeroAmount(record("fizo_percent")) Then
        AppendIssue issues, ET("enrollment.issue.fizo_percent_missing", "Включено ФИЗО, но не определён размер надбавки."), severity, STATUS_BLOCKED
    End If

    If IsEnabledRecord(record, "secrecy_enabled") And IsBlankOrZeroAmount(record("secrecy_percent")) Then
        AppendIssue issues, ET("enrollment.issue.secrecy_percent_missing", "Включена секретность, но не определён размер надбавки."), severity, STATUS_BLOCKED
    End If

    If IsEnabledRecord(record, "achievement_enabled") And IsBlankOrZeroAmount(record("achievement_amount")) Then
        AppendIssue issues, ET("enrollment.issue.achievement_amount_missing", "Включены особые достижения, но не заполнен размер выплаты."), severity, STATUS_BLOCKED
    End If
    If IsMedalAchievement(record) Then
        AppendIssueIfBlank issues, severity, record("achievement_award_date"), "A medal requires the award-order date.", STATUS_BLOCKED
        AppendIssueIfBlank issues, severity, record("achievement_document_reference"), "A medal requires the award-order number or reference.", STATUS_BLOCKED
    End If
End Sub

Private Function IsMedalAchievement(ByVal record As Object) As Boolean
    Select Case UCase$(GetEnrollmentReferenceCode("ACHIEVEMENT", SafeText(record("achievement_param"))))
        Case "COMBAT_DISTINCTION", "DEMINING", "MILITARY_VALOR_I", "MILITARY_VALOR_II"
            IsMedalAchievement = True
    End Select
End Function

Private Function IsBlankOrZeroAmount(ByVal valueText As Variant) As Boolean
    Dim normalized As String

    normalized = Trim$(CStr(valueText))
    normalized = Replace$(normalized, "%", "")
    normalized = Replace$(normalized, " ", "")

    If normalized = "" Then
        IsBlankOrZeroAmount = True
        Exit Function
    End If

    normalized = Replace$(normalized, ".", Application.DecimalSeparator)
    normalized = Replace$(normalized, ",", Application.DecimalSeparator)
    If IsNumeric(normalized) Then
        IsBlankOrZeroAmount = (CDbl(normalized) = 0)
    End If
End Function

Private Sub ValidateConfiguredPaymentDefinitions(ByVal record As Object, ByRef issues As String, ByRef severity As Long)
    Dim defCode As Variant
    Dim definition As Object
    Dim requiredTokens() As String
    Dim i As Long
    Dim token As String
    Dim fieldKey As String
    Dim issueText As String
    Dim issueSeverity As Long

    For Each defCode In EnrollmentDefinitionCodes()
        If CStr(defCode) <> "extra_monthly" And CStr(defCode) <> "extra_onetime" Then
            Set definition = GetEnrollmentPaymentDefinition(CStr(defCode))
            If IsConfiguredPaymentActive(record, definition) Then
                issueSeverity = DefinitionValidationSeverity(definition)
                requiredTokens = Split(SafeText(definition("required_docs")), ",")
                For i = LBound(requiredTokens) To UBound(requiredTokens)
                    token = Trim$(requiredTokens(i))
                    If token <> "" Then
                        fieldKey = ResolveDefinitionRequiredFieldKey(definition, token)
                        If fieldKey <> "" Then
                            If SafeText(record(fieldKey)) = "" Then
                                issueText = ResolveDefinitionIssueText(definition, token)
                                If issueText = "" Then issueText = BuildDefinitionRequiredFieldMessage(definition, token)
                                AppendIssue issues, issueText, severity, issueSeverity
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next defCode
End Sub

Private Function ResolveDefinitionIssueText(ByVal definition As Object, ByVal token As String) As String
    Select Case SafeText(definition("code")) & "|" & LCase$(token)
        Case "premium|premium_end"
            ResolveDefinitionIssueText = ET("enrollment.issue.premium_end_missing", "Не заполнена дата окончания премии.")
        Case "premium|basis"
            ResolveDefinitionIssueText = ET("enrollment.issue.premium_basis_missing", "Включена премия, но не заполнено основание премирования.")
        Case "class|param"
            ResolveDefinitionIssueText = ET("enrollment.issue.class_param_missing", "Включена классность, но не заполнен параметр.")
        Case "fizo|param"
            ResolveDefinitionIssueText = ET("enrollment.issue.fizo_param_missing", "Включено ФИЗО, но не заполнен параметр.")
        Case "secrecy|param"
            ResolveDefinitionIssueText = ET("enrollment.issue.secrecy_param_missing", "Включена секретность, но не заполнен параметр.")
        Case "achievement|param"
            ResolveDefinitionIssueText = ET("enrollment.issue.achievement_param_missing", "Включены особые достижения, но не заполнен параметр.")
        Case "edv|basis_section2"
            ResolveDefinitionIssueText = ET("enrollment.issue.basis_section2_missing", "Включена ЕДВ, но не заполнен блок оснований для §2.")
        Case "edv|contract_basis"
            ResolveDefinitionIssueText = ET("enrollment.issue.edv_contract_basis_required", "Для ЕДВ нужно заполнить основание контракта / нормативный блок.")
    End Select
End Function

Private Function BuildDefinitionRequiredFieldMessage(ByVal definition As Object, ByVal token As String) As String
    Dim templateText As String

    templateText = ET("enrollment.issue.payment_field_missing", "Для выплаты ""{label}"" не заполнено поле ""{field}"".")
    templateText = Replace$(templateText, "{label}", SafeText(definition("label")))
    BuildDefinitionRequiredFieldMessage = Replace$(templateText, "{field}", ResolveDefinitionRequiredFieldLabel(token))
End Function

Private Function ResolveDefinitionRequiredFieldLabel(ByVal token As String) As String
    Select Case LCase$(token)
        Case "amount"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.amount", "размер")
        Case "date"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.date", "дата")
        Case "basis"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.basis", "основание")
        Case "param"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.param", "параметр")
        Case "name"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.name", "наименование")
        Case "start"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.start", "дата начала")
        Case "premium_end"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.premium_end", "дата окончания премии")
        Case "basis_section2"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.basis_section2", "основание для §2")
        Case "contract_basis"
            ResolveDefinitionRequiredFieldLabel = ET("enrollment.field.contract_basis", "основание контракта")
        Case Else
            ResolveDefinitionRequiredFieldLabel = token
    End Select
End Function

Private Function ResolveDefinitionRequiredFieldKey(ByVal definition As Object, ByVal token As String) As String
    Dim bindingKey As String

    bindingKey = SafeText(definition("journal_binding"))
    Select Case LCase$(token)
        Case "basis"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_basis"
        Case "param"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_param"
        Case "amount"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_amount"
        Case "date"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_date"
        Case "name"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_name"
        Case "start"
            ResolveDefinitionRequiredFieldKey = bindingKey & "_start"
        Case "premium_end"
            ResolveDefinitionRequiredFieldKey = "premium_end"
        Case "basis_section2"
            ResolveDefinitionRequiredFieldKey = "basis_section2"
        Case "contract_basis"
            ResolveDefinitionRequiredFieldKey = "contract_basis"
    End Select
End Function

Private Function DefinitionValidationSeverity(ByVal definition As Object) As Long
    Select Case LCase$(SafeText(definition("validation_severity")))
        Case "blocked"
            DefinitionValidationSeverity = STATUS_BLOCKED
        Case "warning"
            DefinitionValidationSeverity = STATUS_WARNING
        Case Else
            DefinitionValidationSeverity = STATUS_READY
    End Select
End Function

Private Sub ApplyConfiguredPaymentStartDates(ByVal record As Object, ByVal enrollDate As Date, ByVal standardStart As Date)
    Dim defCode As Variant
    Dim definition As Object
    Dim fieldKey As String
    Dim defaultDate As Date

    For Each defCode In EnrollmentDefinitionCodes()
        If CStr(defCode) <> "extra_monthly" And CStr(defCode) <> "extra_onetime" Then
            Set definition = GetEnrollmentPaymentDefinition(CStr(defCode))
            fieldKey = ResolveConfiguredDateFieldKey(definition)
            If fieldKey <> "" Then
                If SafeText(record(fieldKey)) = "" Then
                    If IsConfiguredPaymentActive(record, definition) Then
                        defaultDate = ResolveConfiguredStartDate(definition, record, enrollDate, standardStart)
                        If defaultDate > 0 Then record(fieldKey) = Format$(defaultDate, "dd.mm.yyyy")
                    End If
                End If
            End If
        End If
    Next defCode
End Sub

Private Function ResolveConfiguredDateFieldKey(ByVal definition As Object) As String
    Select Case SafeText(definition("code"))
        Case "premium"
            ResolveConfiguredDateFieldKey = "premium_start"
        Case "std_duty", "std_special", "std_tariff", "std_contract430", "lift", "per_diem", "edv"
            ResolveConfiguredDateFieldKey = SafeText(definition("journal_binding")) & "_date"
    End Select
End Function

Private Function ResolveConfiguredStartDate(ByVal definition As Object, ByVal record As Object, ByVal enrollDate As Date, ByVal standardStart As Date) As Date
    Dim sourceKey As String

    sourceKey = LCase$(SafeText(definition("start_date_source")))
    Select Case sourceKey
        Case "standard_start_date"
            ResolveConfiguredStartDate = standardStart
            If ResolveConfiguredStartDate = 0 Then ResolveConfiguredStartDate = enrollDate
        Case "manual", "manual_start_date"
            ResolveConfiguredStartDate = mdlHelper.ParseDateSafe(record("manual_start_date"))
        Case "enroll_date"
            ResolveConfiguredStartDate = enrollDate
        Case "accept", "accept_date", "duty_acceptance_date"
            ResolveConfiguredStartDate = mdlHelper.ParseDateSafe(record("accept_date"))
        Case "order", "order_date", "personnel_order_date"
            ResolveConfiguredStartDate = mdlHelper.ParseDateSafe(record("order_date"))
        Case Else
            If sourceKey <> "" Then
                If record.Exists(sourceKey) Then
                    ResolveConfiguredStartDate = mdlHelper.ParseDateSafe(record(sourceKey))
                End If
            End If
    End Select
End Function

Public Function IsConfiguredPaymentActive(ByVal record As Object, ByVal definition As Object) As Boolean
    Dim enabledFieldKey As String

    enabledFieldKey = SafeText(definition("journal_binding")) & "_enabled"
    If enabledFieldKey = "_enabled" Then Exit Function
    If Not record.Exists(enabledFieldKey) Then Exit Function

    IsConfiguredPaymentActive = NormalizeYesNo(record(enabledFieldKey)) = YES_VALUE
End Function

Private Function IsAllBlank(ParamArray values() As Variant) As Boolean
    Dim i As Long

    IsAllBlank = True
    For i = LBound(values) To UBound(values)
        If SafeText(values(i)) <> "" Then
            IsAllBlank = False
            Exit Function
        End If
    Next i
End Function

Private Function JoinEnabledStandardTypes(ByVal record As Object) As String
    Dim resultText As String

    If IsEnabledRecord(record, "std_duty_enabled") Then resultText = AppendListValue(resultText, ET("enrollment.std.duty", "Надбавка по воинской должности"))
    If IsEnabledRecord(record, "std_special_enabled") Then resultText = AppendListValue(resultText, ET("enrollment.std.special", "Особые условия"))
    If IsEnabledRecord(record, "std_tariff_enabled") Then resultText = AppendListValue(resultText, ET("enrollment.std.tariff", "1-4 тарифный разряд"))
    If IsEnabledRecord(record, "std_contract430_enabled") Then resultText = AppendListValue(resultText, ET("enrollment.std.contract430", "430 ДСП"))

    JoinEnabledStandardTypes = resultText
End Function

Private Function AppendListValue(ByVal sourceText As String, ByVal valueText As String) As String
    If Trim$(valueText) = "" Then
        AppendListValue = sourceText
    ElseIf Trim$(sourceText) = "" Then
        AppendListValue = valueText
    Else
        AppendListValue = sourceText & "; " & valueText
    End If
End Function

Private Function IsTariffOneToFour(ByVal tariffText As Variant) As Boolean
    Dim normalized As String
    normalized = Trim$(CStr(tariffText))
    If Not IsNumeric(normalized) Then Exit Function
    IsTariffOneToFour = CLng(normalized) >= 1 And CLng(normalized) <= 4
End Function

Private Function ExtractContractDate(ByVal record As Object) As Date
    ExtractContractDate = mdlHelper.ParseDateSafe(record("contract_basis"))
    If ExtractContractDate = 0 Then ExtractContractDate = mdlHelper.ParseDateSafe(record("order_date"))
End Function

Private Function IsContract430Candidate(ByVal record As Object, ByVal contractDate As Date) As Boolean
    If SafeText(record("contract_kind")) <> "" Then
        If contractDate = 0 Or contractDate >= DateSerial(2022, 9, 21) Then
            IsContract430Candidate = True
            Exit Function
        End If
    End If

    If InStr(1, SafeText(record("contract_basis")), "21.09.2022", vbTextCompare) > 0 Then
        IsContract430Candidate = True
    End If
End Function

Private Function ResolveClassPercent(ByVal paramText As String) As Long
    Dim normalized As String
    normalized = UCase$(Trim$(paramText))

    If InStr(normalized, "МАСТ") > 0 Then
        ResolveClassPercent = 30
    ElseIf InStr(normalized, "1") > 0 Or InStr(normalized, "I") > 0 Then
        ResolveClassPercent = 20
    ElseIf InStr(normalized, "2") > 0 Or InStr(normalized, "II") > 0 Then
        ResolveClassPercent = 10
    ElseIf InStr(normalized, "3") > 0 Or InStr(normalized, "III") > 0 Then
        ResolveClassPercent = 5
    End If
End Function

Private Function ResolveFizoPercent(ByVal paramText As String) As Long
    Dim normalized As String
    normalized = UCase$(Trim$(paramText))

    If InStr(normalized, "ВЫС") > 0 Then
        ResolveFizoPercent = 70
    ElseIf InStr(normalized, "1") > 0 Or InStr(normalized, "I") > 0 Then
        ResolveFizoPercent = 50
    ElseIf InStr(normalized, "2") > 0 Or InStr(normalized, "II") > 0 Then
        ResolveFizoPercent = 30
    ElseIf InStr(normalized, "3") > 0 Or InStr(normalized, "III") > 0 Then
        ResolveFizoPercent = 15
    End If
End Function

Private Function ResolveSecrecyPercent(ByVal paramText As String) As Long
    Dim normalized As String
    normalized = UCase$(Trim$(paramText))

    If InStr(normalized, "1") > 0 Then
        ResolveSecrecyPercent = 25
    ElseIf InStr(normalized, "2") > 0 Then
        ResolveSecrecyPercent = 20
    ElseIf InStr(normalized, "3") > 0 Then
        ResolveSecrecyPercent = 10
    Else
        ResolveSecrecyPercent = 10
    End If
End Function

Private Function ResolveAchievementAmountText(ByVal paramText As String) As String
    If IsNumeric(Trim$(paramText)) Then
        ResolveAchievementAmountText = Trim$(paramText)
    Else
        ResolveAchievementAmountText = "60"
    End If
End Function

Private Function BuildSection1Basis(ByVal record As Object) As String
    Dim resultText As String

    If SafeText(record("order_issuer")) <> "" Or SafeText(record("order_number")) <> "" Or SafeText(record("order_date")) <> "" Then
        resultText = ET("enrollment.basis.order_prefix", "Приказ") & " " & SafeText(record("order_issuer"))
        If SafeText(record("order_date")) <> "" Then resultText = Trim$(resultText & " от " & SafeText(record("order_date")))
        If SafeText(record("order_number")) <> "" Then resultText = Trim$(resultText & " № " & SafeText(record("order_number")))
    End If
    If SafeText(record("prescription_number")) <> "" Or SafeText(record("prescription_date")) <> "" Then
        resultText = AppendListValue(resultText, ET("enrollment.basis.prescription_prefix", "предписание") & " № " & SafeText(record("prescription_number")) & IIf(SafeText(record("prescription_date")) <> "", " от " & SafeText(record("prescription_date")), ""))
    ElseIf SafeText(record("assignment_info")) <> "" Then
        resultText = AppendListValue(resultText, SafeText(record("assignment_info")))
    End If
    If SafeText(record("report_number")) <> "" Or SafeText(record("report_date")) <> "" Then
        resultText = AppendListValue(resultText, ET("enrollment.basis.report_prefix", "рапорт") & " № " & SafeText(record("report_number")) & IIf(SafeText(record("report_date")) <> "", " от " & SafeText(record("report_date")), ""))
    ElseIf SafeText(record("report_info")) <> "" Then
        resultText = AppendListValue(resultText, SafeText(record("report_info")))
    End If

    BuildSection1Basis = resultText
End Function

Private Function BuildSection2Basis(ByVal record As Object) As String
    Dim resultText As String
    resultText = BuildSection1Basis(record)
    If SafeText(record("contract_basis")) <> "" Then
        resultText = AppendListValue(resultText, SafeText(record("contract_basis")))
    End If
    BuildSection2Basis = resultText
End Function

Private Function DefaultLiftAmount() As String
    DefaultLiftAmount = "1 " & mdlHelper.Ru(1054, 1044, 1057)
End Function

Private Function DefaultPerDiemAmount() As String
    DefaultPerDiemAmount = "1 " & mdlHelper.Ru(1089, 1091, 1090, 1082, 1080)
End Function

Private Sub EnsureEnrollmentSettings()
    Dim ws As Worksheet

    Set ws = GetSettingsWorksheet()
    If ws Is Nothing Then Exit Sub

    EnsureSetting ws, "enrollment.unit_number", "81510", "Military unit number for enrollment order header"
    EnsureSetting ws, "enrollment.city", EnrollmentDefaultCity(), "City in the enrollment order header"
    EnsureSetting ws, "enrollment.signatory_name", EnrollmentDefaultSignatoryName(), "Signatory full name"
    EnsureSetting ws, "enrollment.signatory_rank", EnrollmentDefaultSignatoryRank(), "Signatory military rank"
    EnsureSetting ws, "enrollment.signatory_position", EnrollmentDefaultSignatoryPosition(), "Signatory position"
    EnsureSetting ws, "enrollment.header_text", EnrollmentDefaultHeaderText(), "Order header lines separated by |"
    EnsureSetting ws, "enrollment.filename_template", EnrollmentDefaultFilenameTemplate(), "Output file name template"
    EnsureSetting ws, "enrollment.template_file", EnrollmentDefaultTemplateFileName(), "Enrollment Word template file"
    EnsureSetting ws, "enrollment.template_body_marker", "[ENROLLMENT_ORDER_BODY]", "Word template marker replaced by generated enrollment order body"
    EnsureEnrollmentPaymentDefinitions ws
End Sub

Public Function GetEnrollmentSetting(ByVal settingKey As String, ByVal defaultValue As String) As String
    Dim ws As Worksheet
    Dim rowNum As Long

    Set ws = GetSettingsWorksheet()
    If ws Is Nothing Then
        GetEnrollmentSetting = defaultValue
        Exit Function
    End If

    rowNum = FindSettingRow(ws, settingKey)
    If rowNum = 0 Then
        GetEnrollmentSetting = defaultValue
    Else
        GetEnrollmentSetting = SafeText(ws.Cells(rowNum, 2).Value)
        If GetEnrollmentSetting = "" Then GetEnrollmentSetting = defaultValue
    End If
End Function

Private Function GetSettingsWorksheet() As Worksheet
    On Error Resume Next
    Set GetSettingsWorksheet = ThisWorkbook.Worksheets(mdlHelper.Ru(1053, 1072, 1089, 1090, 1088, 1086, 1081, 1082, 1080))
    On Error GoTo 0
End Function

Private Sub ClearBackendCellSafe(ByVal targetCell As Range)
    On Error Resume Next
    If targetCell.MergeCells Then
        targetCell.MergeArea.Cells(1, 1).ClearContents
    Else
        targetCell.ClearContents
    End If
    On Error GoTo 0
End Sub

Private Function EnrollmentDefaultCity() As String
    EnrollmentDefaultCity = mdlHelper.Ru(1043, 1088, 1086, 1079, 1085, 1099, 1081)
End Function

Private Function EnrollmentDefaultSignatoryName() As String
    EnrollmentDefaultSignatoryName = mdlHelper.Ru(1045) & "." & mdlHelper.Ru(1050, 1086, 1088, 1086, 1087, 1077, 1094)
End Function

Private Function EnrollmentDefaultSignatoryRank() As String
    EnrollmentDefaultSignatoryRank = mdlHelper.Ru(1084, 1072, 1081, 1086, 1088)
End Function

Private Function EnrollmentDefaultSignatoryPosition() As String
    EnrollmentDefaultSignatoryPosition = mdlHelper.Ru(1042, 1056, 1048, 1054) & " " & mdlHelper.Ru(1050, 1054, 1052, 1040, 1053, 1044, 1048, 1056, 1040) & " " & mdlHelper.Ru(1042, 1054, 1049, 1057, 1050, 1054, 1049) & " " & mdlHelper.Ru(1063, 1040, 1057, 1058, 1048) & " 81510"
End Function

Private Function EnrollmentDefaultHeaderText() As String
    EnrollmentDefaultHeaderText = mdlHelper.Ru(1055, 1056, 1054, 1045, 1050, 1058) & " " & mdlHelper.Ru(1055, 1056, 1048, 1050, 1040, 1047, 1040) & "|" & _
                                  mdlHelper.Ru(1050, 1054, 1052, 1040, 1053, 1044, 1048, 1056, 1040) & " " & mdlHelper.Ru(1042, 1054, 1049, 1057, 1050, 1054, 1049) & " " & mdlHelper.Ru(1063, 1040, 1057, 1058, 1048) & " {unit}|(" & _
                                  mdlHelper.Ru(1087, 1086, 32, 1089, 1090, 1088, 1086, 1077, 1074, 1086, 1081, 32, 1095, 1072, 1089, 1090, 1080) & ")"
End Function

Private Function EnrollmentDefaultFilenameTemplate() As String
    EnrollmentDefaultFilenameTemplate = mdlHelper.Ru(1055, 1088, 1080, 1082, 1072, 1079) & "_" & mdlHelper.Ru(1086) & "_" & mdlHelper.Ru(1079, 1072, 1095, 1080, 1089, 1083, 1077, 1085, 1080, 1080) & "_{orderDraftId}_{date}"
End Function

Private Function EnrollmentDefaultTemplateFileName() As String
    EnrollmentDefaultTemplateFileName = mdlHelper.Ru(1064, 1072, 1073, 1083, 1086, 1085, 95, 1047, 1072, 1095, 1080, 1089, 1083, 1077, 1085, 1080, 1077, 46, 100, 111, 99, 120)
End Function

Private Sub EnsureEnrollmentPaymentDefinitions(ByVal ws As Worksheet)
    EnsureEnrollmentPaymentDefinition ws, "core", "core", "Section1Core", "core", "", "manual", "blocked", ET("enrollment.word.block.core", "Основное кадровое действие")
    EnsureEnrollmentPaymentDefinition ws, "std_duty", "standard", "Section1MonthlyStandard", "std_duty", "basis", "standard_start_date", "warning", ET("payments.type.std_duty", "Надбавка по воинской должности")
    EnsureEnrollmentPaymentDefinition ws, "std_special", "standard", "Section1MonthlyStandard", "std_special", "basis", "standard_start_date", "warning", ET("payments.type.std_special", "Особые условия службы")
    EnsureEnrollmentPaymentDefinition ws, "std_tariff", "standard", "Section1MonthlyStandard", "std_tariff", "basis", "standard_start_date", "warning", ET("payments.type.std_tariff", "Надбавка 1-4 тарифных разрядов")
    EnsureEnrollmentPaymentDefinition ws, "std_contract430", "standard", "Section1MonthlyStandard", "std_contract430", "basis", "standard_start_date", "warning", ET("payments.type.std_contract430", "Надбавка 430 приказ")
    EnsureEnrollmentPaymentDefinition ws, "class", "personal", "Section1MonthlyPersonal", "class", "param,basis", "standard_start_date", "blocked", ET("payments.type.class_qualification", "Классная квалификация")
    EnsureEnrollmentPaymentDefinition ws, "fizo", "personal", "Section1MonthlyPersonal", "fizo", "param,basis", "standard_start_date", "blocked", ET("payments.type.fizo", "ФИЗО")
    EnsureEnrollmentPaymentDefinition ws, "secrecy", "personal", "Section1MonthlyPersonal", "secrecy", "param,basis", "standard_start_date", "blocked", ET("payments.type.secrecy", "Секретность")
    EnsureEnrollmentPaymentDefinition ws, "achievement", "personal", "Section1MonthlyPersonal", "achievement", "param,basis", "standard_start_date", "blocked", ET("payments.type.achievement", "Особые достижения")
    EnsureEnrollmentPaymentDefinition ws, "lift", "onetime", "Section1OneTime", "lift", "amount,date,basis", "enroll_date", "warning", ET("enrollment.field.lift_enabled", "Подъёмное пособие")
    EnsureEnrollmentPaymentDefinition ws, "per_diem", "onetime", "Section1OneTime", "per_diem", "amount,date,basis", "enroll_date", "warning", ET("enrollment.field.per_diem_enabled", "Суточные")
    EnsureEnrollmentPaymentDefinition ws, "edv", "onetime", "Section2Edv400k", "edv", "amount,date,basis_section2,contract_basis", "enroll_date", "blocked", ET("enrollment.field.edv_enabled", "ЕДВ 400000")
    EnsureEnrollmentPaymentDefinition ws, "premium", "premium", "Section1Premium", "premium", "premium_end,basis", "enroll_date", "blocked", ET("enrollment.field.premium", "Премия")
    EnsureEnrollmentPaymentDefinition ws, "requisites", "requisites", "Section1Requisites", "requisites", "", "manual", "blocked", ET("enrollment.word.block.requisites", "Реквизиты и основания")
    EnsureEnrollmentPaymentDefinition ws, "extra_monthly", "personal", "Section1MonthlyPersonal", "extra_monthly", "name,amount,start,basis", "manual", "blocked", ET("enrollment.page.extras", "Дополнительные выплаты")
    EnsureEnrollmentPaymentDefinition ws, "extra_onetime", "onetime", "Section1OneTime", "extra_one_time", "name,amount,date,basis", "manual", "blocked", ET("enrollment.page.extras", "Дополнительные выплаты")
    UpgradeSettingValueIfEquals ws, "enrollment.def.core.text_template", ET("enrollment.word.block.core", "Основное кадровое действие"), "{core_text}"
    UpgradeSettingValueIfEquals ws, "enrollment.def.requisites.text_template", ET("enrollment.word.block.requisites", "Реквизиты и основания"), "{requisites_text}"
    UpgradeSettingValueIfEquals ws, "enrollment.def.edv.required_docs", "amount,date,basis_section2", "amount,date,basis_section2,contract_basis"
End Sub

Private Sub EnsureEnrollmentPaymentDefinition(ByVal ws As Worksheet, ByVal defCode As String, ByVal paymentKind As String, ByVal wordBlockTarget As String, ByVal journalBinding As String, ByVal requiredDocs As String, ByVal startDateSource As String, ByVal validationSeverity As String, ByVal labelText As String)
    EnsureSetting ws, "enrollment.def." & defCode & ".payment_kind", paymentKind, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".word_block_target", wordBlockTarget, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".journal_binding", journalBinding, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".required_docs", requiredDocs, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".start_date_source", startDateSource, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".validation_severity", validationSeverity, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".label", labelText, "Enrollment payment definition"
    EnsureSetting ws, "enrollment.def." & defCode & ".text_template", DefaultEnrollmentDefinitionTextTemplate(defCode), "Enrollment payment text template"
End Sub

Public Function GetEnrollmentPaymentDefinition(ByVal defCode As String) As Object
    Dim definition As Object

    Set definition = CreateObject("Scripting.Dictionary")
    definition.CompareMode = vbTextCompare
    definition("code") = defCode
    definition("payment_kind") = GetEnrollmentSetting("enrollment.def." & defCode & ".payment_kind", "")
    definition("word_block_target") = GetEnrollmentSetting("enrollment.def." & defCode & ".word_block_target", "")
    definition("journal_binding") = GetEnrollmentSetting("enrollment.def." & defCode & ".journal_binding", defCode)
    definition("required_docs") = GetEnrollmentSetting("enrollment.def." & defCode & ".required_docs", "")
    definition("start_date_source") = GetEnrollmentSetting("enrollment.def." & defCode & ".start_date_source", "")
    definition("validation_severity") = GetEnrollmentSetting("enrollment.def." & defCode & ".validation_severity", "")
    definition("label") = GetEnrollmentSetting("enrollment.def." & defCode & ".label", defCode)
    definition("text_template") = GetEnrollmentSetting("enrollment.def." & defCode & ".text_template", DefaultEnrollmentDefinitionTextTemplate(defCode))

    Set GetEnrollmentPaymentDefinition = definition
End Function

Public Function GetEnrollmentPaymentDefinitionsByBlock(ByVal wordBlockTarget As String) As Collection
    Dim code As Variant
    Dim definition As Object

    Set GetEnrollmentPaymentDefinitionsByBlock = New Collection
    For Each code In EnrollmentDefinitionCodes()
        Set definition = GetEnrollmentPaymentDefinition(CStr(code))
        If StrComp(SafeText(definition("word_block_target")), wordBlockTarget, vbTextCompare) = 0 Then
            GetEnrollmentPaymentDefinitionsByBlock.Add definition
        End If
    Next code
End Function

Private Function EnrollmentDefinitionCodes() As Variant
    EnrollmentDefinitionCodes = Array("core", "std_duty", "std_special", "std_tariff", "std_contract430", "class", "fizo", "secrecy", "achievement", "lift", "per_diem", "edv", "premium", "requisites", "extra_monthly", "extra_onetime")
End Function

Private Function DefaultEnrollmentDefinitionTextTemplate(ByVal defCode As String) As String
    Select Case LCase$(Trim$(defCode))
        Case "core"
            DefaultEnrollmentDefinitionTextTemplate = "{core_text}"
        Case "std_duty"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.monthly.std_duty", "Установить ежемесячную надбавку к денежному довольствию в размере {percent}% оклада по воинской должности.")
        Case "std_special"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.monthly.std_special", "Установить ежемесячную надбавку за особые условия военной службы в размере {percent}% оклада по воинской должности.")
        Case "std_tariff"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.monthly.std_tariff", "Установить ежемесячную надбавку по должностям 1-4 тарифных разрядов в размере {percent}% оклада по воинской должности.")
        Case "std_contract430"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.monthly.std_contract430", "Установить ежемесячную надбавку за особые достижения в службе по 430 ДСП в размере {percent}%.")
        Case "class"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.personal.class", "Установить ежемесячную надбавку за классную квалификацию") & " ({param}) " & ET("enrollment.word.personal.amount_prefix", "в размере") & " {amount_with_unit}."
        Case "fizo"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.personal.fizo", "Установить ежемесячную надбавку за уровень физической подготовленности") & " ({param}) " & ET("enrollment.word.personal.amount_prefix", "в размере") & " {amount_with_unit}."
        Case "secrecy"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.personal.secrecy", "Установить ежемесячную надбавку за работу со сведениями, составляющими государственную тайну") & " ({param}) " & ET("enrollment.word.personal.amount_prefix", "в размере") & " {amount_with_unit}."
        Case "achievement"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.personal.achievement", "Установить ежемесячную надбавку за особые достижения в службе / медаль") & " ({param}) " & ET("enrollment.word.personal.amount_prefix", "в размере") & " {amount_with_unit}."
        Case "lift"
            DefaultEnrollmentDefinitionTextTemplate = "    " & ET("enrollment.word.onetime.lift", "подъёмное пособие в размере") & " {amount}."
        Case "per_diem"
            DefaultEnrollmentDefinitionTextTemplate = "    " & ET("enrollment.word.onetime.per_diem", "суточные в размере") & " {amount} " & ET("enrollment.word.onetime.per_diem_days_prefix", "за") & " {days} " & ET("enrollment.word.onetime.per_diem_days_suffix", "сут.") & "."
        Case "edv"
            DefaultEnrollmentDefinitionTextTemplate = "{rank} {fio}, {personal_number}, {position}, " & ET("enrollment.word.section2.edv_text", "выплатить единовременную денежную выплату в размере") & " {amount} " & ET("enrollment.word.section2.currency", "рублей.")
        Case "premium"
            DefaultEnrollmentDefinitionTextTemplate = ET("enrollment.word.section1.premium_text", "Достоин выплаты ежемесячной премии в размере") & " {percent}% " & ET("enrollment.word.section1.premium_from", "с") & " {start} " & ET("enrollment.word.section1.premium_to", "по") & " {end}."
        Case "requisites"
            DefaultEnrollmentDefinitionTextTemplate = "{requisites_text}"
    End Select
End Function

Private Sub EnsureSetting(ByVal ws As Worksheet, ByVal settingKey As String, ByVal defaultValue As String, ByVal description As String)
    Dim rowNum As Long
    rowNum = FindSettingRow(ws, settingKey)
    If rowNum = 0 Then
        rowNum = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If rowNum < 2 Then rowNum = 2
        ws.Cells(rowNum, 1).Value = settingKey
        ws.Cells(rowNum, 2).Value = defaultValue
        ws.Cells(rowNum, 3).Value = description
    End If
End Sub

Private Sub UpgradeSettingValueIfEquals(ByVal ws As Worksheet, ByVal settingKey As String, ByVal legacyValue As String, ByVal upgradedValue As String)
    Dim rowNum As Long

    rowNum = FindSettingRow(ws, settingKey)
    If rowNum = 0 Then Exit Sub
    If StrComp(SafeText(ws.Cells(rowNum, 2).Value), legacyValue, vbTextCompare) = 0 Then
        ws.Cells(rowNum, 2).Value = upgradedValue
    End If
End Sub

Private Function FindSettingRow(ByVal ws As Worksheet, ByVal settingKey As String) As Long
    Dim lastRow As Long
    Dim rowNum As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For rowNum = 1 To lastRow
        If StrComp(SafeText(ws.Cells(rowNum, 1).Value), settingKey, vbTextCompare) = 0 Then
            FindSettingRow = rowNum
            Exit Function
        End If
    Next rowNum
End Function

Private Function SafeText(ByVal valueText As Variant) As String
    If IsError(valueText) Then Exit Function
    If IsNull(valueText) Then Exit Function
    SafeText = Trim$(CStr(valueText))
End Function

Private Function ET(ByVal key As String, ByVal fallback As String) As String
    ET = t(key, fallback)
End Function
