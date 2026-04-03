Option Explicit

Public Enum KATSArCategory
    katsArvode = 1
    katsArvodeHelg = 2
    katsTidsspillan = 3
    katsTidsspillanOvrigTid = 4
End Enum

Public Enum KATSSectionKind
    katsSectionUnknown = 0
    katsSectionExpenses = 1
    katsSectionExpensesNoVat = 2
    katsSectionWorklog = 3
    katsSectionArvode = 4
    katsSectionArvodeTotal = 5
    katsSectionRecipient = 6
    katsSectionSignature = 7
End Enum

Public Type KATSSectionRef
    Kind As KATSSectionKind
    TagName As String
    Exists As Boolean
    HasTable As Boolean
End Type

Public Type KATSExpenseRowInput
    RowIndex As Long
    Desc As String
    Qty As Currency
    Rate As Currency
    ExistingAmount As Currency
End Type

Public Type KATSExpenseBlockInput
    Section As KATSSectionRef
    Heading As String
    SummaryRow As Long
    ApplyMileageRule As Boolean
    Count As Long
    Rows(1 To 200) As KATSExpenseRowInput
End Type

Public Type KATSExpenseBlockOutput
    Total As Currency
    Count As Long
    RowIndex(1 To 200) As Long
    Amount(1 To 200) As Currency
End Type

Public Type KATSWorklogInput
    Section As KATSSectionRef
    SummaryRow(1 To 4) As Long
    CategoryHours(1 To 4) As Currency
    CategoryAdjustment(1 To 4) As Currency
    HearingFound As Boolean
    HearingRow As Long
    HearingStart As Date
    HearingMinutes As Long
    IsTaxa As Boolean
End Type

Public Type KATSArvodeInput
    Section As KATSSectionRef
    RateText(1 To 6) As String
End Type

Public Type KATSRecipientInput
    Section As KATSSectionRef
    RawText As String
End Type

Public Type KATSSignatureInput
    Section As KATSSectionRef
End Type

Public Type KATSDocumentModel
    Worklog As KATSWorklogInput
    Expenses As KATSExpenseBlockInput
    ExpensesNoVat As KATSExpenseBlockInput
    Arvode As KATSArvodeInput
    Recipient As KATSRecipientInput
    Signature As KATSSignatureInput
End Type

Public Type KATSComputedState
    CategoryHours(1 To 4) As Currency
    CategoryAdjustment(1 To 4) As Currency

    HearingFound As Boolean
    HearingStart As Date
    HearingMinutes As Long
    IsTaxa As Boolean

    ExpenseOutput As KATSExpenseBlockOutput
    ExpenseNoVatOutput As KATSExpenseBlockOutput
    ExpenseExVatTotal As Currency
    ExpenseNoVatTotal As Currency

    ArvodeSpec(1 To 6) As String
    ArvodeAmount(1 To 6) As Currency
    ArvodeKeepRow(1 To 6) As Boolean
    ArvodeExMoms As Currency

    Postort As String
End Type
