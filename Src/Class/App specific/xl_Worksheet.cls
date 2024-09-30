VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_Worksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private WS As Worksheet
Private ErrorCatalog(1, 99) As Variant
Private Initialized As Boolean
Private Handler As New std_Error


    Public Function AssignHandler(Optional n_ShowError As Boolean = True, Optional n_LogError As Boolean = True, Optional n_LoggingDestination As Variant, Optional n_ShowDestination As Variant) As Boolean
        On Error GoTo Error
        Let Handler.ShowError = n_ShowError
        Let Handler.LogError = n_LogError
        Set Handler.ShowDestination = n_ShowDestination
        Set Handler.LoggingDestination = n_LoggingDestination
        AssignHandler = True
        Exit Function
        Error:
    End Function

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = Handler.IS_ERROR
    End Property

    Public Property Set Worksheet(n_Worksheet As Worksheet)
        Set WS = n_Worksheet
    End Property

    Public Property Get Worksheet() As Worksheet
        Set Worksheet = WS
    End Property
    
    Public Function Activate(Optional WorkbookName As String = "", Optional WorksheetName As String = "") As Boolean

        Dim Temp As Worksheet
        Dim ErrorValue As String
        On Error GoTo Error
        If WorksheetName = "" Then
            Set Temp = WS
            ErrorValue = WS.Name
        Else
            If WorkbookName = "" Then
                Set Temp = Exists(WorksheetName, , True, True, False)
                ErrorValue = WorksheetName
            Else
                Set Temp = Exists(WorksheetName, WorkbookName, False, True, False)
                ErrorValue = WorkbookName & " | " & WorksheetName
            End If
        End If
        Temp.Activate
        Exit Function

        Error:
        Activate = Handler.Handle(ErrorCatalog, 1, WorksheetName)

    End Function

    Public Function Exists(Name As String, Optional Workbook As String = "", Optional SearchAllWorkbooks As Boolean = False, Optional ShouldExist As Boolean = True, Optional ReturnAsBoolean As Boolean = True) As Variant
        Dim WB As Workbook
        Dim WSs As Worksheet
        Dim Found As Boolean
        If SearchAllWorkbooks Then
            For Each WB In Workbooks
                For Each WSs In Worksheets
                    If WB.WSs.Name = Name Then Found = True: GoTo Cont
                Next WSs
            Next WB
        Else
            For Each WSs In Workbooks(Workbook).Worksheets
                If WSs.Name = Name Then Found = True: GoTo Cont
            Next WSs 
        End If
        Cont:
        If ReturnAsBoolean Then
            Exists = Not(Found Xor ShouldExist)
        Else
            If Found Then
                Set Exists = WSs
            Else
                Set Exists = Nothing
            End If
        End If
    End Function


    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
    
        If Initialized = False Then
        ' System-Errors
            ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "xl_Worksheet"
            ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Could not activate worksheet"
            ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "PLACEHOLDER"
            ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "PLACEHOLDER"
            ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "PLACEHOLDER"
            ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "PLACEHOLDER"
            ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "PLACEHOLDER"
            ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "PLACEHOLDER"
            ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "PLACEHOLDER"
            ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "PLACEHOLDER"
            ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "PLACEHOLDER"
            ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "PLACEHOLDER"
            ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "PLACEHOLDER"
            ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "PLACEHOLDER"
            ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "PLACEHOLDER"
            ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "PLACEHOLDER"
            ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "PLACEHOLDER"
            ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "PLACEHOLDER"
            ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "PLACEHOLDER"
            ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "PLACEHOLDER"
            ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "PLACEHOLDER"
            ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "PLACEHOLDER"
            ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "PLACEHOLDER"
            ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "PLACEHOLDER"
            ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "PLACEHOLDER"
            ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "PLACEHOLDER"
            ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "PLACEHOLDER"
            ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "PLACEHOLDER"
            ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "PLACEHOLDER"
            ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "PLACEHOLDER"
            ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "PLACEHOLDER"
            ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "PLACEHOLDER"
            ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
            ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
            ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
            ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "PLACEHOLDER"
            ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "PLACEHOLDER"
            ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "PLACEHOLDER"
            ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "PLACEHOLDER"
            ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "PLACEHOLDER"
            ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "PLACEHOLDER"
            ErrorCatalog(0, 0041) = 1000: ErrorCatalog(1, 0041) = "PLACEHOLDER"
            ErrorCatalog(0, 0042) = 1000: ErrorCatalog(1, 0042) = "PLACEHOLDER"
            ErrorCatalog(0, 0043) = 1000: ErrorCatalog(1, 0043) = "PLACEHOLDER"
            ErrorCatalog(0, 0044) = 1000: ErrorCatalog(1, 0044) = "PLACEHOLDER"
            ErrorCatalog(0, 0045) = 1000: ErrorCatalog(1, 0045) = "PLACEHOLDER"
            ErrorCatalog(0, 0046) = 1000: ErrorCatalog(1, 0046) = "PLACEHOLDER"
            ErrorCatalog(0, 0047) = 1000: ErrorCatalog(1, 0047) = "PLACEHOLDER"
            ErrorCatalog(0, 0048) = 1000: ErrorCatalog(1, 0048) = "PLACEHOLDER"
            ErrorCatalog(0, 0049) = 1000: ErrorCatalog(1, 0049) = "PLACEHOLDER"
            ErrorCatalog(0, 0050) = 1000: ErrorCatalog(1, 0050) = "PLACEHOLDER"
            ErrorCatalog(0, 0051) = 1000: ErrorCatalog(1, 0051) = "PLACEHOLDER"
            ErrorCatalog(0, 0052) = 1000: ErrorCatalog(1, 0052) = "PLACEHOLDER"
            ErrorCatalog(0, 0053) = 1000: ErrorCatalog(1, 0053) = "PLACEHOLDER"
            ErrorCatalog(0, 0054) = 1000: ErrorCatalog(1, 0054) = "PLACEHOLDER"
            ErrorCatalog(0, 0055) = 1000: ErrorCatalog(1, 0055) = "PLACEHOLDER"
            ErrorCatalog(0, 0056) = 1000: ErrorCatalog(1, 0056) = "PLACEHOLDER"
            ErrorCatalog(0, 0057) = 1000: ErrorCatalog(1, 0057) = "PLACEHOLDER"
            ErrorCatalog(0, 0058) = 1000: ErrorCatalog(1, 0058) = "PLACEHOLDER"
            ErrorCatalog(0, 0059) = 1000: ErrorCatalog(1, 0059) = "PLACEHOLDER"
            ErrorCatalog(0, 0060) = 1000: ErrorCatalog(1, 0060) = "PLACEHOLDER"
            ErrorCatalog(0, 0061) = 1000: ErrorCatalog(1, 0061) = "PLACEHOLDER"
            ErrorCatalog(0, 0062) = 1000: ErrorCatalog(1, 0062) = "PLACEHOLDER"
            ErrorCatalog(0, 0063) = 1000: ErrorCatalog(1, 0063) = "PLACEHOLDER"
            ErrorCatalog(0, 0064) = 1000: ErrorCatalog(1, 0064) = "PLACEHOLDER"
            ErrorCatalog(0, 0065) = 1000: ErrorCatalog(1, 0065) = "PLACEHOLDER"
            ErrorCatalog(0, 0066) = 1000: ErrorCatalog(1, 0066) = "PLACEHOLDER"
            ErrorCatalog(0, 0067) = 1000: ErrorCatalog(1, 0067) = "PLACEHOLDER"
            ErrorCatalog(0, 0068) = 1000: ErrorCatalog(1, 0068) = "PLACEHOLDER"
            ErrorCatalog(0, 0069) = 1000: ErrorCatalog(1, 0069) = "PLACEHOLDER"
            ErrorCatalog(0, 0070) = 1000: ErrorCatalog(1, 0070) = "PLACEHOLDER"
            ErrorCatalog(0, 0071) = 1000: ErrorCatalog(1, 0071) = "PLACEHOLDER"
            ErrorCatalog(0, 0072) = 1000: ErrorCatalog(1, 0072) = "PLACEHOLDER"
            ErrorCatalog(0, 0073) = 1000: ErrorCatalog(1, 0073) = "PLACEHOLDER"
            ErrorCatalog(0, 0074) = 1000: ErrorCatalog(1, 0074) = "PLACEHOLDER"
            ErrorCatalog(0, 0075) = 1000: ErrorCatalog(1, 0075) = "PLACEHOLDER"
            ErrorCatalog(0, 0076) = 1000: ErrorCatalog(1, 0076) = "PLACEHOLDER"
            ErrorCatalog(0, 0077) = 1000: ErrorCatalog(1, 0077) = "PLACEHOLDER"
            ErrorCatalog(0, 0078) = 1000: ErrorCatalog(1, 0078) = "PLACEHOLDER"
            ErrorCatalog(0, 0079) = 1000: ErrorCatalog(1, 0079) = "PLACEHOLDER"
            ErrorCatalog(0, 0080) = 1000: ErrorCatalog(1, 0080) = "PLACEHOLDER"
            ErrorCatalog(0, 0081) = 1000: ErrorCatalog(1, 0081) = "PLACEHOLDER"
            ErrorCatalog(0, 0082) = 1000: ErrorCatalog(1, 0082) = "PLACEHOLDER"
            ErrorCatalog(0, 0083) = 1000: ErrorCatalog(1, 0083) = "PLACEHOLDER"
            ErrorCatalog(0, 0084) = 1000: ErrorCatalog(1, 0084) = "PLACEHOLDER"
            ErrorCatalog(0, 0085) = 1000: ErrorCatalog(1, 0085) = "PLACEHOLDER"
            ErrorCatalog(0, 0086) = 1000: ErrorCatalog(1, 0086) = "PLACEHOLDER"
            ErrorCatalog(0, 0087) = 1000: ErrorCatalog(1, 0087) = "PLACEHOLDER"
            ErrorCatalog(0, 0088) = 1000: ErrorCatalog(1, 0088) = "PLACEHOLDER"
            ErrorCatalog(0, 0089) = 1000: ErrorCatalog(1, 0089) = "PLACEHOLDER"
            ErrorCatalog(0, 0090) = 1000: ErrorCatalog(1, 0090) = "PLACEHOLDER"
            ErrorCatalog(0, 0091) = 1000: ErrorCatalog(1, 0091) = "PLACEHOLDER"
            ErrorCatalog(0, 0092) = 1000: ErrorCatalog(1, 0092) = "PLACEHOLDER"
            ErrorCatalog(0, 0093) = 1000: ErrorCatalog(1, 0093) = "PLACEHOLDER"
            ErrorCatalog(0, 0094) = 1000: ErrorCatalog(1, 0094) = "PLACEHOLDER"
            ErrorCatalog(0, 0095) = 1000: ErrorCatalog(1, 0095) = "PLACEHOLDER"
            ErrorCatalog(0, 0096) = 1000: ErrorCatalog(1, 0096) = "PLACEHOLDER"
            ErrorCatalog(0, 0097) = 1000: ErrorCatalog(1, 0097) = "PLACEHOLDER"
            ErrorCatalog(0, 0098) = 1000: ErrorCatalog(1, 0098) = "PLACEHOLDER"
            ErrorCatalog(0, 0099) = 1000: ErrorCatalog(1, 0099) = "PLACEHOLDER"
            Initialized = True
        End If
    
    End Sub
'