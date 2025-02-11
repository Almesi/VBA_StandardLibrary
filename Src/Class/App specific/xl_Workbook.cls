VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_Workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private WB As Workbook
Private ErrorCatalog(1, 99) As Variant
Private p_Handler As New std_Error

    Public Property Let Handler(n_Handler As std_Error)
        Set p_Handler = n_Handler
    End Property

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = p_Handler.IS_ERROR
    End Property

    Public Property Let Workbook(n_Workbook As Workbook)
        Set WB = n_Workbook
    End Property

    Public Property Get Workbook() As Workbook
        Set Workbook = WB
    End Property

    Public Property Let WorkbookViaName(Name As String)
        Dim Temp As Workbook
        Set Temp = Exists(Name, True, False)
        If Temp Is Nothing Then
        Else
            Set WB = Temp
        End If
    End Property

    Public Function Openn(FilePath As Variant, Optional UpdateLinks As Variant = Empty, Optional ReadOnly As Variant = Empty, Optional Format As Variant = Empty, Optional Password As Variant = Empty, Optional WriteResPassword As Variant = Empty, Optional IgnoreReadOnlyRecommende As Variant = Empty, Optional Origin As Variant = Empty, Optional Delimiter As Variant = Empty, Optional Editable As Variant = Empty, Optional Notify As Variant = Empty, Optional Converter As Variant = Empty, Optional AddToMru As Variant = Empty, Optional n_Local As Variant = Empty, Optional CorruptLoad As Variant = Empty, Optional ThrowError As Boolean = True) As Variant
        On Error GoTo Error
        Workbooks.Open FilePath
        Set WB = ActiveWorkbook
        Exit Function

        Error:
        Openn = p_Handler.Handle(ErrorCatalog, 1, ThrowError, FilePath)
    End Function

    Public Function Closee(Optional Name As String = "", Optional SaveChanges As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
        Dim Temp As Workbook
        Dim ErrorValue As String

        On Error GoTo Error
        If Name = "" Then
            ErrorValue = WB.Name
            WB.Close SaveChanges
        Else
            Set Temp = Exists(Name)
            If Temp Is Nothing Then
                ErrorValue = Name
                GoTo Error
            Else
                ErrorValue = Temp.Name
                Temp.Close
            End If 
        End If
        Exit Function

        Error:
        Closee = p_Handler.Handle(ErrorCatalog, 2, ThrowError, ErrorValue)
        
    End Function

    Public Function Exists(NameOrPath As String, Optional ShouldExist As Boolean = True, Optional ReturnAsBoolean As Boolean = True) As Variant
        Dim WBs As Workbook
        Dim Found As Boolean
        Dim FSO As Object

        Set FSO = CreateObject("Scripting.FileSystemObject")
        If FSO.FileExists(NameOrPath) Then Found = True: GoTo SkipWB

        For Each WBs In Workbooks
            If WBs.Name = NameOrPath Then Found = True: Exit For
        Next WBs
        SkipWB:
        If ReturnAsBoolean Then
            Exists = Not (Found Xor ShouldExist)
        Else
            If Found Then
                Set Exists = WBs
            Else
                Set Exists = Nothing
            End If
        End If
    End Function

    Public Function Add(FilePath As String, Optional Unique As Boolean = True, Optional ThrowError As Boolean = True) As Boolean
        Dim i As Long, j As Long
        Dim FolderSplit() As String
        Dim FileEnding As String
        On Error GoTo Error
        If Exists(FilePath) Then
            If Unique Then
                Add = p_Handler.Handle(ErrorCatalog, 3, ThrowError, FilePath)
            Else
                Set WB = Workbooks.Add
                FileEnding = Mid(FolderSplit(Ubound(FolderSplit)), 1, InStr(1, FolderSplit(Ubound(FolderSplit)), ".") - 1)
                FolderSplit = Split(Replace(FilePath, "/", "\"), "\")
                FolderSplit(Ubound(FolderSplit) - 1) = FolderSplit(Ubound(FolderSplit) - 1) & i
                Do Until Exists(FilePath)
                    i = i + 1
                    For j = 0 To Ubound(FolderSplit) - 2
                        FilePath = FilePath & FolderSplit(j)
                    Next j
                    FilePath = FilePath & Mid(FolderSplit(j), 1, Len(FolderSplit(j - 1))) & i
                    FilePath = FilePath & FolderSplit(j + 1)
                Loop
                WB.SaveAs(FilePath)
            End If
        Else
            Set WB = Workbooks.Add
            WB.SaveAs(FilePath)
        End If
        Exit Function

        Error:
        Add = p_Handler.Handle(ErrorCatalog, 4, ThrowError, FilePath)
    End Function

    Private Sub Class_Initialize()
        ProtInit
        Set p_Handler = std_Error
    End Sub

    Private Sub Class_Terminate()
    End Sub
        
    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "xl_Workbook"
        ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Could not open workbook"
        ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Could not close workbook"
        ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "File already exists"
        ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Couldnt add and save Workbook"
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
    End Sub
'