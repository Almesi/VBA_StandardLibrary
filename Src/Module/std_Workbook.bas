Attribute VB_Name = "std_Workbook"

Option Explicit

Private ErrorCatalog(1, 99) As Variant
Private Initialized As Boolean

    ' Handles possible Errors when working with workbooks
    Public Function Workbook_Error(Name As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WB As WorkBook
        ProtInit
        If Name = Empty Then Workbook_Error = std_Error.Handle(ErrorCatalog, 01, "Name"): Exit Function
        If ShouldExist = True Then
            For Each WB in WorkBooks
                If WB.Name = Name Then Exit Function
            Next
            Workbook_Error = std_Error.Handle(ErrorCatalog, 02, Name)
        Else
            For Each WB in WorkBooks
                If WB.Name = Name Then Workbook_Error = std_Error.Handle(ErrorCatalog, 03, Name): Exit Function
            Next
        End If

    End Function

    ' Handles possible Errors when working with workheets
    Public Function Worksheet_Error(WorkbookName As String, SheetName As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WS As Worksheet
        ProtInit
        If Workbook_Error(WorkbookName, True) = std_Error.IS_ERROR Then Worksheet_Error = std_Error.IS_ERROR: Exit Function
        If SheetName = Empty Then Worksheet_Error = std_Error.Handle(ErrorCatalog, 01, "SheetName"   ):  Exit Function
        With Workbooks(WorkbookName)
            If ShouldExist = True Then
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Exit Function
                Next
                Worksheet_Error = std_Error.Handle(ErrorCatalog, 04, WorkbookName, SheetName)
            Else
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Worksheet_Error = std_Error.Handle(ErrorCatalog, 05, WorkbookName, SheetName): Exit Function
                Next
            End If
        End With

    End Function

    Public Function Workbook_Open(o_FilePath As Variant, Optional o_UpdateLinks As Variant = Empty, Optional o_ReadOnly As Variant = Empty, Optional o_Format As Variant = Empty, Optional o_Password As Variant = Empty, Optional o_WriteResPassword As Variant = Empty, Optional o_IgnoreReadOnlyRecommende As Variant = Empty, Optional o_Origin As Variant = Empty, Optional o_Delimiter As Variant = Empty, Optional o_Editable As Variant = Empty, Optional o_Notify As Variant = Empty, Optional o_Converter As Variant = Empty, Optional o_AddToMru As Variant = Empty, Optional o_Local As Variant = Empty, Optional o_CorruptLoad As Variant = Empty) As Boolean
        On Error GoTo Error
        If Workbook_Error(FilePath, True) <> std_Error.IS_ERROR Then
            Workbooks.Open o_FilePath, o_UpdateLinks, o_ReadOnly, o_Format, o_Password, o_WriteResPassword, o_IgnoreReadOnlyRecommende, o_Origin, o_Delimiter, o_Editable, o_Notify, o_Converter, o_AddToMru, o_Local, o_CorruptLoad
        End If
        Exit Function

        Error:
        Workbook_Open = std_Error.Handle(ErrorCatalog, 6, o_FilePath, o_UpdateLinks, o_ReadOnly, o_Format, o_Password, o_WriteResPassword, o_IgnoreReadOnlyRecommende, o_Origin, o_Delimiter, o_Editable, o_Notify, o_Converter, o_AddToMru, o_Local, o_CorruptLoad)
    End Function

    Public Function Worksheet_Activate(WorkbookName As String, SheetName As String) As Boolean

        On Error GoTo Error
        If Worksheet_Error(WorkbookName, SheetName, True) <> std_Error.IS_ERROR Then
            Workbooks(WorkbookName).Sheets(SheetName).Activate
        End If
        Exit Function

        Error:
        Worksheet_Activate = std_Error.Handle(ErrorCatalog, 7, Name)

    End Function

    Public Function Workbook_Close(Name As String) As Boolean

        Dim WB As Workbook
        Dim Found As Boolean

        On Error GoTo Error
        For Each WB In Workbooks
            If WB.Name = Name Then
                Found = True
                Exit For
            End If
        Next
        If Found = True Then WB.Close
        Exit Function

        Error:
        Workbook_Close = std_Error.Handle(ErrorCatalog, 8, Name)

    End Function
        
    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
    
        If Initialized = False Then
        ' System-Errors
            ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Workbook"
            ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Variable Empty"
            ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Workbook doesnt Exist"
            ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Workbook Exists"
            ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Worksheet doesnt Exist"
            ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "Worksheet exists"
            ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "could not open workbook"
            ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "could not activate worksheet"
            ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "could not close Workbook"
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