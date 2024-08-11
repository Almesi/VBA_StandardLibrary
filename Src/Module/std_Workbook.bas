Attribute VB_Name = "std_Workbook"

Option Explicit

Private WorkbookCatalog(1, 99) As Variant

Private Initialized As Boolean

    ' Handles possible Errors when working with workbooks
    Public Function Workbook_Error(WorkbookName As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WB As WorkBook
        ProtInit
        If WorkbookName = Empty Then Workbook_Error = std_Error.Handle(WorkbookCatalog, 01, "WorkbookName"): Exit Function
        If ShouldExist = True Then
            For Each WB in WorkBooks
                If WB.Name = WorkbookName Then Exit Function
            Next
            Workbook_Error = std_Error.Handle(WorkbookCatalog, 02, WorkbookName)
        Else
            For Each WB in WorkBooks
                If WB.Name = WorkbookName Then Workbook_Error = std_Error.Handle(WorkbookCatalog, 03, WorkbookName): Exit Function
            Next
        End If

    End Function

    ' Handles possible Errors when working with workheets
    Public Function Worksheet_Error(WorkbookName As String, SheetName As String, Optional ShouldExist As Boolean = True) As Boolean

        Dim WS As Worksheet
        ProtInit
        If Workbook_Error(WorkBookName, True) = IS_ERROR Then Worksheet_Error = IS_ERROR: Exit Function
        If SheetName = Empty Then Worksheet_Error = std_Error.Handle(WorkbookCatalog, 01, "SheetName"   ):  Exit Function
        With Workbooks(WorkbookName)
            If ShouldExist = True Then
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Exit Function
                Next
                Worksheet_Error = std_Error.Handle(WorkbookCatalog, 04, WorkbookName, SheetName)
            Else
                For Each WS in .Worksheets
                    If WS.Name = SheetName Then Worksheet_Error = std_Error.Handle(WorkbookCatalog, 05, WorkbookName, SheetName): Exit Function
                Next
            End If
        End With

    End Function




        
    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
    
        If Initialized = False Then
        ' System-Errors
            WorkbookCatalog(0, 0000) = 0002: WorkbookCatalog(1, 0000) = "std_Workbook"
            WorkbookCatalog(0, 0001) = 1000: WorkbookCatalog(1, 0001) = "Variable Empty"
            WorkbookCatalog(0, 0002) = 1000: WorkbookCatalog(1, 0002) = "Workbook doesnt Exist"
            WorkbookCatalog(0, 0003) = 1000: WorkbookCatalog(1, 0003) = "Workbook Exists"
            WorkbookCatalog(0, 0004) = 1000: WorkbookCatalog(1, 0004) = "Worksheet doesnt Exist"
            WorkbookCatalog(0, 0005) = 1000: WorkbookCatalog(1, 0005) = "Worksheet exists"
            WorkbookCatalog(0, 0006) = 1000: WorkbookCatalog(1, 0006) = "PLACEHOLDER"
            WorkbookCatalog(0, 0007) = 1000: WorkbookCatalog(1, 0007) = "PLACEHOLDER"
            WorkbookCatalog(0, 0008) = 1000: WorkbookCatalog(1, 0008) = "PLACEHOLDER"
            WorkbookCatalog(0, 0009) = 1000: WorkbookCatalog(1, 0009) = "PLACEHOLDER"
            WorkbookCatalog(0, 0010) = 1000: WorkbookCatalog(1, 0010) = "PLACEHOLDER"
            WorkbookCatalog(0, 0011) = 1000: WorkbookCatalog(1, 0011) = "PLACEHOLDER"
            WorkbookCatalog(0, 0012) = 1000: WorkbookCatalog(1, 0012) = "PLACEHOLDER"
            WorkbookCatalog(0, 0013) = 1000: WorkbookCatalog(1, 0013) = "PLACEHOLDER"
            WorkbookCatalog(0, 0014) = 1000: WorkbookCatalog(1, 0014) = "PLACEHOLDER"
            WorkbookCatalog(0, 0015) = 1000: WorkbookCatalog(1, 0015) = "PLACEHOLDER"
            WorkbookCatalog(0, 0016) = 1000: WorkbookCatalog(1, 0016) = "PLACEHOLDER"
            WorkbookCatalog(0, 0017) = 1000: WorkbookCatalog(1, 0017) = "PLACEHOLDER"
            WorkbookCatalog(0, 0018) = 1000: WorkbookCatalog(1, 0018) = "PLACEHOLDER"
            WorkbookCatalog(0, 0019) = 1000: WorkbookCatalog(1, 0019) = "PLACEHOLDER"
            WorkbookCatalog(0, 0020) = 1000: WorkbookCatalog(1, 0020) = "PLACEHOLDER"
            WorkbookCatalog(0, 0021) = 1000: WorkbookCatalog(1, 0021) = "PLACEHOLDER"
            WorkbookCatalog(0, 0022) = 1000: WorkbookCatalog(1, 0022) = "PLACEHOLDER"
            WorkbookCatalog(0, 0023) = 1000: WorkbookCatalog(1, 0023) = "PLACEHOLDER"
            WorkbookCatalog(0, 0024) = 1000: WorkbookCatalog(1, 0024) = "PLACEHOLDER"
            WorkbookCatalog(0, 0025) = 1000: WorkbookCatalog(1, 0025) = "PLACEHOLDER"
            WorkbookCatalog(0, 0026) = 1000: WorkbookCatalog(1, 0026) = "PLACEHOLDER"
            WorkbookCatalog(0, 0027) = 1000: WorkbookCatalog(1, 0027) = "PLACEHOLDER"
            WorkbookCatalog(0, 0028) = 1000: WorkbookCatalog(1, 0028) = "PLACEHOLDER"
            WorkbookCatalog(0, 0029) = 1000: WorkbookCatalog(1, 0029) = "PLACEHOLDER"
            WorkbookCatalog(0, 0030) = 1000: WorkbookCatalog(1, 0030) = "PLACEHOLDER"
            WorkbookCatalog(0, 0031) = 1000: WorkbookCatalog(1, 0031) = "PLACEHOLDER"
            WorkbookCatalog(0, 0032) = 1000: WorkbookCatalog(1, 0032) = "PLACEHOLDER"
            WorkbookCatalog(0, 0033) = 1000: WorkbookCatalog(1, 0033) = "PLACEHOLDER"
            WorkbookCatalog(0, 0034) = 1000: WorkbookCatalog(1, 0034) = "PLACEHOLDER"
            WorkbookCatalog(0, 0035) = 1000: WorkbookCatalog(1, 0035) = "PLACEHOLDER"
            WorkbookCatalog(0, 0036) = 1000: WorkbookCatalog(1, 0036) = "PLACEHOLDER"
            WorkbookCatalog(0, 0037) = 1000: WorkbookCatalog(1, 0037) = "PLACEHOLDER"
            WorkbookCatalog(0, 0038) = 1000: WorkbookCatalog(1, 0038) = "PLACEHOLDER"
            WorkbookCatalog(0, 0039) = 1000: WorkbookCatalog(1, 0039) = "PLACEHOLDER"
            WorkbookCatalog(0, 0040) = 1000: WorkbookCatalog(1, 0040) = "PLACEHOLDER"
            WorkbookCatalog(0, 0041) = 1000: WorkbookCatalog(1, 0041) = "PLACEHOLDER"
            WorkbookCatalog(0, 0042) = 1000: WorkbookCatalog(1, 0042) = "PLACEHOLDER"
            WorkbookCatalog(0, 0043) = 1000: WorkbookCatalog(1, 0043) = "PLACEHOLDER"
            WorkbookCatalog(0, 0044) = 1000: WorkbookCatalog(1, 0044) = "PLACEHOLDER"
            WorkbookCatalog(0, 0045) = 1000: WorkbookCatalog(1, 0045) = "PLACEHOLDER"
            WorkbookCatalog(0, 0046) = 1000: WorkbookCatalog(1, 0046) = "PLACEHOLDER"
            WorkbookCatalog(0, 0047) = 1000: WorkbookCatalog(1, 0047) = "PLACEHOLDER"
            WorkbookCatalog(0, 0048) = 1000: WorkbookCatalog(1, 0048) = "PLACEHOLDER"
            WorkbookCatalog(0, 0049) = 1000: WorkbookCatalog(1, 0049) = "PLACEHOLDER"
            WorkbookCatalog(0, 0050) = 1000: WorkbookCatalog(1, 0050) = "PLACEHOLDER"
            WorkbookCatalog(0, 0051) = 1000: WorkbookCatalog(1, 0051) = "PLACEHOLDER"
            WorkbookCatalog(0, 0052) = 1000: WorkbookCatalog(1, 0052) = "PLACEHOLDER"
            WorkbookCatalog(0, 0053) = 1000: WorkbookCatalog(1, 0053) = "PLACEHOLDER"
            WorkbookCatalog(0, 0054) = 1000: WorkbookCatalog(1, 0054) = "PLACEHOLDER"
            WorkbookCatalog(0, 0055) = 1000: WorkbookCatalog(1, 0055) = "PLACEHOLDER"
            WorkbookCatalog(0, 0056) = 1000: WorkbookCatalog(1, 0056) = "PLACEHOLDER"
            WorkbookCatalog(0, 0057) = 1000: WorkbookCatalog(1, 0057) = "PLACEHOLDER"
            WorkbookCatalog(0, 0058) = 1000: WorkbookCatalog(1, 0058) = "PLACEHOLDER"
            WorkbookCatalog(0, 0059) = 1000: WorkbookCatalog(1, 0059) = "PLACEHOLDER"
            WorkbookCatalog(0, 0060) = 1000: WorkbookCatalog(1, 0060) = "PLACEHOLDER"
            WorkbookCatalog(0, 0061) = 1000: WorkbookCatalog(1, 0061) = "PLACEHOLDER"
            WorkbookCatalog(0, 0062) = 1000: WorkbookCatalog(1, 0062) = "PLACEHOLDER"
            WorkbookCatalog(0, 0063) = 1000: WorkbookCatalog(1, 0063) = "PLACEHOLDER"
            WorkbookCatalog(0, 0064) = 1000: WorkbookCatalog(1, 0064) = "PLACEHOLDER"
            WorkbookCatalog(0, 0065) = 1000: WorkbookCatalog(1, 0065) = "PLACEHOLDER"
            WorkbookCatalog(0, 0066) = 1000: WorkbookCatalog(1, 0066) = "PLACEHOLDER"
            WorkbookCatalog(0, 0067) = 1000: WorkbookCatalog(1, 0067) = "PLACEHOLDER"
            WorkbookCatalog(0, 0068) = 1000: WorkbookCatalog(1, 0068) = "PLACEHOLDER"
            WorkbookCatalog(0, 0069) = 1000: WorkbookCatalog(1, 0069) = "PLACEHOLDER"
            WorkbookCatalog(0, 0070) = 1000: WorkbookCatalog(1, 0070) = "PLACEHOLDER"
            WorkbookCatalog(0, 0071) = 1000: WorkbookCatalog(1, 0071) = "PLACEHOLDER"
            WorkbookCatalog(0, 0072) = 1000: WorkbookCatalog(1, 0072) = "PLACEHOLDER"
            WorkbookCatalog(0, 0073) = 1000: WorkbookCatalog(1, 0073) = "PLACEHOLDER"
            WorkbookCatalog(0, 0074) = 1000: WorkbookCatalog(1, 0074) = "PLACEHOLDER"
            WorkbookCatalog(0, 0075) = 1000: WorkbookCatalog(1, 0075) = "PLACEHOLDER"
            WorkbookCatalog(0, 0076) = 1000: WorkbookCatalog(1, 0076) = "PLACEHOLDER"
            WorkbookCatalog(0, 0077) = 1000: WorkbookCatalog(1, 0077) = "PLACEHOLDER"
            WorkbookCatalog(0, 0078) = 1000: WorkbookCatalog(1, 0078) = "PLACEHOLDER"
            WorkbookCatalog(0, 0079) = 1000: WorkbookCatalog(1, 0079) = "PLACEHOLDER"
            WorkbookCatalog(0, 0080) = 1000: WorkbookCatalog(1, 0080) = "PLACEHOLDER"
            WorkbookCatalog(0, 0081) = 1000: WorkbookCatalog(1, 0081) = "PLACEHOLDER"
            WorkbookCatalog(0, 0082) = 1000: WorkbookCatalog(1, 0082) = "PLACEHOLDER"
            WorkbookCatalog(0, 0083) = 1000: WorkbookCatalog(1, 0083) = "PLACEHOLDER"
            WorkbookCatalog(0, 0084) = 1000: WorkbookCatalog(1, 0084) = "PLACEHOLDER"
            WorkbookCatalog(0, 0085) = 1000: WorkbookCatalog(1, 0085) = "PLACEHOLDER"
            WorkbookCatalog(0, 0086) = 1000: WorkbookCatalog(1, 0086) = "PLACEHOLDER"
            WorkbookCatalog(0, 0087) = 1000: WorkbookCatalog(1, 0087) = "PLACEHOLDER"
            WorkbookCatalog(0, 0088) = 1000: WorkbookCatalog(1, 0088) = "PLACEHOLDER"
            WorkbookCatalog(0, 0089) = 1000: WorkbookCatalog(1, 0089) = "PLACEHOLDER"
            WorkbookCatalog(0, 0090) = 1000: WorkbookCatalog(1, 0090) = "PLACEHOLDER"
            WorkbookCatalog(0, 0091) = 1000: WorkbookCatalog(1, 0091) = "PLACEHOLDER"
            WorkbookCatalog(0, 0092) = 1000: WorkbookCatalog(1, 0092) = "PLACEHOLDER"
            WorkbookCatalog(0, 0093) = 1000: WorkbookCatalog(1, 0093) = "PLACEHOLDER"
            WorkbookCatalog(0, 0094) = 1000: WorkbookCatalog(1, 0094) = "PLACEHOLDER"
            WorkbookCatalog(0, 0095) = 1000: WorkbookCatalog(1, 0095) = "PLACEHOLDER"
            WorkbookCatalog(0, 0096) = 1000: WorkbookCatalog(1, 0096) = "PLACEHOLDER"
            WorkbookCatalog(0, 0097) = 1000: WorkbookCatalog(1, 0097) = "PLACEHOLDER"
            WorkbookCatalog(0, 0098) = 1000: WorkbookCatalog(1, 0098) = "PLACEHOLDER"
            WorkbookCatalog(0, 0099) = 1000: WorkbookCatalog(1, 0099) = "PLACEHOLDER"
            Initialized = True
        End If
    
    End Sub
'