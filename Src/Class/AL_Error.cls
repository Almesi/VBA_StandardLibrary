VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AL_Error"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' This Module is used to handle Errors
' The main philosophy is to raise Errors and (maybe) stop code execution, depending on the severity of the Error
' The main use is for programmers to significantly reduce Errorhandling-code and to standardize it
    ' The programmer needs (at least) 2 Values to raise the Error: The Category of the Error and the index of the Error. The Severity and the Message will be pulled from AL_Error_Catalog
' Just raising the Error might be nice, but to abstract even higher there are more Errorfunctions
    ' For example Variable Compares Numbers or Text and raises the Error
        ' The resulting code for the programmer would then be reduced to just:

            ' If Variable(a, "=", b) = IS_ERROR Then
                ' Do Something

        ' instead of:

            ' If a = b Then
                ' Show Error
                ' Print Error
                ' Handle Error Severity
                ' Do Something







' Array of all Errormessages
    ' First Dimension is Category
    ' Second Dimension is Severity/Message
    ' Third Dimension is Index
Private AL_Error_Catalog(7, 1, 99) As Variant

' Used to determine, if it is an Error
Public IS_ERROR As Boolean

' Used once to initialize all Errormessages
Private Initialized As Boolean
Private ShowError As Boolean
Private PrintError As Boolean

' Stops process if this value is lower than the Error severity
Private Const AL_SEVERITY_BREAK As Long = 1000

' Runs a Question and handles Error accordingly. Should be below AL_SEVERITY_BREAK
Private Const AL_ERROR_QUESTION As Long = 1

' Standard Value if no ErrorValue is passed
Private Const AL_EMPTY_ERROR As Variant = Empty


' Handles Errormessage
Public Function Handle(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR) As Boolean

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant

    ErrorArray = Error_Get(ErrorCategory, ErrorIndex)
    If ErrorArray(0) = AL_ERROR_QUESTION Then
        Handle = Ask(ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4)
    Else
        Show  ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4
        Handle = IS_ERROR
    End If
    Printt ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4 
    If ErrorArray(0) >= AL_SEVERITY_BREAK Then
        End
    End If

End Function

' Compare Numbers or Text and Handle Errors
Public Function Variable(FirstValue As Variant, Optional Operator As String = Empty, Optional SecondValue As Variant = Empty, Optional MinValue As Variant = Empty, Optional MaxValue As Variant = Empty) As Boolean

    ProtInit
    If Operator <> Empty Then
        Select Case UCase(Operator)
            Case "=", "IS"
                If FirstValue =  SecondValue Then Variable = Handle(0, 15, FirstValue, SecondValue): Exit Function
            Case "<>", "NOT", "!="
                If FirstValue <> SecondValue Then Variable = Handle(0, 06, FirstValue, SecondValue): Exit Function
            Case "<"
                If FirstValue <  SecondValue Then Variable = Handle(0, 10, FirstValue, SecondValue): Exit Function
            Case ">"
                If FirstValue >  SecondValue Then Variable = Handle(0, 09, FirstValue, SecondValue): Exit Function
            Case "=<", "<="
                If FirstValue =< SecondValue Then Variable = Handle(0, 08, FirstValue, SecondValue): Exit Function
            Case ">=", "=>"
                If FirstValue >= SecondValue Then Variable = Handle(0, 07, FirstValue, SecondValue): Exit Function
            Case Else
        End Select
    End If
    If MinValue <> Empty Then
        If FirstValue < MinValue             Then Variable = Handle(0, 09, FirstValue, MinValue):    Exit Function
    End If
    If MaxValue <> Empty Then
        If FirstValue > MaxValue             Then Variable = Handle(0, 10, FirstValue, MaxValue):    Exit Function
    End If
    If FirstValue = Empty                    Then Variable = Handle(0, 02, "FirstValue"):              Exit Function

End Function

' Compare Numbers or Text and Handle Errors
Public Function Object(FirstValue As Object, Optional Operator As String = Empty, Optional SecondValue As Object = Empty) As Boolean

    ProtInit
    If Operator <> Empty Then
        Select Case UCase(Operator)
            Case "=", "IS"
                If FirstValue     IS SecondValue Then Object = Handle(0, 15, "FirstValue", "SecondValue"): Exit Function
            Case "<>", "NOT", "!="
                If Not FirstValue IS SecondValue Then Object = Handle(0, 06, "FirstValue", "SecondValue"): Exit Function
        End Select
    End If
    If FirstValue Is Nothing                      Then Object = Handle(0, 03, "FirstValue"):               Exit Function

End Function

' Handles possible Errors when working with worksheets
Public Function WorkBook(WorkbookName As String, ShouldExist As Boolean) As Boolean

    Dim WB As WorkBook
    ProtInit
    If WorkbookName = Empty Then WorkBook = Handle(0, 02, "WorkbookName"): Exit Function
    If ShouldExist = True Then
        For Each WB in WorkBooks
            If WB.Name = WorkbookName Then Exit Function
        Next
        WorkBook = Handle(3, 00, WorkbookName)
    Else
        For Each WB in WorkBooks
            If WB.Name = WorkbookName Then WorkBook = Handle(3, 01, WorkbookName): Exit Function
        Next
    End If

End Function

' Handles possible Errors when working with worksheets
Public Function WorkSheet(WorkbookName As String, SheetName As String, ShouldExist As Boolean) As Boolean

    Dim WS As Worksheet
    ProtInit
    If Workbook(WorkBookName, True) = IS_ERROR Then WorkSheet = IS_ERROR: Exit Function
    If SheetName = Empty Then WorkSheet = Handle(0, 02, "SheetName"   ):  Exit Function
    With Workbooks(WorkbookName)
        If ShouldExist = True Then
            For Each WS in .Worksheets
                If WS.Name = SheetName Then Exit Function
            Next
            WorkSheet = Handle(4, 01, WorkbookName, SheetName)
        Else
            For Each WS in .Worksheets
                If WS.Name = SheetName Then WorkSheet = Handle(4, 00, WorkbookName, SheetName): Exit Function
            Next
        End If
    End With

End Function

' Print Error to Immediate
Private Sub Printt(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR)

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant

    ProtInit
    If PrintError = False Then Exit Sub
    ErrorArray = Error_Get(ErrorCategory, ErrorIndex)
    ErrorMessage = "( " & "Category: " & Error_GetCategory(ErrorCategory)    & " ) | " & Chr(13) & _
                   "( " & "Severity: " & ErrorArray(0)                       & " ) | " & Chr(13) & _
                   "( " & "Index: "    & ErrorArray(1)                       & " ) | " & Chr(13) & _
                   "( " & "Message: "  & ErrorArray(2)                       & " ) | " & Chr(13)
    If ErrorValue1 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value1: " & ErrorValue1 & " ) | " & Chr(13)
    If ErrorValue2 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value2: " & ErrorValue2 & " ) | " & Chr(13)
    If ErrorValue3 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value3: " & ErrorValue3 & " ) | " & Chr(13)
    If ErrorValue4 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value4: " & ErrorValue4 & " ) | " & Chr(13)
    ErrorMessage = ErrorMessage & "------------------------------------------------------------------------------"
    Debug.Print ErrorMessage

End Sub

' Print Error as MessageBox
Private Sub Show(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR)

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant
    ' For the msgbox to run, has no other use
    Dim Temp As Variant

    ProtInit
    If ShowError = False Then Exit Sub
    ErrorArray = Error_Get(ErrorCategory, ErrorIndex)
    ErrorMessage = "( " & "Category: " & Error_GetCategory(ErrorCategory)    & " ) | " & Chr(13) & _
                   "( " & "Severity: " & ErrorArray(0)                       & " ) | " & Chr(13) & _
                   "( " & "Index: "    & ErrorArray(1)                       & " ) | " & Chr(13) & _
                   "( " & "Message: "  & ErrorArray(2)                       & " ) | " & Chr(13)
    If ErrorValue1 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value1: " & ErrorValue1 & " ) | " & Chr(13)
    If ErrorValue2 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value2: " & ErrorValue2 & " ) | " & Chr(13)
    If ErrorValue3 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value3: " & ErrorValue3 & " ) | " & Chr(13)
    If ErrorValue4 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value4: " & ErrorValue4 & " ) | " & Chr(13)
    ErrorMessage = ErrorMessage & "------------------------------------------------------------------------------"
    Temp = MsgBox(ErrorMessage, vbExclamation, "ERROR")

End Sub

' Asks Yes/No Question (No will raise an Error)
Private Function Ask(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR) As Boolean

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant
    Dim Temp As Boolean

    ProtInit
    ErrorArray = Error_Get(ErrorCategory, ErrorIndex)
    ErrorMessage = "( " & "Category: " & Error_GetCategory(ErrorCategory)    & " ) | " & Chr(13) & _
                   "( " & "Severity: " & ErrorArray(0)                       & " ) | " & Chr(13) & _
                   "( " & "Index: "    & ErrorArray(1)                       & " ) | " & Chr(13) & _
                   "( " & "Message: "  & ErrorArray(2)                       & " ) | " & Chr(13)
    If ErrorValue1 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value1: " & ErrorValue1 & " ) | " & Chr(13)
    If ErrorValue2 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value2: " & ErrorValue2 & " ) | " & Chr(13)
    If ErrorValue3 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value3: " & ErrorValue3 & " ) | " & Chr(13)
    If ErrorValue4 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & "( " & "Value4: " & ErrorValue4 & " ) | " & Chr(13)
    ErrorMessage = ErrorMessage & "------------------------------------------------------------------------------"
    Temp = MsgBox(ErrorMessage, vbYesNo, "Question")
    If Temp = False Then Ask = IS_ERROR

End Function

' Gets Errormessage
Private Function Error_Get(ErrorCategory As Long, ErrorIndex As Long) As Variant()
    
    Dim Temp(2) As Variant
    Select Case ErrorCategory
        Case 0,1,2,3,4,5,6,7:    Temp(0) = AL_Error_Catalog(ErrorCategory, 0, ErrorIndex): Temp(1) = ErrorIndex: Temp(2) = AL_Error_Catalog(ErrorCategory, 1, ErrorIndex)
        Case Else
                                 Temp(0) = 0: Temp(1) = 0: Temp(2) = "ErrorCategory not defined"
    End Select
    Error_Get = Temp

End Function

' Gets Name of Errorcategory
Private Function Error_GetCategory(ErrorCategory As Long) As String

    Select Case ErrorCategory
        Case 0: Error_GetCategory = "System"
        Case 1: Error_GetCategory = "Workbook"
        Case 2: Error_GetCategory = "Worksheet"
        Case 3: Error_GetCategory = "Linker"
        Case 4: Error_GetCategory = "Compiler"
        Case 5: Error_GetCategory = "Module"
        Case 6: Error_GetCategory = "Class"
        Case 7: Error_GetCategory = "Userform"
        Case Else
                Error_GetCategory = "UNKNOWN"
    End Select

End Function

' Runs once to Initialize all Errormessages
Private Sub ProtInit()
    
    If Initialized = False Then
    ' System-Errors
        AL_Error_Catalog(0000, 0000, 0000) = 1000: AL_Error_Catalog(0000, 0001, 0000) = "ErrorCategory doesnt exist"
        AL_Error_Catalog(0000, 0000, 0001) = 1000: AL_Error_Catalog(0000, 0001, 0001) = "Value isnt available"
        AL_Error_Catalog(0000, 0000, 0002) = 1000: AL_Error_Catalog(0000, 0001, 0002) = "Value is Empty"
        AL_Error_Catalog(0000, 0000, 0003) = 1000: AL_Error_Catalog(0000, 0001, 0003) = "Value is Nothing"
        AL_Error_Catalog(0000, 0000, 0004) = 1000: AL_Error_Catalog(0000, 0001, 0004) = "Value Overflow"
        AL_Error_Catalog(0000, 0000, 0005) = 1000: AL_Error_Catalog(0000, 0001, 0005) = "Value Underflow"
        AL_Error_Catalog(0000, 0000, 0006) = 1000: AL_Error_Catalog(0000, 0001, 0006) = "Value1 doesnt equal Value2"
        AL_Error_Catalog(0000, 0000, 0007) = 1000: AL_Error_Catalog(0000, 0001, 0007) = "Value1 is smaller than or equal to Value2"
        AL_Error_Catalog(0000, 0000, 0008) = 1000: AL_Error_Catalog(0000, 0001, 0008) = "Value1 is bigger than or equal to Value2"
        AL_Error_Catalog(0000, 0000, 0009) = 1000: AL_Error_Catalog(0000, 0001, 0009) = "Value1 is smaller than Value2"
        AL_Error_Catalog(0000, 0000, 0010) = 1000: AL_Error_Catalog(0000, 0001, 0010) = "Value1 is bigger than Value2"
        AL_Error_Catalog(0000, 0000, 0011) = 1000: AL_Error_Catalog(0000, 0001, 0011) = "Value1 is Value2"
        AL_Error_Catalog(0000, 0000, 0012) = 1000: AL_Error_Catalog(0000, 0001, 0012) = "Several Values are Empty"
        AL_Error_Catalog(0000, 0000, 0013) = 1000: AL_Error_Catalog(0000, 0001, 0013) = "To many Values arent Empty"
        AL_Error_Catalog(0000, 0000, 0014) = 1000: AL_Error_Catalog(0000, 0001, 0014) = "Value Is Something"
        AL_Error_Catalog(0000, 0000, 0015) = 1000: AL_Error_Catalog(0000, 0001, 0015) = "Value1 equals Value2"
        AL_Error_Catalog(0000, 0000, 0016) = 1000: AL_Error_Catalog(0000, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0017) = 1000: AL_Error_Catalog(0000, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0018) = 1000: AL_Error_Catalog(0000, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0019) = 1000: AL_Error_Catalog(0000, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0020) = 1000: AL_Error_Catalog(0000, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0021) = 1000: AL_Error_Catalog(0000, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0022) = 1000: AL_Error_Catalog(0000, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0023) = 1000: AL_Error_Catalog(0000, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0024) = 1000: AL_Error_Catalog(0000, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0025) = 1000: AL_Error_Catalog(0000, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0026) = 1000: AL_Error_Catalog(0000, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0027) = 1000: AL_Error_Catalog(0000, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0028) = 1000: AL_Error_Catalog(0000, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0029) = 1000: AL_Error_Catalog(0000, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0030) = 1000: AL_Error_Catalog(0000, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0031) = 1000: AL_Error_Catalog(0000, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0032) = 1000: AL_Error_Catalog(0000, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0033) = 1000: AL_Error_Catalog(0000, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0034) = 1000: AL_Error_Catalog(0000, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0035) = 1000: AL_Error_Catalog(0000, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0036) = 1000: AL_Error_Catalog(0000, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0037) = 1000: AL_Error_Catalog(0000, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0038) = 1000: AL_Error_Catalog(0000, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0039) = 1000: AL_Error_Catalog(0000, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0040) = 1000: AL_Error_Catalog(0000, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0041) = 1000: AL_Error_Catalog(0000, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0042) = 1000: AL_Error_Catalog(0000, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0043) = 1000: AL_Error_Catalog(0000, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0044) = 1000: AL_Error_Catalog(0000, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0045) = 1000: AL_Error_Catalog(0000, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0046) = 1000: AL_Error_Catalog(0000, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0047) = 1000: AL_Error_Catalog(0000, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0048) = 1000: AL_Error_Catalog(0000, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0049) = 1000: AL_Error_Catalog(0000, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0050) = 1000: AL_Error_Catalog(0000, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0051) = 1000: AL_Error_Catalog(0000, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0052) = 1000: AL_Error_Catalog(0000, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0053) = 1000: AL_Error_Catalog(0000, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0054) = 1000: AL_Error_Catalog(0000, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0055) = 1000: AL_Error_Catalog(0000, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0056) = 1000: AL_Error_Catalog(0000, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0057) = 1000: AL_Error_Catalog(0000, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0058) = 1000: AL_Error_Catalog(0000, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0059) = 1000: AL_Error_Catalog(0000, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0060) = 1000: AL_Error_Catalog(0000, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0061) = 1000: AL_Error_Catalog(0000, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0062) = 1000: AL_Error_Catalog(0000, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0063) = 1000: AL_Error_Catalog(0000, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0064) = 1000: AL_Error_Catalog(0000, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0065) = 1000: AL_Error_Catalog(0000, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0066) = 1000: AL_Error_Catalog(0000, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0067) = 1000: AL_Error_Catalog(0000, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0068) = 1000: AL_Error_Catalog(0000, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0069) = 1000: AL_Error_Catalog(0000, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0070) = 1000: AL_Error_Catalog(0000, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0071) = 1000: AL_Error_Catalog(0000, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0072) = 1000: AL_Error_Catalog(0000, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0073) = 1000: AL_Error_Catalog(0000, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0074) = 1000: AL_Error_Catalog(0000, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0075) = 1000: AL_Error_Catalog(0000, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0076) = 1000: AL_Error_Catalog(0000, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0077) = 1000: AL_Error_Catalog(0000, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0078) = 1000: AL_Error_Catalog(0000, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0079) = 1000: AL_Error_Catalog(0000, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0080) = 1000: AL_Error_Catalog(0000, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0081) = 1000: AL_Error_Catalog(0000, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0082) = 1000: AL_Error_Catalog(0000, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0083) = 1000: AL_Error_Catalog(0000, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0084) = 1000: AL_Error_Catalog(0000, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0085) = 1000: AL_Error_Catalog(0000, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0086) = 1000: AL_Error_Catalog(0000, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0087) = 1000: AL_Error_Catalog(0000, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0088) = 1000: AL_Error_Catalog(0000, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0089) = 1000: AL_Error_Catalog(0000, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0090) = 1000: AL_Error_Catalog(0000, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0091) = 1000: AL_Error_Catalog(0000, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0092) = 1000: AL_Error_Catalog(0000, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0093) = 1000: AL_Error_Catalog(0000, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0094) = 1000: AL_Error_Catalog(0000, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0095) = 1000: AL_Error_Catalog(0000, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0096) = 1000: AL_Error_Catalog(0000, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0097) = 1000: AL_Error_Catalog(0000, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0098) = 1000: AL_Error_Catalog(0000, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0099) = 1000: AL_Error_Catalog(0000, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0000) = 1000: AL_Error_Catalog(0001, 0001, 0000) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0001) = 1000: AL_Error_Catalog(0001, 0001, 0001) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0002) = 1000: AL_Error_Catalog(0001, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0003) = 1000: AL_Error_Catalog(0001, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0004) = 1000: AL_Error_Catalog(0001, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0005) = 1000: AL_Error_Catalog(0001, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0006) = 1000: AL_Error_Catalog(0001, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0007) = 1000: AL_Error_Catalog(0001, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0008) = 1000: AL_Error_Catalog(0001, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0009) = 1000: AL_Error_Catalog(0001, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0010) = 1000: AL_Error_Catalog(0001, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0011) = 1000: AL_Error_Catalog(0001, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0012) = 1000: AL_Error_Catalog(0001, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0013) = 1000: AL_Error_Catalog(0001, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0014) = 1000: AL_Error_Catalog(0001, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0015) = 1000: AL_Error_Catalog(0001, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0016) = 1000: AL_Error_Catalog(0001, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0017) = 1000: AL_Error_Catalog(0001, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0018) = 1000: AL_Error_Catalog(0001, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0019) = 1000: AL_Error_Catalog(0001, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0020) = 1000: AL_Error_Catalog(0001, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0021) = 1000: AL_Error_Catalog(0001, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0022) = 1000: AL_Error_Catalog(0001, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0023) = 1000: AL_Error_Catalog(0001, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0024) = 1000: AL_Error_Catalog(0001, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0025) = 1000: AL_Error_Catalog(0001, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0026) = 1000: AL_Error_Catalog(0001, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0027) = 1000: AL_Error_Catalog(0001, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0028) = 1000: AL_Error_Catalog(0001, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0029) = 1000: AL_Error_Catalog(0001, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0030) = 1000: AL_Error_Catalog(0001, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0031) = 1000: AL_Error_Catalog(0001, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0032) = 1000: AL_Error_Catalog(0001, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0033) = 1000: AL_Error_Catalog(0001, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0034) = 1000: AL_Error_Catalog(0001, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0035) = 1000: AL_Error_Catalog(0001, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0036) = 1000: AL_Error_Catalog(0001, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0037) = 1000: AL_Error_Catalog(0001, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0038) = 1000: AL_Error_Catalog(0001, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0039) = 1000: AL_Error_Catalog(0001, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0040) = 1000: AL_Error_Catalog(0001, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0041) = 1000: AL_Error_Catalog(0001, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0042) = 1000: AL_Error_Catalog(0001, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0043) = 1000: AL_Error_Catalog(0001, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0044) = 1000: AL_Error_Catalog(0001, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0045) = 1000: AL_Error_Catalog(0001, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0046) = 1000: AL_Error_Catalog(0001, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0047) = 1000: AL_Error_Catalog(0001, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0048) = 1000: AL_Error_Catalog(0001, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0049) = 1000: AL_Error_Catalog(0001, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0050) = 1000: AL_Error_Catalog(0001, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0051) = 1000: AL_Error_Catalog(0001, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0052) = 1000: AL_Error_Catalog(0001, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0053) = 1000: AL_Error_Catalog(0001, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0054) = 1000: AL_Error_Catalog(0001, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0055) = 1000: AL_Error_Catalog(0001, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0056) = 1000: AL_Error_Catalog(0001, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0057) = 1000: AL_Error_Catalog(0001, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0058) = 1000: AL_Error_Catalog(0001, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0059) = 1000: AL_Error_Catalog(0001, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0060) = 1000: AL_Error_Catalog(0001, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0061) = 1000: AL_Error_Catalog(0001, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0062) = 1000: AL_Error_Catalog(0001, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0063) = 1000: AL_Error_Catalog(0001, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0064) = 1000: AL_Error_Catalog(0001, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0065) = 1000: AL_Error_Catalog(0001, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0066) = 1000: AL_Error_Catalog(0001, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0067) = 1000: AL_Error_Catalog(0001, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0068) = 1000: AL_Error_Catalog(0001, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0069) = 1000: AL_Error_Catalog(0001, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0070) = 1000: AL_Error_Catalog(0001, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0071) = 1000: AL_Error_Catalog(0001, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0072) = 1000: AL_Error_Catalog(0001, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0073) = 1000: AL_Error_Catalog(0001, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0074) = 1000: AL_Error_Catalog(0001, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0075) = 1000: AL_Error_Catalog(0001, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0076) = 1000: AL_Error_Catalog(0001, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0077) = 1000: AL_Error_Catalog(0001, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0078) = 1000: AL_Error_Catalog(0001, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0079) = 1000: AL_Error_Catalog(0001, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0080) = 1000: AL_Error_Catalog(0001, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0081) = 1000: AL_Error_Catalog(0001, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0082) = 1000: AL_Error_Catalog(0001, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0083) = 1000: AL_Error_Catalog(0001, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0084) = 1000: AL_Error_Catalog(0001, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0085) = 1000: AL_Error_Catalog(0001, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0086) = 1000: AL_Error_Catalog(0001, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0087) = 1000: AL_Error_Catalog(0001, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0088) = 1000: AL_Error_Catalog(0001, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0089) = 1000: AL_Error_Catalog(0001, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0090) = 1000: AL_Error_Catalog(0001, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0091) = 1000: AL_Error_Catalog(0001, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0092) = 1000: AL_Error_Catalog(0001, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0093) = 1000: AL_Error_Catalog(0001, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0094) = 1000: AL_Error_Catalog(0001, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0095) = 1000: AL_Error_Catalog(0001, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0096) = 1000: AL_Error_Catalog(0001, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0097) = 1000: AL_Error_Catalog(0001, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0098) = 1000: AL_Error_Catalog(0001, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0099) = 1000: AL_Error_Catalog(0001, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0000) = 1000: AL_Error_Catalog(0002, 0001, 0000) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0001) = 1000: AL_Error_Catalog(0002, 0001, 0001) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0002) = 1000: AL_Error_Catalog(0002, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0003) = 1000: AL_Error_Catalog(0002, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0004) = 1000: AL_Error_Catalog(0002, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0005) = 1000: AL_Error_Catalog(0002, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0006) = 1000: AL_Error_Catalog(0002, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0007) = 1000: AL_Error_Catalog(0002, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0008) = 1000: AL_Error_Catalog(0002, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0009) = 1000: AL_Error_Catalog(0002, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0010) = 1000: AL_Error_Catalog(0002, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0011) = 1000: AL_Error_Catalog(0002, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0012) = 1000: AL_Error_Catalog(0002, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0013) = 1000: AL_Error_Catalog(0002, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0014) = 1000: AL_Error_Catalog(0002, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0015) = 1000: AL_Error_Catalog(0002, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0016) = 1000: AL_Error_Catalog(0002, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0017) = 1000: AL_Error_Catalog(0002, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0018) = 1000: AL_Error_Catalog(0002, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0019) = 1000: AL_Error_Catalog(0002, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0020) = 1000: AL_Error_Catalog(0002, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0021) = 1000: AL_Error_Catalog(0002, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0022) = 1000: AL_Error_Catalog(0002, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0023) = 1000: AL_Error_Catalog(0002, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0024) = 1000: AL_Error_Catalog(0002, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0025) = 1000: AL_Error_Catalog(0002, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0026) = 1000: AL_Error_Catalog(0002, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0027) = 1000: AL_Error_Catalog(0002, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0028) = 1000: AL_Error_Catalog(0002, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0029) = 1000: AL_Error_Catalog(0002, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0030) = 1000: AL_Error_Catalog(0002, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0031) = 1000: AL_Error_Catalog(0002, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0032) = 1000: AL_Error_Catalog(0002, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0033) = 1000: AL_Error_Catalog(0002, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0034) = 1000: AL_Error_Catalog(0002, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0035) = 1000: AL_Error_Catalog(0002, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0036) = 1000: AL_Error_Catalog(0002, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0037) = 1000: AL_Error_Catalog(0002, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0038) = 1000: AL_Error_Catalog(0002, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0039) = 1000: AL_Error_Catalog(0002, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0040) = 1000: AL_Error_Catalog(0002, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0041) = 1000: AL_Error_Catalog(0002, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0042) = 1000: AL_Error_Catalog(0002, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0043) = 1000: AL_Error_Catalog(0002, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0044) = 1000: AL_Error_Catalog(0002, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0045) = 1000: AL_Error_Catalog(0002, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0046) = 1000: AL_Error_Catalog(0002, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0047) = 1000: AL_Error_Catalog(0002, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0048) = 1000: AL_Error_Catalog(0002, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0049) = 1000: AL_Error_Catalog(0002, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0050) = 1000: AL_Error_Catalog(0002, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0051) = 1000: AL_Error_Catalog(0002, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0052) = 1000: AL_Error_Catalog(0002, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0053) = 1000: AL_Error_Catalog(0002, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0054) = 1000: AL_Error_Catalog(0002, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0055) = 1000: AL_Error_Catalog(0002, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0056) = 1000: AL_Error_Catalog(0002, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0057) = 1000: AL_Error_Catalog(0002, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0058) = 1000: AL_Error_Catalog(0002, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0059) = 1000: AL_Error_Catalog(0002, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0060) = 1000: AL_Error_Catalog(0002, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0061) = 1000: AL_Error_Catalog(0002, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0062) = 1000: AL_Error_Catalog(0002, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0063) = 1000: AL_Error_Catalog(0002, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0064) = 1000: AL_Error_Catalog(0002, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0065) = 1000: AL_Error_Catalog(0002, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0066) = 1000: AL_Error_Catalog(0002, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0067) = 1000: AL_Error_Catalog(0002, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0068) = 1000: AL_Error_Catalog(0002, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0069) = 1000: AL_Error_Catalog(0002, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0070) = 1000: AL_Error_Catalog(0002, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0071) = 1000: AL_Error_Catalog(0002, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0072) = 1000: AL_Error_Catalog(0002, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0073) = 1000: AL_Error_Catalog(0002, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0074) = 1000: AL_Error_Catalog(0002, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0075) = 1000: AL_Error_Catalog(0002, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0076) = 1000: AL_Error_Catalog(0002, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0077) = 1000: AL_Error_Catalog(0002, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0078) = 1000: AL_Error_Catalog(0002, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0079) = 1000: AL_Error_Catalog(0002, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0080) = 1000: AL_Error_Catalog(0002, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0081) = 1000: AL_Error_Catalog(0002, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0082) = 1000: AL_Error_Catalog(0002, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0083) = 1000: AL_Error_Catalog(0002, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0084) = 1000: AL_Error_Catalog(0002, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0085) = 1000: AL_Error_Catalog(0002, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0086) = 1000: AL_Error_Catalog(0002, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0087) = 1000: AL_Error_Catalog(0002, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0088) = 1000: AL_Error_Catalog(0002, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0089) = 1000: AL_Error_Catalog(0002, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0090) = 1000: AL_Error_Catalog(0002, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0091) = 1000: AL_Error_Catalog(0002, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0092) = 1000: AL_Error_Catalog(0002, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0093) = 1000: AL_Error_Catalog(0002, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0094) = 1000: AL_Error_Catalog(0002, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0095) = 1000: AL_Error_Catalog(0002, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0096) = 1000: AL_Error_Catalog(0002, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0097) = 1000: AL_Error_Catalog(0002, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0098) = 1000: AL_Error_Catalog(0002, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0099) = 1000: AL_Error_Catalog(0002, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0000) = 1000: AL_Error_Catalog(0003, 0001, 0000) = "Workbook doesnt exist"
        AL_Error_Catalog(0003, 0000, 0001) = 1000: AL_Error_Catalog(0003, 0001, 0001) = "Workbook already exists"
        AL_Error_Catalog(0003, 0000, 0002) = 1000: AL_Error_Catalog(0003, 0001, 0002) = "Dependency missing"
        AL_Error_Catalog(0003, 0000, 0003) = 1000: AL_Error_Catalog(0003, 0001, 0003) = "Object not Initialized"
        AL_Error_Catalog(0003, 0000, 0004) = 1000: AL_Error_Catalog(0003, 0001, 0004) = "Not available in Workbook"
        AL_Error_Catalog(0003, 0000, 0005) = 1000: AL_Error_Catalog(0003, 0001, 0005) = "Component doesnt exists"
        AL_Error_Catalog(0003, 0000, 0006) = 1000: AL_Error_Catalog(0003, 0001, 0006) = "Component already exists"
        AL_Error_Catalog(0003, 0000, 0007) = 1000: AL_Error_Catalog(0003, 0001, 0007) = "Couldnt Create Component"
        AL_Error_Catalog(0003, 0000, 0008) = 1000: AL_Error_Catalog(0003, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0009) = 1000: AL_Error_Catalog(0003, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0010) = 1000: AL_Error_Catalog(0003, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0011) = 1000: AL_Error_Catalog(0003, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0012) = 1000: AL_Error_Catalog(0003, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0013) = 1000: AL_Error_Catalog(0003, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0014) = 1000: AL_Error_Catalog(0003, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0015) = 1000: AL_Error_Catalog(0003, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0016) = 1000: AL_Error_Catalog(0003, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0017) = 1000: AL_Error_Catalog(0003, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0018) = 1000: AL_Error_Catalog(0003, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0019) = 1000: AL_Error_Catalog(0003, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0020) = 1000: AL_Error_Catalog(0003, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0021) = 1000: AL_Error_Catalog(0003, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0022) = 1000: AL_Error_Catalog(0003, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0023) = 1000: AL_Error_Catalog(0003, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0024) = 1000: AL_Error_Catalog(0003, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0025) = 1000: AL_Error_Catalog(0003, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0026) = 1000: AL_Error_Catalog(0003, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0027) = 1000: AL_Error_Catalog(0003, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0028) = 1000: AL_Error_Catalog(0003, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0029) = 1000: AL_Error_Catalog(0003, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0030) = 1000: AL_Error_Catalog(0003, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0031) = 1000: AL_Error_Catalog(0003, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0032) = 1000: AL_Error_Catalog(0003, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0033) = 1000: AL_Error_Catalog(0003, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0034) = 1000: AL_Error_Catalog(0003, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0035) = 1000: AL_Error_Catalog(0003, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0036) = 1000: AL_Error_Catalog(0003, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0037) = 1000: AL_Error_Catalog(0003, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0038) = 1000: AL_Error_Catalog(0003, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0039) = 1000: AL_Error_Catalog(0003, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0040) = 1000: AL_Error_Catalog(0003, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0041) = 1000: AL_Error_Catalog(0003, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0042) = 1000: AL_Error_Catalog(0003, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0043) = 1000: AL_Error_Catalog(0003, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0044) = 1000: AL_Error_Catalog(0003, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0045) = 1000: AL_Error_Catalog(0003, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0046) = 1000: AL_Error_Catalog(0003, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0047) = 1000: AL_Error_Catalog(0003, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0048) = 1000: AL_Error_Catalog(0003, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0049) = 1000: AL_Error_Catalog(0003, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0050) = 1000: AL_Error_Catalog(0003, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0051) = 1000: AL_Error_Catalog(0003, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0052) = 1000: AL_Error_Catalog(0003, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0053) = 1000: AL_Error_Catalog(0003, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0054) = 1000: AL_Error_Catalog(0003, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0055) = 1000: AL_Error_Catalog(0003, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0056) = 1000: AL_Error_Catalog(0003, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0057) = 1000: AL_Error_Catalog(0003, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0058) = 1000: AL_Error_Catalog(0003, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0059) = 1000: AL_Error_Catalog(0003, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0060) = 1000: AL_Error_Catalog(0003, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0061) = 1000: AL_Error_Catalog(0003, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0062) = 1000: AL_Error_Catalog(0003, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0063) = 1000: AL_Error_Catalog(0003, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0064) = 1000: AL_Error_Catalog(0003, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0065) = 1000: AL_Error_Catalog(0003, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0066) = 1000: AL_Error_Catalog(0003, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0067) = 1000: AL_Error_Catalog(0003, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0068) = 1000: AL_Error_Catalog(0003, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0069) = 1000: AL_Error_Catalog(0003, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0070) = 1000: AL_Error_Catalog(0003, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0071) = 1000: AL_Error_Catalog(0003, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0072) = 1000: AL_Error_Catalog(0003, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0073) = 1000: AL_Error_Catalog(0003, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0074) = 1000: AL_Error_Catalog(0003, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0075) = 1000: AL_Error_Catalog(0003, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0076) = 1000: AL_Error_Catalog(0003, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0077) = 1000: AL_Error_Catalog(0003, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0078) = 1000: AL_Error_Catalog(0003, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0079) = 1000: AL_Error_Catalog(0003, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0080) = 1000: AL_Error_Catalog(0003, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0081) = 1000: AL_Error_Catalog(0003, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0082) = 1000: AL_Error_Catalog(0003, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0083) = 1000: AL_Error_Catalog(0003, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0084) = 1000: AL_Error_Catalog(0003, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0085) = 1000: AL_Error_Catalog(0003, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0086) = 1000: AL_Error_Catalog(0003, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0087) = 1000: AL_Error_Catalog(0003, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0088) = 1000: AL_Error_Catalog(0003, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0089) = 1000: AL_Error_Catalog(0003, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0090) = 1000: AL_Error_Catalog(0003, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0091) = 1000: AL_Error_Catalog(0003, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0092) = 1000: AL_Error_Catalog(0003, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0093) = 1000: AL_Error_Catalog(0003, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0094) = 1000: AL_Error_Catalog(0003, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0095) = 1000: AL_Error_Catalog(0003, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0096) = 1000: AL_Error_Catalog(0003, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0097) = 1000: AL_Error_Catalog(0003, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0098) = 1000: AL_Error_Catalog(0003, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0099) = 1000: AL_Error_Catalog(0003, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0000) = 1000: AL_Error_Catalog(0004, 0001, 0000) = "Worksheet already exists"
        AL_Error_Catalog(0004, 0000, 0001) = 1000: AL_Error_Catalog(0004, 0001, 0001) = "Worksheet doesnt exist"
        AL_Error_Catalog(0004, 0000, 0002) = 1000: AL_Error_Catalog(0004, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0003) = 1000: AL_Error_Catalog(0004, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0004) = 1000: AL_Error_Catalog(0004, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0005) = 1000: AL_Error_Catalog(0004, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0006) = 1000: AL_Error_Catalog(0004, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0007) = 1000: AL_Error_Catalog(0004, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0008) = 1000: AL_Error_Catalog(0004, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0009) = 1000: AL_Error_Catalog(0004, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0010) = 1000: AL_Error_Catalog(0004, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0011) = 1000: AL_Error_Catalog(0004, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0012) = 1000: AL_Error_Catalog(0004, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0013) = 1000: AL_Error_Catalog(0004, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0014) = 1000: AL_Error_Catalog(0004, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0015) = 1000: AL_Error_Catalog(0004, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0016) = 1000: AL_Error_Catalog(0004, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0017) = 1000: AL_Error_Catalog(0004, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0018) = 1000: AL_Error_Catalog(0004, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0019) = 1000: AL_Error_Catalog(0004, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0020) = 1000: AL_Error_Catalog(0004, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0021) = 1000: AL_Error_Catalog(0004, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0022) = 1000: AL_Error_Catalog(0004, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0023) = 1000: AL_Error_Catalog(0004, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0024) = 1000: AL_Error_Catalog(0004, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0025) = 1000: AL_Error_Catalog(0004, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0026) = 1000: AL_Error_Catalog(0004, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0027) = 1000: AL_Error_Catalog(0004, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0028) = 1000: AL_Error_Catalog(0004, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0029) = 1000: AL_Error_Catalog(0004, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0030) = 1000: AL_Error_Catalog(0004, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0031) = 1000: AL_Error_Catalog(0004, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0032) = 1000: AL_Error_Catalog(0004, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0033) = 1000: AL_Error_Catalog(0004, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0034) = 1000: AL_Error_Catalog(0004, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0035) = 1000: AL_Error_Catalog(0004, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0036) = 1000: AL_Error_Catalog(0004, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0037) = 1000: AL_Error_Catalog(0004, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0038) = 1000: AL_Error_Catalog(0004, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0039) = 1000: AL_Error_Catalog(0004, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0040) = 1000: AL_Error_Catalog(0004, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0041) = 1000: AL_Error_Catalog(0004, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0042) = 1000: AL_Error_Catalog(0004, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0043) = 1000: AL_Error_Catalog(0004, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0044) = 1000: AL_Error_Catalog(0004, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0045) = 1000: AL_Error_Catalog(0004, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0046) = 1000: AL_Error_Catalog(0004, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0047) = 1000: AL_Error_Catalog(0004, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0048) = 1000: AL_Error_Catalog(0004, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0049) = 1000: AL_Error_Catalog(0004, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0050) = 1000: AL_Error_Catalog(0004, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0051) = 1000: AL_Error_Catalog(0004, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0052) = 1000: AL_Error_Catalog(0004, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0053) = 1000: AL_Error_Catalog(0004, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0054) = 1000: AL_Error_Catalog(0004, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0055) = 1000: AL_Error_Catalog(0004, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0056) = 1000: AL_Error_Catalog(0004, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0057) = 1000: AL_Error_Catalog(0004, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0058) = 1000: AL_Error_Catalog(0004, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0059) = 1000: AL_Error_Catalog(0004, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0060) = 1000: AL_Error_Catalog(0004, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0061) = 1000: AL_Error_Catalog(0004, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0062) = 1000: AL_Error_Catalog(0004, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0063) = 1000: AL_Error_Catalog(0004, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0064) = 1000: AL_Error_Catalog(0004, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0065) = 1000: AL_Error_Catalog(0004, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0066) = 1000: AL_Error_Catalog(0004, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0067) = 1000: AL_Error_Catalog(0004, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0068) = 1000: AL_Error_Catalog(0004, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0069) = 1000: AL_Error_Catalog(0004, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0070) = 1000: AL_Error_Catalog(0004, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0071) = 1000: AL_Error_Catalog(0004, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0072) = 1000: AL_Error_Catalog(0004, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0073) = 1000: AL_Error_Catalog(0004, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0074) = 1000: AL_Error_Catalog(0004, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0075) = 1000: AL_Error_Catalog(0004, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0076) = 1000: AL_Error_Catalog(0004, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0077) = 1000: AL_Error_Catalog(0004, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0078) = 1000: AL_Error_Catalog(0004, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0079) = 1000: AL_Error_Catalog(0004, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0080) = 1000: AL_Error_Catalog(0004, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0081) = 1000: AL_Error_Catalog(0004, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0082) = 1000: AL_Error_Catalog(0004, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0083) = 1000: AL_Error_Catalog(0004, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0084) = 1000: AL_Error_Catalog(0004, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0085) = 1000: AL_Error_Catalog(0004, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0086) = 1000: AL_Error_Catalog(0004, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0087) = 1000: AL_Error_Catalog(0004, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0088) = 1000: AL_Error_Catalog(0004, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0089) = 1000: AL_Error_Catalog(0004, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0090) = 1000: AL_Error_Catalog(0004, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0091) = 1000: AL_Error_Catalog(0004, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0092) = 1000: AL_Error_Catalog(0004, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0093) = 1000: AL_Error_Catalog(0004, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0094) = 1000: AL_Error_Catalog(0004, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0095) = 1000: AL_Error_Catalog(0004, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0096) = 1000: AL_Error_Catalog(0004, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0097) = 1000: AL_Error_Catalog(0004, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0098) = 1000: AL_Error_Catalog(0004, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0099) = 1000: AL_Error_Catalog(0004, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0000) = 1000: AL_Error_Catalog(0005, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0005, 0000, 0001) = 1000: AL_Error_Catalog(0005, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0005, 0000, 0002) = 1000: AL_Error_Catalog(0005, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0005, 0000, 0003) = 1000: AL_Error_Catalog(0005, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0005, 0000, 0004) = 1000: AL_Error_Catalog(0005, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0005, 0000, 0005) = 1000: AL_Error_Catalog(0005, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0006) = 1000: AL_Error_Catalog(0005, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0007) = 1000: AL_Error_Catalog(0005, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0008) = 1000: AL_Error_Catalog(0005, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0009) = 1000: AL_Error_Catalog(0005, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0010) = 1000: AL_Error_Catalog(0005, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0011) = 1000: AL_Error_Catalog(0005, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0012) = 1000: AL_Error_Catalog(0005, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0013) = 1000: AL_Error_Catalog(0005, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0014) = 1000: AL_Error_Catalog(0005, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0015) = 1000: AL_Error_Catalog(0005, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0016) = 1000: AL_Error_Catalog(0005, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0017) = 1000: AL_Error_Catalog(0005, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0018) = 1000: AL_Error_Catalog(0005, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0019) = 1000: AL_Error_Catalog(0005, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0020) = 1000: AL_Error_Catalog(0005, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0021) = 1000: AL_Error_Catalog(0005, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0022) = 1000: AL_Error_Catalog(0005, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0023) = 1000: AL_Error_Catalog(0005, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0024) = 1000: AL_Error_Catalog(0005, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0025) = 1000: AL_Error_Catalog(0005, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0026) = 1000: AL_Error_Catalog(0005, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0027) = 1000: AL_Error_Catalog(0005, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0028) = 1000: AL_Error_Catalog(0005, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0029) = 1000: AL_Error_Catalog(0005, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0030) = 1000: AL_Error_Catalog(0005, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0031) = 1000: AL_Error_Catalog(0005, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0032) = 1000: AL_Error_Catalog(0005, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0033) = 1000: AL_Error_Catalog(0005, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0034) = 1000: AL_Error_Catalog(0005, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0035) = 1000: AL_Error_Catalog(0005, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0036) = 1000: AL_Error_Catalog(0005, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0037) = 1000: AL_Error_Catalog(0005, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0038) = 1000: AL_Error_Catalog(0005, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0039) = 1000: AL_Error_Catalog(0005, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0040) = 1000: AL_Error_Catalog(0005, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0041) = 1000: AL_Error_Catalog(0005, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0042) = 1000: AL_Error_Catalog(0005, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0043) = 1000: AL_Error_Catalog(0005, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0044) = 1000: AL_Error_Catalog(0005, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0045) = 1000: AL_Error_Catalog(0005, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0046) = 1000: AL_Error_Catalog(0005, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0047) = 1000: AL_Error_Catalog(0005, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0048) = 1000: AL_Error_Catalog(0005, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0049) = 1000: AL_Error_Catalog(0005, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0050) = 1000: AL_Error_Catalog(0005, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0051) = 1000: AL_Error_Catalog(0005, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0052) = 1000: AL_Error_Catalog(0005, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0053) = 1000: AL_Error_Catalog(0005, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0054) = 1000: AL_Error_Catalog(0005, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0055) = 1000: AL_Error_Catalog(0005, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0056) = 1000: AL_Error_Catalog(0005, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0057) = 1000: AL_Error_Catalog(0005, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0058) = 1000: AL_Error_Catalog(0005, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0059) = 1000: AL_Error_Catalog(0005, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0060) = 1000: AL_Error_Catalog(0005, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0061) = 1000: AL_Error_Catalog(0005, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0062) = 1000: AL_Error_Catalog(0005, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0063) = 1000: AL_Error_Catalog(0005, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0064) = 1000: AL_Error_Catalog(0005, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0065) = 1000: AL_Error_Catalog(0005, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0066) = 1000: AL_Error_Catalog(0005, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0067) = 1000: AL_Error_Catalog(0005, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0068) = 1000: AL_Error_Catalog(0005, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0069) = 1000: AL_Error_Catalog(0005, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0070) = 1000: AL_Error_Catalog(0005, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0071) = 1000: AL_Error_Catalog(0005, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0072) = 1000: AL_Error_Catalog(0005, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0073) = 1000: AL_Error_Catalog(0005, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0074) = 1000: AL_Error_Catalog(0005, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0075) = 1000: AL_Error_Catalog(0005, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0076) = 1000: AL_Error_Catalog(0005, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0077) = 1000: AL_Error_Catalog(0005, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0078) = 1000: AL_Error_Catalog(0005, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0079) = 1000: AL_Error_Catalog(0005, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0080) = 1000: AL_Error_Catalog(0005, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0081) = 1000: AL_Error_Catalog(0005, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0082) = 1000: AL_Error_Catalog(0005, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0083) = 1000: AL_Error_Catalog(0005, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0084) = 1000: AL_Error_Catalog(0005, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0085) = 1000: AL_Error_Catalog(0005, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0086) = 1000: AL_Error_Catalog(0005, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0087) = 1000: AL_Error_Catalog(0005, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0088) = 1000: AL_Error_Catalog(0005, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0089) = 1000: AL_Error_Catalog(0005, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0090) = 1000: AL_Error_Catalog(0005, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0091) = 1000: AL_Error_Catalog(0005, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0092) = 1000: AL_Error_Catalog(0005, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0093) = 1000: AL_Error_Catalog(0005, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0094) = 1000: AL_Error_Catalog(0005, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0095) = 1000: AL_Error_Catalog(0005, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0096) = 1000: AL_Error_Catalog(0005, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0097) = 1000: AL_Error_Catalog(0005, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0098) = 1000: AL_Error_Catalog(0005, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0099) = 1000: AL_Error_Catalog(0005, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0000) = 1000: AL_Error_Catalog(0006, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0006, 0000, 0001) = 1000: AL_Error_Catalog(0006, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0006, 0000, 0002) = 1000: AL_Error_Catalog(0006, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0006, 0000, 0003) = 1000: AL_Error_Catalog(0006, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0006, 0000, 0004) = 1000: AL_Error_Catalog(0006, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0006, 0000, 0005) = 1000: AL_Error_Catalog(0006, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0006) = 1000: AL_Error_Catalog(0006, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0007) = 1000: AL_Error_Catalog(0006, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0008) = 1000: AL_Error_Catalog(0006, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0009) = 1000: AL_Error_Catalog(0006, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0010) = 1000: AL_Error_Catalog(0006, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0011) = 1000: AL_Error_Catalog(0006, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0012) = 1000: AL_Error_Catalog(0006, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0013) = 1000: AL_Error_Catalog(0006, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0014) = 1000: AL_Error_Catalog(0006, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0015) = 1000: AL_Error_Catalog(0006, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0016) = 1000: AL_Error_Catalog(0006, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0017) = 1000: AL_Error_Catalog(0006, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0018) = 1000: AL_Error_Catalog(0006, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0019) = 1000: AL_Error_Catalog(0006, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0020) = 1000: AL_Error_Catalog(0006, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0021) = 1000: AL_Error_Catalog(0006, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0022) = 1000: AL_Error_Catalog(0006, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0023) = 1000: AL_Error_Catalog(0006, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0024) = 1000: AL_Error_Catalog(0006, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0025) = 1000: AL_Error_Catalog(0006, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0026) = 1000: AL_Error_Catalog(0006, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0027) = 1000: AL_Error_Catalog(0006, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0028) = 1000: AL_Error_Catalog(0006, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0029) = 1000: AL_Error_Catalog(0006, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0030) = 1000: AL_Error_Catalog(0006, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0031) = 1000: AL_Error_Catalog(0006, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0032) = 1000: AL_Error_Catalog(0006, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0033) = 1000: AL_Error_Catalog(0006, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0034) = 1000: AL_Error_Catalog(0006, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0035) = 1000: AL_Error_Catalog(0006, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0036) = 1000: AL_Error_Catalog(0006, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0037) = 1000: AL_Error_Catalog(0006, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0038) = 1000: AL_Error_Catalog(0006, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0039) = 1000: AL_Error_Catalog(0006, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0040) = 1000: AL_Error_Catalog(0006, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0041) = 1000: AL_Error_Catalog(0006, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0042) = 1000: AL_Error_Catalog(0006, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0043) = 1000: AL_Error_Catalog(0006, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0044) = 1000: AL_Error_Catalog(0006, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0045) = 1000: AL_Error_Catalog(0006, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0046) = 1000: AL_Error_Catalog(0006, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0047) = 1000: AL_Error_Catalog(0006, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0048) = 1000: AL_Error_Catalog(0006, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0049) = 1000: AL_Error_Catalog(0006, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0050) = 1000: AL_Error_Catalog(0006, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0051) = 1000: AL_Error_Catalog(0006, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0052) = 1000: AL_Error_Catalog(0006, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0053) = 1000: AL_Error_Catalog(0006, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0054) = 1000: AL_Error_Catalog(0006, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0055) = 1000: AL_Error_Catalog(0006, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0056) = 1000: AL_Error_Catalog(0006, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0057) = 1000: AL_Error_Catalog(0006, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0058) = 1000: AL_Error_Catalog(0006, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0059) = 1000: AL_Error_Catalog(0006, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0060) = 1000: AL_Error_Catalog(0006, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0061) = 1000: AL_Error_Catalog(0006, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0062) = 1000: AL_Error_Catalog(0006, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0063) = 1000: AL_Error_Catalog(0006, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0064) = 1000: AL_Error_Catalog(0006, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0065) = 1000: AL_Error_Catalog(0006, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0066) = 1000: AL_Error_Catalog(0006, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0067) = 1000: AL_Error_Catalog(0006, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0068) = 1000: AL_Error_Catalog(0006, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0069) = 1000: AL_Error_Catalog(0006, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0070) = 1000: AL_Error_Catalog(0006, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0071) = 1000: AL_Error_Catalog(0006, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0072) = 1000: AL_Error_Catalog(0006, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0073) = 1000: AL_Error_Catalog(0006, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0074) = 1000: AL_Error_Catalog(0006, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0075) = 1000: AL_Error_Catalog(0006, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0076) = 1000: AL_Error_Catalog(0006, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0077) = 1000: AL_Error_Catalog(0006, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0078) = 1000: AL_Error_Catalog(0006, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0079) = 1000: AL_Error_Catalog(0006, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0080) = 1000: AL_Error_Catalog(0006, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0081) = 1000: AL_Error_Catalog(0006, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0082) = 1000: AL_Error_Catalog(0006, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0083) = 1000: AL_Error_Catalog(0006, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0084) = 1000: AL_Error_Catalog(0006, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0085) = 1000: AL_Error_Catalog(0006, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0086) = 1000: AL_Error_Catalog(0006, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0087) = 1000: AL_Error_Catalog(0006, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0088) = 1000: AL_Error_Catalog(0006, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0089) = 1000: AL_Error_Catalog(0006, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0090) = 1000: AL_Error_Catalog(0006, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0091) = 1000: AL_Error_Catalog(0006, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0092) = 1000: AL_Error_Catalog(0006, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0093) = 1000: AL_Error_Catalog(0006, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0094) = 1000: AL_Error_Catalog(0006, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0095) = 1000: AL_Error_Catalog(0006, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0096) = 1000: AL_Error_Catalog(0006, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0097) = 1000: AL_Error_Catalog(0006, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0098) = 1000: AL_Error_Catalog(0006, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0099) = 1000: AL_Error_Catalog(0006, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0000) = 1000: AL_Error_Catalog(0007, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0007, 0000, 0001) = 1000: AL_Error_Catalog(0007, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0007, 0000, 0002) = 1000: AL_Error_Catalog(0007, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0007, 0000, 0003) = 1000: AL_Error_Catalog(0007, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0007, 0000, 0004) = 1000: AL_Error_Catalog(0007, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0007, 0000, 0005) = 1000: AL_Error_Catalog(0007, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0006) = 1000: AL_Error_Catalog(0007, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0007) = 1000: AL_Error_Catalog(0007, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0008) = 1000: AL_Error_Catalog(0007, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0009) = 1000: AL_Error_Catalog(0007, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0010) = 1000: AL_Error_Catalog(0007, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0011) = 1000: AL_Error_Catalog(0007, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0012) = 1000: AL_Error_Catalog(0007, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0013) = 1000: AL_Error_Catalog(0007, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0014) = 1000: AL_Error_Catalog(0007, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0015) = 1000: AL_Error_Catalog(0007, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0016) = 1000: AL_Error_Catalog(0007, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0017) = 1000: AL_Error_Catalog(0007, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0018) = 1000: AL_Error_Catalog(0007, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0019) = 1000: AL_Error_Catalog(0007, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0020) = 1000: AL_Error_Catalog(0007, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0021) = 1000: AL_Error_Catalog(0007, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0022) = 1000: AL_Error_Catalog(0007, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0023) = 1000: AL_Error_Catalog(0007, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0024) = 1000: AL_Error_Catalog(0007, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0025) = 1000: AL_Error_Catalog(0007, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0026) = 1000: AL_Error_Catalog(0007, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0027) = 1000: AL_Error_Catalog(0007, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0028) = 1000: AL_Error_Catalog(0007, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0029) = 1000: AL_Error_Catalog(0007, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0030) = 1000: AL_Error_Catalog(0007, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0031) = 1000: AL_Error_Catalog(0007, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0032) = 1000: AL_Error_Catalog(0007, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0033) = 1000: AL_Error_Catalog(0007, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0034) = 1000: AL_Error_Catalog(0007, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0035) = 1000: AL_Error_Catalog(0007, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0036) = 1000: AL_Error_Catalog(0007, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0037) = 1000: AL_Error_Catalog(0007, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0038) = 1000: AL_Error_Catalog(0007, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0039) = 1000: AL_Error_Catalog(0007, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0040) = 1000: AL_Error_Catalog(0007, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0041) = 1000: AL_Error_Catalog(0007, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0042) = 1000: AL_Error_Catalog(0007, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0043) = 1000: AL_Error_Catalog(0007, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0044) = 1000: AL_Error_Catalog(0007, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0045) = 1000: AL_Error_Catalog(0007, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0046) = 1000: AL_Error_Catalog(0007, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0047) = 1000: AL_Error_Catalog(0007, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0048) = 1000: AL_Error_Catalog(0007, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0049) = 1000: AL_Error_Catalog(0007, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0050) = 1000: AL_Error_Catalog(0007, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0051) = 1000: AL_Error_Catalog(0007, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0052) = 1000: AL_Error_Catalog(0007, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0053) = 1000: AL_Error_Catalog(0007, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0054) = 1000: AL_Error_Catalog(0007, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0055) = 1000: AL_Error_Catalog(0007, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0056) = 1000: AL_Error_Catalog(0007, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0057) = 1000: AL_Error_Catalog(0007, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0058) = 1000: AL_Error_Catalog(0007, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0059) = 1000: AL_Error_Catalog(0007, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0060) = 1000: AL_Error_Catalog(0007, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0061) = 1000: AL_Error_Catalog(0007, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0062) = 1000: AL_Error_Catalog(0007, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0063) = 1000: AL_Error_Catalog(0007, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0064) = 1000: AL_Error_Catalog(0007, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0065) = 1000: AL_Error_Catalog(0007, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0066) = 1000: AL_Error_Catalog(0007, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0067) = 1000: AL_Error_Catalog(0007, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0068) = 1000: AL_Error_Catalog(0007, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0069) = 1000: AL_Error_Catalog(0007, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0070) = 1000: AL_Error_Catalog(0007, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0071) = 1000: AL_Error_Catalog(0007, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0072) = 1000: AL_Error_Catalog(0007, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0073) = 1000: AL_Error_Catalog(0007, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0074) = 1000: AL_Error_Catalog(0007, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0075) = 1000: AL_Error_Catalog(0007, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0076) = 1000: AL_Error_Catalog(0007, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0077) = 1000: AL_Error_Catalog(0007, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0078) = 1000: AL_Error_Catalog(0007, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0079) = 1000: AL_Error_Catalog(0007, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0080) = 1000: AL_Error_Catalog(0007, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0081) = 1000: AL_Error_Catalog(0007, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0082) = 1000: AL_Error_Catalog(0007, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0083) = 1000: AL_Error_Catalog(0007, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0084) = 1000: AL_Error_Catalog(0007, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0085) = 1000: AL_Error_Catalog(0007, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0086) = 1000: AL_Error_Catalog(0007, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0087) = 1000: AL_Error_Catalog(0007, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0088) = 1000: AL_Error_Catalog(0007, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0089) = 1000: AL_Error_Catalog(0007, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0090) = 1000: AL_Error_Catalog(0007, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0091) = 1000: AL_Error_Catalog(0007, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0092) = 1000: AL_Error_Catalog(0007, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0093) = 1000: AL_Error_Catalog(0007, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0094) = 1000: AL_Error_Catalog(0007, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0095) = 1000: AL_Error_Catalog(0007, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0096) = 1000: AL_Error_Catalog(0007, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0097) = 1000: AL_Error_Catalog(0007, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0098) = 1000: AL_Error_Catalog(0007, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0099) = 1000: AL_Error_Catalog(0007, 0001, 0099) = "PLACEHOLDER"
    IS_ERROR = True
    ShowError = True
    PrintError = True
    Initialized = True
    End If

End Sub