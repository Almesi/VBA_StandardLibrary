Private AL_ErrorCatalog(7, 1, 99) As Variant
Public Const IS_ERROR As Boolean = True 

' Used to Compares to Objects by defined operand
Private Function AL_Error_ComObj(MainValue As Object, CompareValue As Object, CompareType As String) As Boolean

    Select Case CompareType
        Case "Is", "IS", "iS", "is", "="
            If MainValue Is CompareValue Then
                AL_Error_ComObj AL_Error_System, 0006, MainValue, CompareValue
                AL_Error_ComObj = AL_IS_ERROR
            End If
        Case "not", "noT", "nOt", "nOT", "Not", "NoT", "NOt", "NOT", "<>"
            If Not MainValue Is CompareValue Then
                AL_Error_ComObj AL_Error_System, 0011, MainValue, CompareValue
                AL_Error_ComObj = AL_IS_ERROR
            End If
        Case Empty
            AL_Error_ComObj AL_Error_System, 0002, "CompareType"
            AL_Error_ComObj = AL_IS_ERROR
        Case Else
            AL_Error_ComObj AL_Error_System, 0001, "CompareType"
            AL_Error_ComObj = AL_IS_ERROR
    End Select

End Function

Public Function AL_Error_Component(VBProj As VBIDE.VBProject, ComponentExistence As Boolean, Optional ComponentName As String = Empty, Optional ComponentIndex As Long = Empty) As Boolean
    
    Dim VBComp As VBIDE.VBComponent

    ' Handle VBProj
    If VBProj Is Nothing Then
        AL_Error_ComVar AL_Error_System, 0003, "VBProj"
    End If
    ' Handle Component Index
    If ComponentName = Empty And ComponentIndex <> Empty Then
            If ComponentExistence = AL_EXIST Then
                If ComponentIndex =< VBProj.VBComponents.Count Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0007, ComponentIndex
                    AL_Error_Component = AL_IS_ERROR
                End If
            Else
                If ComponentIndex > VBProj.VBComponents.Count Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0008, ComponentIndex
                    AL_Error_Component = AL_IS_ERROR
                End If
            End If
        Next VBComp
    ' Handle Component Name
    ElseIf ComponentName <> Empty And ComponentIndex = Empty Then
        For Each VBComp In VbProj.VBComponents
            If ComponentExistence = AL_EXIST Then
                If VBComp.Name = ComponentName Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0007, ComponentName
                    AL_Error_Component = AL_IS_ERROR
                    Exit Function
                End If
            Else
                If VBComp.Name = ComponentName Then
                    AL_Error_Component = AL_NO_ERROR
                    Exit Function
                End If
            End If
        Next VBComp
        If ComponentExistence = AL_EXIST Then
            AL_Error_Component = AL_IS_ERROR
            Exit Function
        End If
        AL_Error_Show AL_Error_CategoryWorkbook, 0008, ComponentName
        AL_Error_Component = AL_IS_ERROR
    ' Handle wrong Function input
    ElseIf ComponentName = Empty And ComponentIndex = Empty Then
        AL_Error_ComVar AL_Error_System, 0012, "ComponentName", "ComponentIndex"
        AL_Error_Component = AL_IS_ERROR
    ' Handle wrong Function input
    Else
        AL_Error_ComVar AL_Error_System, 0013, "ComponentName", "ComponentIndex"
        AL_Error_Component = AL_IS_ERROR
    End If

End Function

' Used to Compares to Variant Values by defined operand
Private Function AL_Error_ComVar(MainValue As Variant, CompareValue As Variant, CompareType As String) As Boolean

    Select Case CompareType
        Case "="
            If MainValue = CompareValue Then
                AL_Error_ComVar AL_Error_System, 0006, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case ">"
            If MainValue > CompareValue Then
                AL_Error_ComVar AL_Error_System, 0007, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case "<"
            If MainValue < CompareValue Then
                AL_Error_ComVar AL_Error_System, 0008, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case ">="
            If MainValue >= CompareValue Then
                AL_Error_ComVar AL_Error_System, 0009, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case "=<"
            If MainValue =< CompareValue Then
                AL_Error_ComVar AL_Error_System, 0010, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case "<>"
            If MainValue <> CompareValue Then
                AL_Error_ComVar AL_Error_System, 0011, MainValue, CompareValue
                AL_Error_ComVar = AL_IS_ERROR
            End If
        Case Empty
            AL_Error_ComVar AL_Error_System, 0002, "CompareType"
            AL_Error_ComVar = AL_IS_ERROR
        Case Else
            AL_Error_ComVar AL_Error_System, 0001, "CompareType"
            AL_Error_ComVar = AL_IS_ERROR
    End Select

End Function

Private Function AL_Error_Get(ErrorCategory As Long, ErrorIndex As Long) As Variant()
    
    Dim Temp(2) As Variant
    Select Case ErrorCategory
        Case Error_System:    Temp(Error_Severity) = AL_Error_Category_System(ErrorIndex, Error_Severity):    Temp(Error_Index) = AL_Error_Category_System(ErrorIndex, Error_Index):    Temp(Error_Message) = AL_Error_Category_System(ErrorIndex, Error_Message)
        Case Error_Compiler:  Temp(Error_Severity) = AL_Error_Category_Compiler(ErrorIndex, Error_Severity):  Temp(Error_Index) = AL_Error_Category_Compiler(ErrorIndex, Error_Index):  Temp(Error_Message) = AL_Error_Category_Compiler(ErrorIndex, Error_Message)
        Case Error_Linker:    Temp(Error_Severity) = AL_Error_Category_Linker(ErrorIndex, Error_Severity):    Temp(Error_Index) = AL_Error_Category_Linker(ErrorIndex, Error_Index):    Temp(Error_Message) = AL_Error_Category_Linker(ErrorIndex, Error_Message)
        Case Error_Module:    Temp(Error_Severity) = AL_Error_Category_Module(ErrorIndex, Error_Severity):    Temp(Error_Index) = AL_Error_Category_Module(ErrorIndex, Error_Index):    Temp(Error_Message) = AL_Error_Category_Module(ErrorIndex, Error_Message)
        Case Error_Class:     Temp(Error_Severity) = AL_Error_Category_Class(ErrorIndex, Error_Severity):     Temp(Error_Index) = AL_Error_Category_Class(ErrorIndex, Error_Index):     Temp(Error_Message) = AL_Error_Category_Class(ErrorIndex, Error_Message)
        Case Error_Form:      Temp(Error_Severity) = AL_Error_Category_Form(ErrorIndex, Error_Severity):      Temp(Error_Index) = AL_Error_Category_Form(ErrorIndex, Error_Index):      Temp(Error_Message) = AL_Error_Category_Form(ErrorIndex, Error_Message)
        Case Error_Workbook:  Temp(Error_Severity) = AL_Error_Category_Workbook(ErrorIndex, Error_Severity):  Temp(Error_Index) = AL_Error_Category_Workbook(ErrorIndex, Error_Index):  Temp(Error_Message) = AL_Error_Category_Workbook(ErrorIndex, Error_Message)
        Case Error_WorkSheet: Temp(Error_Severity) = AL_Error_Category_WorkSheet(ErrorIndex, Error_Severity): Temp(Error_Index) = AL_Error_Category_WorkSheet(ErrorIndex, Error_Index): Temp(Error_Message) = AL_Error_Category_WorkSheet(ErrorIndex, Error_Message)
        Case Else
                              Temp(Error_Severity) = 0: Temp(Error_Index) = 0: Temp(Error_Message) = "ErrorCategory not defined"
    End Select
    AL_Error_Get = Temp

End Function

Private Function AL_Error_GetCategory(ErrorCategory As Long) As String

    Select Case ErrorCategory
        Case AL_Error_System:    AL_Error_GetCategory = "System"
        Case AL_Error_Compiler:  AL_Error_GetCategory = "Workbook"
        Case AL_Error_Linker:    AL_Error_GetCategory = "Worksheet"
        Case AL_Error_Module:    AL_Error_GetCategory = "Linker"
        Case AL_Error_Class:     AL_Error_GetCategory = "Compiler"
        Case AL_Error_Form:      AL_Error_GetCategory = "Module"
        Case AL_Error_Workbook:  AL_Error_GetCategory = "Class"
        Case AL_Error_WorkSheet: AL_Error_GetCategory = "Userform"
        Case Else
                              AL_Error_GetCategory = "UNKNOWN"
    End Select

End Function

' Check if Variant has a defined Error
' Either Comparing to CompareObject, or Min-Max
Function AL_Error_Obj(MainObject As Object, Optional CompareObject As Object = Nothing, Optional CompareType As Object = Nothing) As Boolean

    ' Handle MainObject
    If MainObject = Nothing Then
        AL_Error_ComVar AL_Error_System, 0003, "MainObject"
        AL_Error_Obj = AL_IS_ERROR
    End If
    ' Handle CompareObject
    If CompareObject <> Nothing And (MinValue = Nothing And MaxValue = Nothing) Then
        If AL_Error_ComVar(MainObject, CompareObject, CompareType) = AL_IS_ERROR Then
            AL_Error_Obj = AL_IS_ERROR
        End If
    End If
    ' Handle MinValue
    If MinValue <> Nothing And CompareObject = Nothing Then
        If AL_Error_ComVar(MainObject, MinValue, CompareType) = AL_IS_ERROR Then
            AL_Error_Obj = AL_IS_ERROR
        End If
    End If
    ' Handle MaxValue
    If MaxValue <> Nothing And CompareObject = Nothing Then
        If AL_Error_ComVar(MainObject, MaxValue, CompareType) = AL_IS_ERROR Then
            AL_Error_Obj = AL_IS_ERROR
        End If
    End If
    ' Handle in bounds
    If MinValue <> Nothing And MaxValue <> Nothing And CompareObject = Nothing Then
        If AL_Error_ComVar(MainObject, MinValue, ">") = AL_IS_ERROR Or  AL_Error_ComVar(MainObject, MinValue, "<") = AL_IS_ERROR Then
            AL_Error_Obj = AL_IS_ERROR
        End If
    End If

End Function

' Print Error to Immediate
Public Sub AL_Error_Print(ErrorCategory As Long, ErrorIndex As Integer, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR)

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant

    Set ErrorArray = AL_Error_Get(ErrorCategory, ErrorIndex)
    ErrorMessage = "( " & "Category: " & AL_Error_GetCategory(ErrorCategory) & " ) | " & _
                   "( " & "Severity: " & ErrorArray(AL_Error_Severity)       & " ) | " & _
                   "( " & "Index: "    & ErrorArray(AL_Error_Index)          & " ) | " & _
                   "( " & "Message: "  & ErrorArray(AL_Error_Message)        & " )"
    If ErrorValue1 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value1: " & ErrorValue1 & " )": End If
    If ErrorValue2 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value2: " & ErrorValue2 & " )": End If
    If ErrorValue3 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value3: " & ErrorValue3 & " )": End If
    If ErrorValue4 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value4: " & ErrorValue4 & " )": End If
    Debug.Print ErrorMessage
    If ErrorArray(AL_Error_Severity) >= AL_SEVERITY_BREAK Then
        End
    End If

End Sub

Public Function AL_Error_Sheet(WB As Workbook, SheetExistence As Boolean, Optional SheetName As String = Empty, Optional SheetIndex As Long = Empty) As Boolean

    ' Handle WB
    If WB Is Nothing Then
        AL_Error_ComVar AL_Error_System, 0003, "WB"
    End If
    ' Handle Sheet Index
    If SheetName = Empty And SheetIndex <> Empty Then
            If SheetExistence = AL_EXIST Then
                If SheetIndex =< WB.VBSheets.Count Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0000, SheetIndex
                    AL_Error_Sheet = AL_IS_ERROR
                End If
            Else
                If SheetIndex > WB.VBSheets.Count Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0001, SheetIndex
                    AL_Error_Sheet = AL_IS_ERROR
                End If
            End If
        Next VBComp
    ' Handle Sheet Name
    ElseIf SheetName <> Empty And SheetIndex = Empty Then
        For Each VBComp In WB.VBSheets
            If SheetExistence = AL_EXIST Then
                If VBComp.Name = SheetName Then
                    AL_Error_Show AL_Error_CategoryWorkbook, 0007, SheetName
                    AL_Error_Sheet = AL_IS_ERROR
                    Exit Function
                End If
            Else
                If VBComp.Name = SheetName Then
                    AL_Error_Sheet = AL_NO_ERROR
                    Exit Function
                End If
            End If
        Next VBComp
        If SheetExistence = AL_EXIST Then
            AL_Error_Sheet = True
            Exit Function
        End If
        AL_Error_Show AL_Error_CategoryWorkbook, 0008, SheetName
        AL_Error_Sheet = AL_IS_ERROR
    ' Handle wrong Function input
    ElseIf SheetName = Empty And SheetIndex = Empty Then
        AL_Error_ComVar AL_Error_System, 0012, "SheetName", "SheetIndex"
        AL_Error_Sheet = AL_IS_ERROR
    ' Handle wrong Function input
    Else
        AL_Error_ComVar AL_Error_System, 0013, "SheetName", "SheetIndex"
        AL_Error_Sheet = AL_IS_ERROR
    End If

End Function

' Print Error as MessageBox
Public Sub AL_Error_Show(ErrorCategory As Long, ErrorIndex As Integer, Optional ErrorValue1 As Variant = AL_EMPTY_ERROR, Optional ErrorValue2 As Variant = AL_EMPTY_ERROR, Optional ErrorValue3 As Variant = AL_EMPTY_ERROR, Optional ErrorValue4 As Variant = AL_EMPTY_ERROR)

    Dim ErrorMessage As String
    Dim ErrorArray() As Variant

    AL_Error_Initialize
    Set ErrorArray = AL_Error_Get(ErrorCategory, ErrorIndex)
    ErrorMessage = "( " & "Category: " & AL_Error_GetCategory(ErrorCategory) & " ) | " & _
                   "( " & "Severity: " & ErrorArray(AL_Error_Severity)       & " ) | " & _
                   "( " & "Index: "    & ErrorArray(AL_Error_Index)          & " ) | " & _
                   "( " & "Message: "  & ErrorArray(AL_Error_Message)        & " )"
    If ErrorValue1 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value1: " & ErrorValue1 & " )": End If
    If ErrorValue2 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value2: " & ErrorValue2 & " )": End If
    If ErrorValue3 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value3: " & ErrorValue3 & " )": End If
    If ErrorValue4 <> AL_EMPTY_ERROR Then: ErrorMessage = ErrorMessage & " | ( " & "Value4: " & ErrorValue4 & " )": End If
    MsgBox(ErrorMessage, vbExclamation, "ERROR")
    AL_Error_Print ErrorCategory, ErrorIndex, ErrorValue1, ErrorValue2, ErrorValue3, ErrorValue4
    If ErrorArray(AL_Error_Severity) >= AL_SEVERITY_BREAK Then
        End
    End If

End Sub

' Check if Variant has a defined Error
' Either Comparing to CompareValue, or Min-Max
Function AL_Error_Var(MainValue As Variant, Optional CompareValue As Variant = Empty, Optional CompareType As Variant = Empty, Optional MinValue As Variant = Empty, Optional MaxValue As Variant = Empty) As Boolean

    ' Handle MainValue
    If MainValue = Empty Then
        AL_Error_ComVar AL_Error_System, 0002, "MainValue"
        AL_Error_Var = AL_IS_ERROR
    End If
    ' Handle CompareValue
    If CompareValue <> Empty And (MinValue = Empty And MaxValue = Empty) Then
        If AL_Error_ComVar(MainValue, CompareValue, CompareType) = AL_IS_ERROR Then
            AL_Error_Var = AL_IS_ERROR
        End If
    End If
    ' Handle MinValue
    If MinValue <> Empty And CompareValue = Empty Then
        If AL_Error_ComVar(MainValue, MinValue, CompareType) = AL_IS_ERROR Then
            AL_Error_Var = AL_IS_ERROR
        End If
    End If
    ' Handle MaxValue
    If MaxValue <> Empty And CompareValue = Empty Then
        If AL_Error_ComVar(MainValue, MaxValue, CompareType) = AL_IS_ERROR Then
            AL_Error_Var = AL_IS_ERROR
        End If
    End If
    ' Handle in bounds
    If MinValue <> Empty And MaxValue <> Empty And CompareValue = Empty Then
        If AL_Error_ComVar(MainValue, MinValue, ">") = AL_IS_ERROR Or  AL_Error_ComVar(MainValue, MinValue, "<") = AL_IS_ERROR Then
            AL_Error_Var = AL_IS_ERROR
        End If
    End If

End Function

' Runs once to Initialize all Errormessages
Private Sub AL_Error_Initialize()
    
    If AL_Error_Initialized = False Then
    ' System-Errors
        AL_Error_Catalog(0000, 0000, 0000) = 0000: AL_Error_Catalog(0000, 0001, 0000) = "ErrorCategory doesnt exist"
        AL_Error_Catalog(0000, 0000, 0001) = 0000: AL_Error_Catalog(0000, 0001, 0001) = "Value isnt available"
        AL_Error_Catalog(0000, 0000, 0002) = 0000: AL_Error_Catalog(0000, 0001, 0002) = "Value is Empty"
        AL_Error_Catalog(0000, 0000, 0003) = 0000: AL_Error_Catalog(0000, 0001, 0003) = "Value is Nothing"
        AL_Error_Catalog(0000, 0000, 0004) = 0000: AL_Error_Catalog(0000, 0001, 0004) = "Value Overflow"
        AL_Error_Catalog(0000, 0000, 0005) = 0000: AL_Error_Catalog(0000, 0001, 0005) = "Value Underflow"
        AL_Error_Catalog(0000, 0000, 0006) = 0000: AL_Error_Catalog(0000, 0001, 0006) = "Value1 doesnt equal Value2"
        AL_Error_Catalog(0000, 0000, 0007) = 0000: AL_Error_Catalog(0000, 0001, 0007) = "Value1 is smaller than or equal to Value2"
        AL_Error_Catalog(0000, 0000, 0008) = 0000: AL_Error_Catalog(0000, 0001, 0008) = "Value1 is bigger than or equal to Value2"
        AL_Error_Catalog(0000, 0000, 0009) = 0000: AL_Error_Catalog(0000, 0001, 0009) = "Value1 is smaller Value2"
        AL_Error_Catalog(0000, 0000, 0010) = 0000: AL_Error_Catalog(0000, 0001, 0010) = "Value1 is bigger Value2"
        AL_Error_Catalog(0000, 0000, 0011) = 0000: AL_Error_Catalog(0000, 0001, 0011) = "Value1 is Value2"
        AL_Error_Catalog(0000, 0000, 0012) = 0000: AL_Error_Catalog(0000, 0001, 0012) = "Several Values are Empty"
        AL_Error_Catalog(0000, 0000, 0013) = 0000: AL_Error_Catalog(0000, 0001, 0013) = "To many Values arent Empty"
        AL_Error_Catalog(0000, 0000, 0014) = 0000: AL_Error_Catalog(0000, 0001, 0014) = "Value Is Something"
        AL_Error_Catalog(0000, 0000, 0015) = 0000: AL_Error_Catalog(0000, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0016) = 0000: AL_Error_Catalog(0000, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0017) = 0000: AL_Error_Catalog(0000, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0018) = 0000: AL_Error_Catalog(0000, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0019) = 0000: AL_Error_Catalog(0000, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0020) = 0000: AL_Error_Catalog(0000, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0021) = 0000: AL_Error_Catalog(0000, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0022) = 0000: AL_Error_Catalog(0000, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0023) = 0000: AL_Error_Catalog(0000, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0024) = 0000: AL_Error_Catalog(0000, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0025) = 0000: AL_Error_Catalog(0000, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0026) = 0000: AL_Error_Catalog(0000, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0027) = 0000: AL_Error_Catalog(0000, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0028) = 0000: AL_Error_Catalog(0000, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0029) = 0000: AL_Error_Catalog(0000, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0030) = 0000: AL_Error_Catalog(0000, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0031) = 0000: AL_Error_Catalog(0000, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0032) = 0000: AL_Error_Catalog(0000, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0033) = 0000: AL_Error_Catalog(0000, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0034) = 0000: AL_Error_Catalog(0000, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0035) = 0000: AL_Error_Catalog(0000, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0036) = 0000: AL_Error_Catalog(0000, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0037) = 0000: AL_Error_Catalog(0000, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0038) = 0000: AL_Error_Catalog(0000, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0039) = 0000: AL_Error_Catalog(0000, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0040) = 0000: AL_Error_Catalog(0000, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0041) = 0000: AL_Error_Catalog(0000, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0042) = 0000: AL_Error_Catalog(0000, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0043) = 0000: AL_Error_Catalog(0000, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0044) = 0000: AL_Error_Catalog(0000, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0045) = 0000: AL_Error_Catalog(0000, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0046) = 0000: AL_Error_Catalog(0000, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0047) = 0000: AL_Error_Catalog(0000, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0048) = 0000: AL_Error_Catalog(0000, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0049) = 0000: AL_Error_Catalog(0000, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0050) = 0000: AL_Error_Catalog(0000, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0051) = 0000: AL_Error_Catalog(0000, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0052) = 0000: AL_Error_Catalog(0000, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0053) = 0000: AL_Error_Catalog(0000, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0054) = 0000: AL_Error_Catalog(0000, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0055) = 0000: AL_Error_Catalog(0000, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0056) = 0000: AL_Error_Catalog(0000, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0057) = 0000: AL_Error_Catalog(0000, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0058) = 0000: AL_Error_Catalog(0000, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0059) = 0000: AL_Error_Catalog(0000, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0060) = 0000: AL_Error_Catalog(0000, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0061) = 0000: AL_Error_Catalog(0000, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0062) = 0000: AL_Error_Catalog(0000, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0063) = 0000: AL_Error_Catalog(0000, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0064) = 0000: AL_Error_Catalog(0000, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0065) = 0000: AL_Error_Catalog(0000, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0066) = 0000: AL_Error_Catalog(0000, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0067) = 0000: AL_Error_Catalog(0000, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0068) = 0000: AL_Error_Catalog(0000, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0069) = 0000: AL_Error_Catalog(0000, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0070) = 0000: AL_Error_Catalog(0000, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0071) = 0000: AL_Error_Catalog(0000, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0072) = 0000: AL_Error_Catalog(0000, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0073) = 0000: AL_Error_Catalog(0000, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0074) = 0000: AL_Error_Catalog(0000, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0075) = 0000: AL_Error_Catalog(0000, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0076) = 0000: AL_Error_Catalog(0000, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0077) = 0000: AL_Error_Catalog(0000, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0078) = 0000: AL_Error_Catalog(0000, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0079) = 0000: AL_Error_Catalog(0000, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0080) = 0000: AL_Error_Catalog(0000, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0081) = 0000: AL_Error_Catalog(0000, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0082) = 0000: AL_Error_Catalog(0000, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0083) = 0000: AL_Error_Catalog(0000, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0084) = 0000: AL_Error_Catalog(0000, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0085) = 0000: AL_Error_Catalog(0000, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0086) = 0000: AL_Error_Catalog(0000, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0087) = 0000: AL_Error_Catalog(0000, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0088) = 0000: AL_Error_Catalog(0000, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0089) = 0000: AL_Error_Catalog(0000, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0090) = 0000: AL_Error_Catalog(0000, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0091) = 0000: AL_Error_Catalog(0000, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0092) = 0000: AL_Error_Catalog(0000, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0093) = 0000: AL_Error_Catalog(0000, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0094) = 0000: AL_Error_Catalog(0000, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0095) = 0000: AL_Error_Catalog(0000, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0096) = 0000: AL_Error_Catalog(0000, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0097) = 0000: AL_Error_Catalog(0000, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0098) = 0000: AL_Error_Catalog(0000, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0000, 0000, 0099) = 0000: AL_Error_Catalog(0000, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0000) = 0000: AL_Error_Catalog(0001, 0001, 0000) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0001) = 0000: AL_Error_Catalog(0001, 0001, 0001) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0002) = 0000: AL_Error_Catalog(0001, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0003) = 0000: AL_Error_Catalog(0001, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0004) = 0000: AL_Error_Catalog(0001, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0005) = 0000: AL_Error_Catalog(0001, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0006) = 0000: AL_Error_Catalog(0001, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0007) = 0000: AL_Error_Catalog(0001, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0008) = 0000: AL_Error_Catalog(0001, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0009) = 0000: AL_Error_Catalog(0001, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0010) = 0000: AL_Error_Catalog(0001, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0011) = 0000: AL_Error_Catalog(0001, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0012) = 0000: AL_Error_Catalog(0001, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0013) = 0000: AL_Error_Catalog(0001, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0014) = 0000: AL_Error_Catalog(0001, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0015) = 0000: AL_Error_Catalog(0001, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0016) = 0000: AL_Error_Catalog(0001, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0017) = 0000: AL_Error_Catalog(0001, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0018) = 0000: AL_Error_Catalog(0001, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0019) = 0000: AL_Error_Catalog(0001, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0020) = 0000: AL_Error_Catalog(0001, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0021) = 0000: AL_Error_Catalog(0001, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0022) = 0000: AL_Error_Catalog(0001, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0023) = 0000: AL_Error_Catalog(0001, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0024) = 0000: AL_Error_Catalog(0001, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0025) = 0000: AL_Error_Catalog(0001, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0026) = 0000: AL_Error_Catalog(0001, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0027) = 0000: AL_Error_Catalog(0001, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0028) = 0000: AL_Error_Catalog(0001, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0029) = 0000: AL_Error_Catalog(0001, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0030) = 0000: AL_Error_Catalog(0001, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0031) = 0000: AL_Error_Catalog(0001, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0032) = 0000: AL_Error_Catalog(0001, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0033) = 0000: AL_Error_Catalog(0001, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0034) = 0000: AL_Error_Catalog(0001, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0035) = 0000: AL_Error_Catalog(0001, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0036) = 0000: AL_Error_Catalog(0001, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0037) = 0000: AL_Error_Catalog(0001, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0038) = 0000: AL_Error_Catalog(0001, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0039) = 0000: AL_Error_Catalog(0001, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0040) = 0000: AL_Error_Catalog(0001, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0041) = 0000: AL_Error_Catalog(0001, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0042) = 0000: AL_Error_Catalog(0001, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0043) = 0000: AL_Error_Catalog(0001, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0044) = 0000: AL_Error_Catalog(0001, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0045) = 0000: AL_Error_Catalog(0001, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0046) = 0000: AL_Error_Catalog(0001, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0047) = 0000: AL_Error_Catalog(0001, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0048) = 0000: AL_Error_Catalog(0001, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0049) = 0000: AL_Error_Catalog(0001, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0050) = 0000: AL_Error_Catalog(0001, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0051) = 0000: AL_Error_Catalog(0001, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0052) = 0000: AL_Error_Catalog(0001, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0053) = 0000: AL_Error_Catalog(0001, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0054) = 0000: AL_Error_Catalog(0001, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0055) = 0000: AL_Error_Catalog(0001, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0056) = 0000: AL_Error_Catalog(0001, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0057) = 0000: AL_Error_Catalog(0001, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0058) = 0000: AL_Error_Catalog(0001, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0059) = 0000: AL_Error_Catalog(0001, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0060) = 0000: AL_Error_Catalog(0001, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0061) = 0000: AL_Error_Catalog(0001, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0062) = 0000: AL_Error_Catalog(0001, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0063) = 0000: AL_Error_Catalog(0001, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0064) = 0000: AL_Error_Catalog(0001, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0065) = 0000: AL_Error_Catalog(0001, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0066) = 0000: AL_Error_Catalog(0001, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0067) = 0000: AL_Error_Catalog(0001, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0068) = 0000: AL_Error_Catalog(0001, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0069) = 0000: AL_Error_Catalog(0001, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0070) = 0000: AL_Error_Catalog(0001, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0071) = 0000: AL_Error_Catalog(0001, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0072) = 0000: AL_Error_Catalog(0001, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0073) = 0000: AL_Error_Catalog(0001, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0074) = 0000: AL_Error_Catalog(0001, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0075) = 0000: AL_Error_Catalog(0001, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0076) = 0000: AL_Error_Catalog(0001, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0077) = 0000: AL_Error_Catalog(0001, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0078) = 0000: AL_Error_Catalog(0001, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0079) = 0000: AL_Error_Catalog(0001, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0080) = 0000: AL_Error_Catalog(0001, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0081) = 0000: AL_Error_Catalog(0001, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0082) = 0000: AL_Error_Catalog(0001, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0083) = 0000: AL_Error_Catalog(0001, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0084) = 0000: AL_Error_Catalog(0001, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0085) = 0000: AL_Error_Catalog(0001, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0086) = 0000: AL_Error_Catalog(0001, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0087) = 0000: AL_Error_Catalog(0001, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0088) = 0000: AL_Error_Catalog(0001, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0089) = 0000: AL_Error_Catalog(0001, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0090) = 0000: AL_Error_Catalog(0001, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0091) = 0000: AL_Error_Catalog(0001, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0092) = 0000: AL_Error_Catalog(0001, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0093) = 0000: AL_Error_Catalog(0001, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0094) = 0000: AL_Error_Catalog(0001, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0095) = 0000: AL_Error_Catalog(0001, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0096) = 0000: AL_Error_Catalog(0001, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0097) = 0000: AL_Error_Catalog(0001, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0098) = 0000: AL_Error_Catalog(0001, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0001, 0000, 0099) = 0000: AL_Error_Catalog(0001, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0000) = 0000: AL_Error_Catalog(0002, 0001, 0000) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0001) = 0000: AL_Error_Catalog(0002, 0001, 0001) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0002) = 0000: AL_Error_Catalog(0002, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0003) = 0000: AL_Error_Catalog(0002, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0004) = 0000: AL_Error_Catalog(0002, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0005) = 0000: AL_Error_Catalog(0002, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0006) = 0000: AL_Error_Catalog(0002, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0007) = 0000: AL_Error_Catalog(0002, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0008) = 0000: AL_Error_Catalog(0002, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0009) = 0000: AL_Error_Catalog(0002, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0010) = 0000: AL_Error_Catalog(0002, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0011) = 0000: AL_Error_Catalog(0002, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0012) = 0000: AL_Error_Catalog(0002, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0013) = 0000: AL_Error_Catalog(0002, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0014) = 0000: AL_Error_Catalog(0002, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0015) = 0000: AL_Error_Catalog(0002, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0016) = 0000: AL_Error_Catalog(0002, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0017) = 0000: AL_Error_Catalog(0002, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0018) = 0000: AL_Error_Catalog(0002, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0019) = 0000: AL_Error_Catalog(0002, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0020) = 0000: AL_Error_Catalog(0002, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0021) = 0000: AL_Error_Catalog(0002, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0022) = 0000: AL_Error_Catalog(0002, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0023) = 0000: AL_Error_Catalog(0002, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0024) = 0000: AL_Error_Catalog(0002, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0025) = 0000: AL_Error_Catalog(0002, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0026) = 0000: AL_Error_Catalog(0002, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0027) = 0000: AL_Error_Catalog(0002, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0028) = 0000: AL_Error_Catalog(0002, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0029) = 0000: AL_Error_Catalog(0002, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0030) = 0000: AL_Error_Catalog(0002, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0031) = 0000: AL_Error_Catalog(0002, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0032) = 0000: AL_Error_Catalog(0002, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0033) = 0000: AL_Error_Catalog(0002, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0034) = 0000: AL_Error_Catalog(0002, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0035) = 0000: AL_Error_Catalog(0002, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0036) = 0000: AL_Error_Catalog(0002, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0037) = 0000: AL_Error_Catalog(0002, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0038) = 0000: AL_Error_Catalog(0002, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0039) = 0000: AL_Error_Catalog(0002, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0040) = 0000: AL_Error_Catalog(0002, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0041) = 0000: AL_Error_Catalog(0002, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0042) = 0000: AL_Error_Catalog(0002, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0043) = 0000: AL_Error_Catalog(0002, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0044) = 0000: AL_Error_Catalog(0002, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0045) = 0000: AL_Error_Catalog(0002, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0046) = 0000: AL_Error_Catalog(0002, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0047) = 0000: AL_Error_Catalog(0002, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0048) = 0000: AL_Error_Catalog(0002, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0049) = 0000: AL_Error_Catalog(0002, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0050) = 0000: AL_Error_Catalog(0002, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0051) = 0000: AL_Error_Catalog(0002, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0052) = 0000: AL_Error_Catalog(0002, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0053) = 0000: AL_Error_Catalog(0002, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0054) = 0000: AL_Error_Catalog(0002, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0055) = 0000: AL_Error_Catalog(0002, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0056) = 0000: AL_Error_Catalog(0002, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0057) = 0000: AL_Error_Catalog(0002, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0058) = 0000: AL_Error_Catalog(0002, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0059) = 0000: AL_Error_Catalog(0002, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0060) = 0000: AL_Error_Catalog(0002, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0061) = 0000: AL_Error_Catalog(0002, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0062) = 0000: AL_Error_Catalog(0002, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0063) = 0000: AL_Error_Catalog(0002, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0064) = 0000: AL_Error_Catalog(0002, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0065) = 0000: AL_Error_Catalog(0002, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0066) = 0000: AL_Error_Catalog(0002, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0067) = 0000: AL_Error_Catalog(0002, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0068) = 0000: AL_Error_Catalog(0002, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0069) = 0000: AL_Error_Catalog(0002, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0070) = 0000: AL_Error_Catalog(0002, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0071) = 0000: AL_Error_Catalog(0002, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0072) = 0000: AL_Error_Catalog(0002, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0073) = 0000: AL_Error_Catalog(0002, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0074) = 0000: AL_Error_Catalog(0002, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0075) = 0000: AL_Error_Catalog(0002, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0076) = 0000: AL_Error_Catalog(0002, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0077) = 0000: AL_Error_Catalog(0002, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0078) = 0000: AL_Error_Catalog(0002, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0079) = 0000: AL_Error_Catalog(0002, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0080) = 0000: AL_Error_Catalog(0002, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0081) = 0000: AL_Error_Catalog(0002, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0082) = 0000: AL_Error_Catalog(0002, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0083) = 0000: AL_Error_Catalog(0002, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0084) = 0000: AL_Error_Catalog(0002, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0085) = 0000: AL_Error_Catalog(0002, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0086) = 0000: AL_Error_Catalog(0002, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0087) = 0000: AL_Error_Catalog(0002, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0088) = 0000: AL_Error_Catalog(0002, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0089) = 0000: AL_Error_Catalog(0002, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0090) = 0000: AL_Error_Catalog(0002, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0091) = 0000: AL_Error_Catalog(0002, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0092) = 0000: AL_Error_Catalog(0002, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0093) = 0000: AL_Error_Catalog(0002, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0094) = 0000: AL_Error_Catalog(0002, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0095) = 0000: AL_Error_Catalog(0002, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0096) = 0000: AL_Error_Catalog(0002, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0097) = 0000: AL_Error_Catalog(0002, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0098) = 0000: AL_Error_Catalog(0002, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0002, 0000, 0099) = 0000: AL_Error_Catalog(0002, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0000) = 0000: AL_Error_Catalog(0003, 0001, 0000) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0001) = 0000: AL_Error_Catalog(0003, 0001, 0001) = "Workbook doesnt exist"
        AL_Error_Catalog(0003, 0000, 0002) = 0000: AL_Error_Catalog(0003, 0001, 0002) = "Workbook already exists"
        AL_Error_Catalog(0003, 0000, 0003) = 0000: AL_Error_Catalog(0003, 0001, 0003) = "Dependency missing"
        AL_Error_Catalog(0003, 0000, 0004) = 0000: AL_Error_Catalog(0003, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0003, 0000, 0005) = 0000: AL_Error_Catalog(0003, 0001, 0005) = "Not available in Workbook"
        AL_Error_Catalog(0003, 0000, 0006) = 0000: AL_Error_Catalog(0003, 0001, 0006) = "Component doesnt exists"
        AL_Error_Catalog(0003, 0000, 0007) = 0000: AL_Error_Catalog(0003, 0001, 0007) = "Component already exists"
        AL_Error_Catalog(0003, 0000, 0008) = 0000: AL_Error_Catalog(0003, 0001, 0008) = "Couldnt Create Component"
        AL_Error_Catalog(0003, 0000, 0009) = 0000: AL_Error_Catalog(0003, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0010) = 0000: AL_Error_Catalog(0003, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0011) = 0000: AL_Error_Catalog(0003, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0012) = 0000: AL_Error_Catalog(0003, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0013) = 0000: AL_Error_Catalog(0003, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0014) = 0000: AL_Error_Catalog(0003, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0015) = 0000: AL_Error_Catalog(0003, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0016) = 0000: AL_Error_Catalog(0003, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0017) = 0000: AL_Error_Catalog(0003, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0018) = 0000: AL_Error_Catalog(0003, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0019) = 0000: AL_Error_Catalog(0003, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0020) = 0000: AL_Error_Catalog(0003, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0021) = 0000: AL_Error_Catalog(0003, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0022) = 0000: AL_Error_Catalog(0003, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0023) = 0000: AL_Error_Catalog(0003, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0024) = 0000: AL_Error_Catalog(0003, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0025) = 0000: AL_Error_Catalog(0003, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0026) = 0000: AL_Error_Catalog(0003, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0027) = 0000: AL_Error_Catalog(0003, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0028) = 0000: AL_Error_Catalog(0003, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0029) = 0000: AL_Error_Catalog(0003, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0030) = 0000: AL_Error_Catalog(0003, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0031) = 0000: AL_Error_Catalog(0003, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0032) = 0000: AL_Error_Catalog(0003, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0033) = 0000: AL_Error_Catalog(0003, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0034) = 0000: AL_Error_Catalog(0003, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0035) = 0000: AL_Error_Catalog(0003, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0036) = 0000: AL_Error_Catalog(0003, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0037) = 0000: AL_Error_Catalog(0003, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0038) = 0000: AL_Error_Catalog(0003, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0039) = 0000: AL_Error_Catalog(0003, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0040) = 0000: AL_Error_Catalog(0003, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0041) = 0000: AL_Error_Catalog(0003, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0042) = 0000: AL_Error_Catalog(0003, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0043) = 0000: AL_Error_Catalog(0003, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0044) = 0000: AL_Error_Catalog(0003, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0045) = 0000: AL_Error_Catalog(0003, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0046) = 0000: AL_Error_Catalog(0003, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0047) = 0000: AL_Error_Catalog(0003, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0048) = 0000: AL_Error_Catalog(0003, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0049) = 0000: AL_Error_Catalog(0003, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0050) = 0000: AL_Error_Catalog(0003, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0051) = 0000: AL_Error_Catalog(0003, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0052) = 0000: AL_Error_Catalog(0003, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0053) = 0000: AL_Error_Catalog(0003, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0054) = 0000: AL_Error_Catalog(0003, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0055) = 0000: AL_Error_Catalog(0003, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0056) = 0000: AL_Error_Catalog(0003, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0057) = 0000: AL_Error_Catalog(0003, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0058) = 0000: AL_Error_Catalog(0003, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0059) = 0000: AL_Error_Catalog(0003, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0060) = 0000: AL_Error_Catalog(0003, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0061) = 0000: AL_Error_Catalog(0003, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0062) = 0000: AL_Error_Catalog(0003, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0063) = 0000: AL_Error_Catalog(0003, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0064) = 0000: AL_Error_Catalog(0003, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0065) = 0000: AL_Error_Catalog(0003, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0066) = 0000: AL_Error_Catalog(0003, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0067) = 0000: AL_Error_Catalog(0003, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0068) = 0000: AL_Error_Catalog(0003, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0069) = 0000: AL_Error_Catalog(0003, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0070) = 0000: AL_Error_Catalog(0003, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0071) = 0000: AL_Error_Catalog(0003, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0072) = 0000: AL_Error_Catalog(0003, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0073) = 0000: AL_Error_Catalog(0003, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0074) = 0000: AL_Error_Catalog(0003, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0075) = 0000: AL_Error_Catalog(0003, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0076) = 0000: AL_Error_Catalog(0003, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0077) = 0000: AL_Error_Catalog(0003, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0078) = 0000: AL_Error_Catalog(0003, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0079) = 0000: AL_Error_Catalog(0003, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0080) = 0000: AL_Error_Catalog(0003, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0081) = 0000: AL_Error_Catalog(0003, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0082) = 0000: AL_Error_Catalog(0003, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0083) = 0000: AL_Error_Catalog(0003, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0084) = 0000: AL_Error_Catalog(0003, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0085) = 0000: AL_Error_Catalog(0003, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0086) = 0000: AL_Error_Catalog(0003, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0087) = 0000: AL_Error_Catalog(0003, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0088) = 0000: AL_Error_Catalog(0003, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0089) = 0000: AL_Error_Catalog(0003, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0090) = 0000: AL_Error_Catalog(0003, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0091) = 0000: AL_Error_Catalog(0003, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0092) = 0000: AL_Error_Catalog(0003, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0093) = 0000: AL_Error_Catalog(0003, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0094) = 0000: AL_Error_Catalog(0003, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0095) = 0000: AL_Error_Catalog(0003, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0096) = 0000: AL_Error_Catalog(0003, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0097) = 0000: AL_Error_Catalog(0003, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0098) = 0000: AL_Error_Catalog(0003, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0003, 0000, 0099) = 0000: AL_Error_Catalog(0003, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0000) = 0000: AL_Error_Catalog(0004, 0001, 0000) = "Worksheet already exists"
        AL_Error_Catalog(0004, 0000, 0001) = 0000: AL_Error_Catalog(0004, 0001, 0001) = "Worksheet doesnt exist"
        AL_Error_Catalog(0004, 0000, 0002) = 0000: AL_Error_Catalog(0004, 0001, 0002) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0003) = 0000: AL_Error_Catalog(0004, 0001, 0003) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0004) = 0000: AL_Error_Catalog(0004, 0001, 0004) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0005) = 0000: AL_Error_Catalog(0004, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0006) = 0000: AL_Error_Catalog(0004, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0007) = 0000: AL_Error_Catalog(0004, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0008) = 0000: AL_Error_Catalog(0004, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0009) = 0000: AL_Error_Catalog(0004, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0010) = 0000: AL_Error_Catalog(0004, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0011) = 0000: AL_Error_Catalog(0004, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0012) = 0000: AL_Error_Catalog(0004, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0013) = 0000: AL_Error_Catalog(0004, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0014) = 0000: AL_Error_Catalog(0004, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0015) = 0000: AL_Error_Catalog(0004, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0016) = 0000: AL_Error_Catalog(0004, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0017) = 0000: AL_Error_Catalog(0004, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0018) = 0000: AL_Error_Catalog(0004, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0019) = 0000: AL_Error_Catalog(0004, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0020) = 0000: AL_Error_Catalog(0004, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0021) = 0000: AL_Error_Catalog(0004, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0022) = 0000: AL_Error_Catalog(0004, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0023) = 0000: AL_Error_Catalog(0004, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0024) = 0000: AL_Error_Catalog(0004, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0025) = 0000: AL_Error_Catalog(0004, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0026) = 0000: AL_Error_Catalog(0004, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0027) = 0000: AL_Error_Catalog(0004, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0028) = 0000: AL_Error_Catalog(0004, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0029) = 0000: AL_Error_Catalog(0004, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0030) = 0000: AL_Error_Catalog(0004, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0031) = 0000: AL_Error_Catalog(0004, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0032) = 0000: AL_Error_Catalog(0004, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0033) = 0000: AL_Error_Catalog(0004, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0034) = 0000: AL_Error_Catalog(0004, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0035) = 0000: AL_Error_Catalog(0004, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0036) = 0000: AL_Error_Catalog(0004, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0037) = 0000: AL_Error_Catalog(0004, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0038) = 0000: AL_Error_Catalog(0004, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0039) = 0000: AL_Error_Catalog(0004, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0040) = 0000: AL_Error_Catalog(0004, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0041) = 0000: AL_Error_Catalog(0004, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0042) = 0000: AL_Error_Catalog(0004, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0043) = 0000: AL_Error_Catalog(0004, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0044) = 0000: AL_Error_Catalog(0004, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0045) = 0000: AL_Error_Catalog(0004, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0046) = 0000: AL_Error_Catalog(0004, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0047) = 0000: AL_Error_Catalog(0004, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0048) = 0000: AL_Error_Catalog(0004, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0049) = 0000: AL_Error_Catalog(0004, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0050) = 0000: AL_Error_Catalog(0004, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0051) = 0000: AL_Error_Catalog(0004, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0052) = 0000: AL_Error_Catalog(0004, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0053) = 0000: AL_Error_Catalog(0004, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0054) = 0000: AL_Error_Catalog(0004, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0055) = 0000: AL_Error_Catalog(0004, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0056) = 0000: AL_Error_Catalog(0004, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0057) = 0000: AL_Error_Catalog(0004, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0058) = 0000: AL_Error_Catalog(0004, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0059) = 0000: AL_Error_Catalog(0004, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0060) = 0000: AL_Error_Catalog(0004, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0061) = 0000: AL_Error_Catalog(0004, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0062) = 0000: AL_Error_Catalog(0004, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0063) = 0000: AL_Error_Catalog(0004, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0064) = 0000: AL_Error_Catalog(0004, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0065) = 0000: AL_Error_Catalog(0004, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0066) = 0000: AL_Error_Catalog(0004, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0067) = 0000: AL_Error_Catalog(0004, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0068) = 0000: AL_Error_Catalog(0004, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0069) = 0000: AL_Error_Catalog(0004, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0070) = 0000: AL_Error_Catalog(0004, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0071) = 0000: AL_Error_Catalog(0004, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0072) = 0000: AL_Error_Catalog(0004, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0073) = 0000: AL_Error_Catalog(0004, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0074) = 0000: AL_Error_Catalog(0004, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0075) = 0000: AL_Error_Catalog(0004, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0076) = 0000: AL_Error_Catalog(0004, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0077) = 0000: AL_Error_Catalog(0004, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0078) = 0000: AL_Error_Catalog(0004, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0079) = 0000: AL_Error_Catalog(0004, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0080) = 0000: AL_Error_Catalog(0004, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0081) = 0000: AL_Error_Catalog(0004, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0082) = 0000: AL_Error_Catalog(0004, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0083) = 0000: AL_Error_Catalog(0004, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0084) = 0000: AL_Error_Catalog(0004, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0085) = 0000: AL_Error_Catalog(0004, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0086) = 0000: AL_Error_Catalog(0004, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0087) = 0000: AL_Error_Catalog(0004, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0088) = 0000: AL_Error_Catalog(0004, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0089) = 0000: AL_Error_Catalog(0004, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0090) = 0000: AL_Error_Catalog(0004, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0091) = 0000: AL_Error_Catalog(0004, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0092) = 0000: AL_Error_Catalog(0004, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0093) = 0000: AL_Error_Catalog(0004, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0094) = 0000: AL_Error_Catalog(0004, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0095) = 0000: AL_Error_Catalog(0004, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0096) = 0000: AL_Error_Catalog(0004, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0097) = 0000: AL_Error_Catalog(0004, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0098) = 0000: AL_Error_Catalog(0004, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0004, 0000, 0099) = 0000: AL_Error_Catalog(0004, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0000) = 0000: AL_Error_Catalog(0005, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0005, 0000, 0001) = 0000: AL_Error_Catalog(0005, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0005, 0000, 0002) = 0000: AL_Error_Catalog(0005, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0005, 0000, 0003) = 0000: AL_Error_Catalog(0005, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0005, 0000, 0004) = 0000: AL_Error_Catalog(0005, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0005, 0000, 0005) = 0000: AL_Error_Catalog(0005, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0006) = 0000: AL_Error_Catalog(0005, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0007) = 0000: AL_Error_Catalog(0005, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0008) = 0000: AL_Error_Catalog(0005, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0009) = 0000: AL_Error_Catalog(0005, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0010) = 0000: AL_Error_Catalog(0005, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0011) = 0000: AL_Error_Catalog(0005, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0012) = 0000: AL_Error_Catalog(0005, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0013) = 0000: AL_Error_Catalog(0005, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0014) = 0000: AL_Error_Catalog(0005, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0015) = 0000: AL_Error_Catalog(0005, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0016) = 0000: AL_Error_Catalog(0005, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0017) = 0000: AL_Error_Catalog(0005, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0018) = 0000: AL_Error_Catalog(0005, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0019) = 0000: AL_Error_Catalog(0005, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0020) = 0000: AL_Error_Catalog(0005, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0021) = 0000: AL_Error_Catalog(0005, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0022) = 0000: AL_Error_Catalog(0005, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0023) = 0000: AL_Error_Catalog(0005, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0024) = 0000: AL_Error_Catalog(0005, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0025) = 0000: AL_Error_Catalog(0005, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0026) = 0000: AL_Error_Catalog(0005, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0027) = 0000: AL_Error_Catalog(0005, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0028) = 0000: AL_Error_Catalog(0005, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0029) = 0000: AL_Error_Catalog(0005, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0030) = 0000: AL_Error_Catalog(0005, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0031) = 0000: AL_Error_Catalog(0005, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0032) = 0000: AL_Error_Catalog(0005, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0033) = 0000: AL_Error_Catalog(0005, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0034) = 0000: AL_Error_Catalog(0005, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0035) = 0000: AL_Error_Catalog(0005, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0036) = 0000: AL_Error_Catalog(0005, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0037) = 0000: AL_Error_Catalog(0005, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0038) = 0000: AL_Error_Catalog(0005, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0039) = 0000: AL_Error_Catalog(0005, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0040) = 0000: AL_Error_Catalog(0005, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0041) = 0000: AL_Error_Catalog(0005, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0042) = 0000: AL_Error_Catalog(0005, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0043) = 0000: AL_Error_Catalog(0005, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0044) = 0000: AL_Error_Catalog(0005, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0045) = 0000: AL_Error_Catalog(0005, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0046) = 0000: AL_Error_Catalog(0005, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0047) = 0000: AL_Error_Catalog(0005, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0048) = 0000: AL_Error_Catalog(0005, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0049) = 0000: AL_Error_Catalog(0005, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0050) = 0000: AL_Error_Catalog(0005, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0051) = 0000: AL_Error_Catalog(0005, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0052) = 0000: AL_Error_Catalog(0005, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0053) = 0000: AL_Error_Catalog(0005, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0054) = 0000: AL_Error_Catalog(0005, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0055) = 0000: AL_Error_Catalog(0005, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0056) = 0000: AL_Error_Catalog(0005, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0057) = 0000: AL_Error_Catalog(0005, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0058) = 0000: AL_Error_Catalog(0005, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0059) = 0000: AL_Error_Catalog(0005, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0060) = 0000: AL_Error_Catalog(0005, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0061) = 0000: AL_Error_Catalog(0005, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0062) = 0000: AL_Error_Catalog(0005, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0063) = 0000: AL_Error_Catalog(0005, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0064) = 0000: AL_Error_Catalog(0005, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0065) = 0000: AL_Error_Catalog(0005, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0066) = 0000: AL_Error_Catalog(0005, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0067) = 0000: AL_Error_Catalog(0005, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0068) = 0000: AL_Error_Catalog(0005, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0069) = 0000: AL_Error_Catalog(0005, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0070) = 0000: AL_Error_Catalog(0005, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0071) = 0000: AL_Error_Catalog(0005, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0072) = 0000: AL_Error_Catalog(0005, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0073) = 0000: AL_Error_Catalog(0005, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0074) = 0000: AL_Error_Catalog(0005, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0075) = 0000: AL_Error_Catalog(0005, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0076) = 0000: AL_Error_Catalog(0005, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0077) = 0000: AL_Error_Catalog(0005, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0078) = 0000: AL_Error_Catalog(0005, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0079) = 0000: AL_Error_Catalog(0005, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0080) = 0000: AL_Error_Catalog(0005, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0081) = 0000: AL_Error_Catalog(0005, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0082) = 0000: AL_Error_Catalog(0005, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0083) = 0000: AL_Error_Catalog(0005, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0084) = 0000: AL_Error_Catalog(0005, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0085) = 0000: AL_Error_Catalog(0005, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0086) = 0000: AL_Error_Catalog(0005, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0087) = 0000: AL_Error_Catalog(0005, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0088) = 0000: AL_Error_Catalog(0005, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0089) = 0000: AL_Error_Catalog(0005, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0090) = 0000: AL_Error_Catalog(0005, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0091) = 0000: AL_Error_Catalog(0005, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0092) = 0000: AL_Error_Catalog(0005, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0093) = 0000: AL_Error_Catalog(0005, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0094) = 0000: AL_Error_Catalog(0005, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0095) = 0000: AL_Error_Catalog(0005, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0096) = 0000: AL_Error_Catalog(0005, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0097) = 0000: AL_Error_Catalog(0005, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0098) = 0000: AL_Error_Catalog(0005, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0005, 0000, 0099) = 0000: AL_Error_Catalog(0005, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0000) = 0000: AL_Error_Catalog(0006, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0006, 0000, 0001) = 0000: AL_Error_Catalog(0006, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0006, 0000, 0002) = 0000: AL_Error_Catalog(0006, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0006, 0000, 0003) = 0000: AL_Error_Catalog(0006, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0006, 0000, 0004) = 0000: AL_Error_Catalog(0006, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0006, 0000, 0005) = 0000: AL_Error_Catalog(0006, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0006) = 0000: AL_Error_Catalog(0006, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0007) = 0000: AL_Error_Catalog(0006, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0008) = 0000: AL_Error_Catalog(0006, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0009) = 0000: AL_Error_Catalog(0006, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0010) = 0000: AL_Error_Catalog(0006, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0011) = 0000: AL_Error_Catalog(0006, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0012) = 0000: AL_Error_Catalog(0006, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0013) = 0000: AL_Error_Catalog(0006, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0014) = 0000: AL_Error_Catalog(0006, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0015) = 0000: AL_Error_Catalog(0006, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0016) = 0000: AL_Error_Catalog(0006, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0017) = 0000: AL_Error_Catalog(0006, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0018) = 0000: AL_Error_Catalog(0006, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0019) = 0000: AL_Error_Catalog(0006, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0020) = 0000: AL_Error_Catalog(0006, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0021) = 0000: AL_Error_Catalog(0006, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0022) = 0000: AL_Error_Catalog(0006, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0023) = 0000: AL_Error_Catalog(0006, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0024) = 0000: AL_Error_Catalog(0006, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0025) = 0000: AL_Error_Catalog(0006, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0026) = 0000: AL_Error_Catalog(0006, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0027) = 0000: AL_Error_Catalog(0006, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0028) = 0000: AL_Error_Catalog(0006, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0029) = 0000: AL_Error_Catalog(0006, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0030) = 0000: AL_Error_Catalog(0006, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0031) = 0000: AL_Error_Catalog(0006, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0032) = 0000: AL_Error_Catalog(0006, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0033) = 0000: AL_Error_Catalog(0006, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0034) = 0000: AL_Error_Catalog(0006, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0035) = 0000: AL_Error_Catalog(0006, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0036) = 0000: AL_Error_Catalog(0006, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0037) = 0000: AL_Error_Catalog(0006, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0038) = 0000: AL_Error_Catalog(0006, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0039) = 0000: AL_Error_Catalog(0006, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0040) = 0000: AL_Error_Catalog(0006, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0041) = 0000: AL_Error_Catalog(0006, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0042) = 0000: AL_Error_Catalog(0006, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0043) = 0000: AL_Error_Catalog(0006, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0044) = 0000: AL_Error_Catalog(0006, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0045) = 0000: AL_Error_Catalog(0006, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0046) = 0000: AL_Error_Catalog(0006, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0047) = 0000: AL_Error_Catalog(0006, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0048) = 0000: AL_Error_Catalog(0006, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0049) = 0000: AL_Error_Catalog(0006, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0050) = 0000: AL_Error_Catalog(0006, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0051) = 0000: AL_Error_Catalog(0006, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0052) = 0000: AL_Error_Catalog(0006, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0053) = 0000: AL_Error_Catalog(0006, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0054) = 0000: AL_Error_Catalog(0006, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0055) = 0000: AL_Error_Catalog(0006, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0056) = 0000: AL_Error_Catalog(0006, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0057) = 0000: AL_Error_Catalog(0006, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0058) = 0000: AL_Error_Catalog(0006, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0059) = 0000: AL_Error_Catalog(0006, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0060) = 0000: AL_Error_Catalog(0006, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0061) = 0000: AL_Error_Catalog(0006, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0062) = 0000: AL_Error_Catalog(0006, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0063) = 0000: AL_Error_Catalog(0006, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0064) = 0000: AL_Error_Catalog(0006, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0065) = 0000: AL_Error_Catalog(0006, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0066) = 0000: AL_Error_Catalog(0006, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0067) = 0000: AL_Error_Catalog(0006, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0068) = 0000: AL_Error_Catalog(0006, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0069) = 0000: AL_Error_Catalog(0006, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0070) = 0000: AL_Error_Catalog(0006, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0071) = 0000: AL_Error_Catalog(0006, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0072) = 0000: AL_Error_Catalog(0006, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0073) = 0000: AL_Error_Catalog(0006, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0074) = 0000: AL_Error_Catalog(0006, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0075) = 0000: AL_Error_Catalog(0006, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0076) = 0000: AL_Error_Catalog(0006, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0077) = 0000: AL_Error_Catalog(0006, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0078) = 0000: AL_Error_Catalog(0006, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0079) = 0000: AL_Error_Catalog(0006, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0080) = 0000: AL_Error_Catalog(0006, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0081) = 0000: AL_Error_Catalog(0006, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0082) = 0000: AL_Error_Catalog(0006, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0083) = 0000: AL_Error_Catalog(0006, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0084) = 0000: AL_Error_Catalog(0006, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0085) = 0000: AL_Error_Catalog(0006, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0086) = 0000: AL_Error_Catalog(0006, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0087) = 0000: AL_Error_Catalog(0006, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0088) = 0000: AL_Error_Catalog(0006, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0089) = 0000: AL_Error_Catalog(0006, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0090) = 0000: AL_Error_Catalog(0006, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0091) = 0000: AL_Error_Catalog(0006, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0092) = 0000: AL_Error_Catalog(0006, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0093) = 0000: AL_Error_Catalog(0006, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0094) = 0000: AL_Error_Catalog(0006, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0095) = 0000: AL_Error_Catalog(0006, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0096) = 0000: AL_Error_Catalog(0006, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0097) = 0000: AL_Error_Catalog(0006, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0098) = 0000: AL_Error_Catalog(0006, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0006, 0000, 0099) = 0000: AL_Error_Catalog(0006, 0001, 0099) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0000) = 0000: AL_Error_Catalog(0007, 0001, 0000) = "Invalid Value"
        AL_Error_Catalog(0007, 0000, 0001) = 0000: AL_Error_Catalog(0007, 0001, 0001) = "Value is Nothing"
        AL_Error_Catalog(0007, 0000, 0002) = 0000: AL_Error_Catalog(0007, 0001, 0002) = "Value Underflow"
        AL_Error_Catalog(0007, 0000, 0003) = 0000: AL_Error_Catalog(0007, 0001, 0003) = "Value Overflow"
        AL_Error_Catalog(0007, 0000, 0004) = 0000: AL_Error_Catalog(0007, 0001, 0004) = "Object not Initialized"
        AL_Error_Catalog(0007, 0000, 0005) = 0000: AL_Error_Catalog(0007, 0001, 0005) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0006) = 0000: AL_Error_Catalog(0007, 0001, 0006) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0007) = 0000: AL_Error_Catalog(0007, 0001, 0007) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0008) = 0000: AL_Error_Catalog(0007, 0001, 0008) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0009) = 0000: AL_Error_Catalog(0007, 0001, 0009) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0010) = 0000: AL_Error_Catalog(0007, 0001, 0010) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0011) = 0000: AL_Error_Catalog(0007, 0001, 0011) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0012) = 0000: AL_Error_Catalog(0007, 0001, 0012) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0013) = 0000: AL_Error_Catalog(0007, 0001, 0013) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0014) = 0000: AL_Error_Catalog(0007, 0001, 0014) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0015) = 0000: AL_Error_Catalog(0007, 0001, 0015) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0016) = 0000: AL_Error_Catalog(0007, 0001, 0016) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0017) = 0000: AL_Error_Catalog(0007, 0001, 0017) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0018) = 0000: AL_Error_Catalog(0007, 0001, 0018) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0019) = 0000: AL_Error_Catalog(0007, 0001, 0019) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0020) = 0000: AL_Error_Catalog(0007, 0001, 0020) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0021) = 0000: AL_Error_Catalog(0007, 0001, 0021) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0022) = 0000: AL_Error_Catalog(0007, 0001, 0022) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0023) = 0000: AL_Error_Catalog(0007, 0001, 0023) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0024) = 0000: AL_Error_Catalog(0007, 0001, 0024) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0025) = 0000: AL_Error_Catalog(0007, 0001, 0025) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0026) = 0000: AL_Error_Catalog(0007, 0001, 0026) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0027) = 0000: AL_Error_Catalog(0007, 0001, 0027) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0028) = 0000: AL_Error_Catalog(0007, 0001, 0028) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0029) = 0000: AL_Error_Catalog(0007, 0001, 0029) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0030) = 0000: AL_Error_Catalog(0007, 0001, 0030) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0031) = 0000: AL_Error_Catalog(0007, 0001, 0031) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0032) = 0000: AL_Error_Catalog(0007, 0001, 0032) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0033) = 0000: AL_Error_Catalog(0007, 0001, 0033) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0034) = 0000: AL_Error_Catalog(0007, 0001, 0034) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0035) = 0000: AL_Error_Catalog(0007, 0001, 0035) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0036) = 0000: AL_Error_Catalog(0007, 0001, 0036) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0037) = 0000: AL_Error_Catalog(0007, 0001, 0037) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0038) = 0000: AL_Error_Catalog(0007, 0001, 0038) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0039) = 0000: AL_Error_Catalog(0007, 0001, 0039) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0040) = 0000: AL_Error_Catalog(0007, 0001, 0040) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0041) = 0000: AL_Error_Catalog(0007, 0001, 0041) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0042) = 0000: AL_Error_Catalog(0007, 0001, 0042) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0043) = 0000: AL_Error_Catalog(0007, 0001, 0043) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0044) = 0000: AL_Error_Catalog(0007, 0001, 0044) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0045) = 0000: AL_Error_Catalog(0007, 0001, 0045) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0046) = 0000: AL_Error_Catalog(0007, 0001, 0046) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0047) = 0000: AL_Error_Catalog(0007, 0001, 0047) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0048) = 0000: AL_Error_Catalog(0007, 0001, 0048) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0049) = 0000: AL_Error_Catalog(0007, 0001, 0049) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0050) = 0000: AL_Error_Catalog(0007, 0001, 0050) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0051) = 0000: AL_Error_Catalog(0007, 0001, 0051) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0052) = 0000: AL_Error_Catalog(0007, 0001, 0052) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0053) = 0000: AL_Error_Catalog(0007, 0001, 0053) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0054) = 0000: AL_Error_Catalog(0007, 0001, 0054) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0055) = 0000: AL_Error_Catalog(0007, 0001, 0055) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0056) = 0000: AL_Error_Catalog(0007, 0001, 0056) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0057) = 0000: AL_Error_Catalog(0007, 0001, 0057) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0058) = 0000: AL_Error_Catalog(0007, 0001, 0058) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0059) = 0000: AL_Error_Catalog(0007, 0001, 0059) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0060) = 0000: AL_Error_Catalog(0007, 0001, 0060) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0061) = 0000: AL_Error_Catalog(0007, 0001, 0061) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0062) = 0000: AL_Error_Catalog(0007, 0001, 0062) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0063) = 0000: AL_Error_Catalog(0007, 0001, 0063) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0064) = 0000: AL_Error_Catalog(0007, 0001, 0064) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0065) = 0000: AL_Error_Catalog(0007, 0001, 0065) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0066) = 0000: AL_Error_Catalog(0007, 0001, 0066) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0067) = 0000: AL_Error_Catalog(0007, 0001, 0067) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0068) = 0000: AL_Error_Catalog(0007, 0001, 0068) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0069) = 0000: AL_Error_Catalog(0007, 0001, 0069) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0070) = 0000: AL_Error_Catalog(0007, 0001, 0070) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0071) = 0000: AL_Error_Catalog(0007, 0001, 0071) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0072) = 0000: AL_Error_Catalog(0007, 0001, 0072) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0073) = 0000: AL_Error_Catalog(0007, 0001, 0073) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0074) = 0000: AL_Error_Catalog(0007, 0001, 0074) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0075) = 0000: AL_Error_Catalog(0007, 0001, 0075) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0076) = 0000: AL_Error_Catalog(0007, 0001, 0076) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0077) = 0000: AL_Error_Catalog(0007, 0001, 0077) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0078) = 0000: AL_Error_Catalog(0007, 0001, 0078) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0079) = 0000: AL_Error_Catalog(0007, 0001, 0079) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0080) = 0000: AL_Error_Catalog(0007, 0001, 0080) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0081) = 0000: AL_Error_Catalog(0007, 0001, 0081) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0082) = 0000: AL_Error_Catalog(0007, 0001, 0082) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0083) = 0000: AL_Error_Catalog(0007, 0001, 0083) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0084) = 0000: AL_Error_Catalog(0007, 0001, 0084) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0085) = 0000: AL_Error_Catalog(0007, 0001, 0085) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0086) = 0000: AL_Error_Catalog(0007, 0001, 0086) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0087) = 0000: AL_Error_Catalog(0007, 0001, 0087) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0088) = 0000: AL_Error_Catalog(0007, 0001, 0088) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0089) = 0000: AL_Error_Catalog(0007, 0001, 0089) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0090) = 0000: AL_Error_Catalog(0007, 0001, 0090) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0091) = 0000: AL_Error_Catalog(0007, 0001, 0091) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0092) = 0000: AL_Error_Catalog(0007, 0001, 0092) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0093) = 0000: AL_Error_Catalog(0007, 0001, 0093) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0094) = 0000: AL_Error_Catalog(0007, 0001, 0094) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0095) = 0000: AL_Error_Catalog(0007, 0001, 0095) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0096) = 0000: AL_Error_Catalog(0007, 0001, 0096) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0097) = 0000: AL_Error_Catalog(0007, 0001, 0097) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0098) = 0000: AL_Error_Catalog(0007, 0001, 0098) = "PLACEHOLDER"
        AL_Error_Catalog(0007, 0000, 0099) = 0000: AL_Error_Catalog(0007, 0001, 0099) = "PLACEHOLDER"
    AL_Error_Initialized = True
    End If

End Sub