Public AL_ErrorInitialization As Boolean
Public AL_Error_Sheet As Worksheet
Public AL_Error_Range As Range
Public AL_ErrorLib_Range As Range
Public AL_Error_Index As Integer
Public Const Is_Error As Boolean = False
Public Const No_Error As Boolean = True
Public Const AL_ErrorCategory_System As Integer = 1
Public Const AL_ErrorCategory_Workbook As Integer = 2
Public Const AL_ErrorCategory_Worksheet As Integer = 3
Public Const AL_ErrorCategory_Linker As Integer = 4
Public Const AL_ErrorCategory_Compiler As Integer = 5
Public Const AL_ErrorCategory_Module As Integer = 6
Public Const AL_ErrorCategory_Class As Integer = 7
Public Const AL_ErrorCategory_Userform As Integer = 8
    
Sub Build()

    AL_Error_Create
    AL_BuildApplication "YOUR_FILEPATH", "YOUR_FILEPATH", 1

End Sub

Sub AL_ErrorCategory_Class_Sub()
    Dim Distance As Integer

    Distance = 30

    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"

    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Class
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Class"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "Invalid Value"
    
    AL_ErrorLib_Range.Offset(2, Distance + 0).Value = AL_ErrorCategory_Class
    AL_ErrorLib_Range.Offset(2, Distance + 1).Formula = "Class"
    AL_ErrorLib_Range.Offset(2, Distance + 2).Value = 2
    AL_ErrorLib_Range.Offset(2, Distance + 3).Formula = "Value is Nothing"

    AL_ErrorLib_Range.Offset(3, Distance + 0).Value = AL_ErrorCategory_Class
    AL_ErrorLib_Range.Offset(3, Distance + 1).Formula = "Class"
    AL_ErrorLib_Range.Offset(3, Distance + 2).Value = 3
    AL_ErrorLib_Range.Offset(3, Distance + 3).Formula = "Value Underflow"

    AL_ErrorLib_Range.Offset(4, Distance + 0).Value = AL_ErrorCategory_Class
    AL_ErrorLib_Range.Offset(4, Distance + 1).Formula = "Class"
    AL_ErrorLib_Range.Offset(4, Distance + 2).Value = 4
    AL_ErrorLib_Range.Offset(4, Distance + 3).Formula = "Value Overflow"

End Sub

Sub AL_ErrorCategory_Compiler_Sub()
    Dim Distance As Integer

    Distance = 20

    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"

    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Compiler
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Compiler"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "PLACEHOLDER"

End Sub

    Sub AL_ErrorCategory_Linker_Sub()
Dim Distance As Integer
    
Distance = 15
    
    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"
    
    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Linker
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Linker"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "PLACEHOLDER"
    
    End Sub

Sub AL_ErrorCategory_Module_Sub()
    Dim Distance As Integer

    Distance = 25

    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"

    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Module
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Module"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "PLACEHOLDER"

End Sub

    Sub AL_ErrorCategory_System_Sub()
Dim Distance As Integer
    
Distance = 0
    
    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"
    
    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "ErrorCategory doesnt exist"
    
    AL_ErrorLib_Range.Offset(2, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(2, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(2, Distance + 2).Value = 2
    AL_ErrorLib_Range.Offset(2, Distance + 3).Formula = "Value isnt available"
    
    AL_ErrorLib_Range.Offset(3, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(3, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(3, Distance + 2).Value = 3
    AL_ErrorLib_Range.Offset(3, Distance + 3).Formula = "Value isnt defined"
    
    AL_ErrorLib_Range.Offset(4, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(4, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(4, Distance + 2).Value = 4
    AL_ErrorLib_Range.Offset(4, Distance + 3).Formula = "Value is Nothing"
    
    AL_ErrorLib_Range.Offset(5, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(5, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(5, Distance + 2).Value = 5
    AL_ErrorLib_Range.Offset(5, Distance + 3).Formula = "Value Overflow"
    
    AL_ErrorLib_Range.Offset(6, Distance + 0).Value = AL_ErrorCategory_System
    AL_ErrorLib_Range.Offset(6, Distance + 1).Formula = "System"
    AL_ErrorLib_Range.Offset(6, Distance + 2).Value = 6
    AL_ErrorLib_Range.Offset(6, Distance + 3).Formula = "Value Underflow"
    
    End Sub

Sub AL_ErrorCategory_Userform_Sub()
    Dim Distance As Integer

    Distance = 35

    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"

    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Userform
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Userform"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "PLACEHOLDER"

End Sub

    Sub AL_ErrorCategory_Workbook_Sub()
Dim Distance As Integer
    
Distance = 5
    
    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"
    
    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "Errormessage doesnt exist"
    
    AL_ErrorLib_Range.Offset(2, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(2, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(2, Distance + 2).Value = 2
    AL_ErrorLib_Range.Offset(2, Distance + 3).Formula = "Workbook doesnt exist"
    
    AL_ErrorLib_Range.Offset(3, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(3, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(3, Distance + 2).Value = 3
    AL_ErrorLib_Range.Offset(3, Distance + 3).Formula = "Instance already exists"
    
    AL_ErrorLib_Range.Offset(4, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(4, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(4, Distance + 2).Value = 4
    AL_ErrorLib_Range.Offset(4, Distance + 3).Formula = "Dependency missing"
    
    AL_ErrorLib_Range.Offset(5, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(5, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(5, Distance + 2).Value = 5
    AL_ErrorLib_Range.Offset(5, Distance + 3).Formula = "Not available in Workbook"
    
    AL_ErrorLib_Range.Offset(6, Distance + 0).Value = AL_ErrorCategory_Workbook
    AL_ErrorLib_Range.Offset(6, Distance + 1).Formula = "Workbook"
    AL_ErrorLib_Range.Offset(6, Distance + 2).Value = 6
    AL_ErrorLib_Range.Offset(6, Distance + 3).Formula = "Instance doesnt exists"
    
    End Sub

Sub AL_ErrorCategory_Worksheet_Sub()
    Dim Distance As Integer

    Distance = 10

    AL_ErrorLib_Range.Offset(0, Distance + 0).Formula = "Error Category Index"
    AL_ErrorLib_Range.Offset(0, Distance + 1).Formula = "Error Category"
    AL_ErrorLib_Range.Offset(0, Distance + 2).Formula = "Error Type"
    AL_ErrorLib_Range.Offset(0, Distance + 3).Formula = "Error Message"

    AL_ErrorLib_Range.Offset(1, Distance + 0).Value = AL_ErrorCategory_Worksheet
    AL_ErrorLib_Range.Offset(1, Distance + 1).Formula = "Worksheet"
    AL_ErrorLib_Range.Offset(1, Distance + 2).Value = 1
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "Worksheet already exists"

    AL_ErrorLib_Range.Offset(2, Distance + 0).Value = AL_ErrorCategory_Worksheet
    AL_ErrorLib_Range.Offset(2, Distance + 1).Formula = "Worksheet"
    AL_ErrorLib_Range.Offset(2, Distance + 2).Value = 2
    AL_ErrorLib_Range.Offset(2, Distance + 3).Formula = "Worksheet doesnt exist"

End Sub

Function AL_Check_Instance(ByVal InstanceName As String, ByVal InstanceExistence As Boolean) As Boolean

    ' InstanceExistence = True exists and throws error
    ' InstanceExistence = False doesnt exist and throws error
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VbProj = ThisWorkbook.VBProject
    For Each VBComp In VbProj.VBComponents
        If InstanceExistence = True Then
                If VBComp.Name = InstanceName Then
                    AL_Error_Print 2, 3, InstanceName
                    AL_Error_Show 2, 3, InstanceName
                    AL_Check_Instance = Is_Error
                    Exit Function
                End If
            Else
                If VBComp.Name = InstanceName Then
                    AL_Check_Instance = True
                    Exit Function
                End If
        End If
    Next VBComp
    If InstanceExistence = True Then
        AL_Check_Instance = True
        Exit Function
    End If
    AL_Error_Print 2, 6, InstanceName
    AL_Error_Show 2, 6, InstanceName
    AL_Check_Instance = Is_Error

End Function

Function AL_Check_Sheet(ByVal InstanceName As String, ByVal InstanceExistence As Boolean) As Boolean

    ' InstanceExistence = True exists and throws error
    ' InstanceExistence = False doesnt exist and throws error
    Dim ws As Worksheet

    For Each ws In Worksheets
        If InstanceExistence = True Then
                If ws.Name = InstanceName Then
                    AL_Error_Print 3, 1, InstanceName
                    AL_Error_Show 3, 1, InstanceName
                    AL_Check_Sheet = Is_Error
                    Exit Function
                End If
            Else
                If ws.Name = InstanceName Then
                    AL_Check_Sheet = True
                    Exit Function
                End If
                End If
    Next ws
    If InstanceExistence = True Then
        AL_Check_Sheet = True
        Exit Function
            End If
    AL_Error_Print 3, 2, InstanceName
    AL_Error_Show 3, 2, InstanceName
    AL_Check_Sheet = Is_Error

End Function

Function AL_Check_String(ByVal Text As String, ByVal CheckType As Integer) As Boolean

    Select Case CheckType
        Case 0
            If Text = "" Then
                AL_Check_String = Is_Error
                AL_Error_Print 1, 4, InstanceName
                AL_Error_Show 1, 4, InstanceName
                Exit Function
            End If
    End Select
    AL_Check_String = True
            
End Function

Sub AL_Error_Create()

    Dim ws As Worksheet
    AL_Error_Initialize
    If AL_Check_Sheet("Error", True) = True Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Error"
        AL_Error_Initialize
        AL_ErrorCategory_Create
    End If
                
End Sub

Function AL_Error_GetCategory(ByVal ErrorCategory As Integer) As String

    Select Case ErrorCategory
        Case AL_ErrorCategory_System: AL_Error_GetCategory = "System"
        Case AL_ErrorCategory_Workbook: AL_Error_GetCategory = "Workbook"
        Case AL_ErrorCategory_Worksheet: AL_Error_GetCategory = "Worksheet"
        Case AL_ErrorCategory_Linker: AL_Error_GetCategory = "Linker"
        Case AL_ErrorCategory_Compiler: AL_Error_GetCategory = "Compiler"
        Case AL_ErrorCategory_Module: AL_Error_GetCategory = "Module"
        Case AL_ErrorCategory_Class: AL_Error_GetCategory = "Class"
        Case AL_ErrorCategory_Userform: AL_Error_GetCategory = "Userform"
        Case Else
            AL_Error_GetCategory = "UNKNOWN"
    End Select

End Function

Sub AL_Error_Initialize()

    If AL_ErrorInitialization <> True Then
        For Each ws In Worksheets
            If ws.Name = "Error" Then
                Set AL_Error_Sheet = ThisWorkbook.Sheets("Error")
                Set AL_Error_Range = AL_Error_Sheet.Range("A1")
                Set AL_ErrorLib_Range = AL_Error_Sheet.Range("H1")
                AL_ErrorInitialization = True
            End If
        Next ws
    End If
    AL_Error_Index = 0

End Sub

' ErrorCategory describes where the error comes from (Linker, Compiler, Module etc)
' ErrorType describes what error it is
' ErrorValue describes what caused the Error
Sub AL_Error_Print(ByVal ErrorCategory As Integer, ByVal ErrorType As Integer, Optional ByVal ErrorValue1 As Variant = 0, ByVal Optional ErrorValue2 As Variant = 0)

    AL_Error_Initialize
    Dim ErrorCategoryString As String
    ErrorCategoryString = AL_Error_GetCategory(ErrorCategory)
    Do Until AL_Error_Range.Offset(AL_Error_Index, 0).Formula = ""
        AL_Error_Index = AL_Error_Index + 1
    Loop
        AL_Error_Range.Offset(AL_Error_Index, 0).Formula = ErrorCategory
        AL_Error_Range.Offset(AL_Error_Index, 1).Formula = ErrorCategoryString
        AL_Error_Range.Offset(AL_Error_Index, 2).Formula = ErrorType
        AL_Error_Range.Offset(AL_Error_Index, 3).Formula = AL_ErrorLib_Range.Offset(ErrorType, 3 + (ErrorCategory * 5 - 5)).Formula
        If ErrorValue1 = 0 Then
                AL_Error_Range.Offset(AL_Error_Index, 4).Formula = "No_Error"
            Else
                AL_Error_Range.Offset(AL_Error_Index, 4).Formula = ErrorValue1
        End If
        If ErrorValue2 = 0 Then
                AL_Error_Range.Offset(AL_Error_Index, 5).Formula = "No_Error"
            Else
                AL_Error_Range.Offset(AL_Error_Index, 5).Formula = ErrorValue2
        End If

End Sub

' ErrorCategory describes where the error comes from (Linker, Compiler, Module etc)
' ErrorType describes what error it is
' ErrorValue describes what caused the Error
Sub AL_Error_Show(ByVal ErrorCategory As Integer, ByVal ErrorType As Integer, Optional ByVal ErrorValue1 As Variant = 0, ByVal Optional ErrorValue2 As Variant = 0)

    AL_Error_Initialize
    Dim ErrorCategoryString As String
    Dim ErrorMessage As String
    Dim ErrorString1 As String
    Dim ErrorString2 As String
    If ErrorValue1 = 0 Then
                ErrorString1 = "No_Error"
            Else
                ErrorString1 = ErrorValue1
        End If
        If ErrorValue2 = 0 Then
                ErrorString2 = "No_Error"
            Else
                ErrorString2 = ErrorValue2
        End If
    ErrorCategoryString = AL_Error_GetCategory(ErrorCategory)
    ErrorMessage = "( " & ErrorCategory & " ): ( " & ErrorCategoryString & " ) / ( " & ErrorType & " ): ( " & (AL_ErrorLib_Range.Offset(ErrorType, 3 + (ErrorCategory * 5 - 5)).Formula) & " ) / ( " & ErrorValue1 & " ) ( " & ErrorValue2 & " )"
    MsgBox(ErrorMessage)

End Sub

Function AL_Check_Integer(ByVal IntegerValue As Integer)

    Select Case IntegerValue
        Case 0
            AL_Error_Print 1, 3, IntegerValue
            AL_Error_Show 1, 3, IntegerValue
            AL_Check_Integer = Is_Error
    End Select
    AL_Check_Integer = True

End Function

    ' Loops through a folder and Builds a single Module for each Folder
Sub AL_BuildApplication(ByVal FolderPath As String, ByVal OGFilePath As String, ByVal InstanceType As Integer)

    ' Create a FileSystemObject
    If AL_Check_String(FolderPath, 0) = Is_Error Or AL_Check_String(OGFilePath, 0) = Is_Error Then
        Exit Sub
    End If
    If AL_Check_Integer(InstanceType) = Is_Error Then
        Exit Sub
    End If

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim InstanceName As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBCodeModule As VBIDE.CodeModule

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the Folder object
    Set Folder = fso.GetFolder(FolderPath)

    ' Checks if Instance already exists
    InstanceName = fso.GetBaseName(FolderPath)
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = InstanceName Then
                AL_Error_Print 2, 3, InstanceName
                AL_Error_Show 2, 3, InstanceName
                Exit Sub
        End If
    Next VBComp

    ' Create an Instance (Module, Class, Userform)
    Select Case InstanceName
        Case "AL_Modules"
            InstanceType = 1
        Case "AL_Classes"
            InstanceType = 2
        Case "AL_Userforms"
            InstanceType = 3
    End Select
    AL_Include_CreateInstance InstanceType, InstanceName
    Set VBComp = VBProj.VBComponents(InstanceName)
    Set VBCodeModule = VBComp.CodeModule

    ' Loop through each File in the Folder
    For Each File In Folder.Files
        If File.Name Like "AL_*" Then
                AL_Include_AddCode VBCodeModule, File.Path
        End If
    Next File
    
    ' Recursively loop through each SubFolder
    For Each SubFolder In Folder.SubFolders
        If SubFolder.Name Like "AL_*" Then
            AL_BuildApplication SubFolder.Path, FolderPath, InstanceType
        End If
    Next SubFolder

End Sub

' Adds Code to a given Instance
Sub AL_Include_AddCode(VBCodeModule As VBIDE.CodeModule, FilePath As String)

    If AL_Check_String(FilePath, 0) = Is_Error Then
        Exit Sub
    End If
    If VBCodeModule Is Nothing Then
        AL_Error_Print 1, 4, "AL_Include_AddCode"
        AL_Error_Print 1, 4, "AL_Include_AddCode"
    End If
    Dim FileLine As String
    Dim Index As Long
    Dim FileNumber As Integer

    Index = 1
    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber
    Do Until EOF(FileNumber)
        Line Input #FileNumber, FileLine
        VBCodeModule.InsertLines Index, FileLine
        Index = Index + 1
    Loop
    Close #FileNumber

End Sub

Sub AL_ErrorCategory_Create()

    AL_ErrorCategory_System_Sub
    AL_ErrorCategory_Workbook_Sub
    AL_ErrorCategory_Worksheet_Sub
    AL_ErrorCategory_Linker_Sub
    AL_ErrorCategory_Compiler_Sub
    AL_ErrorCategory_Module_Sub
    AL_ErrorCategory_Class_Sub
    AL_ErrorCategory_Userform_Sub

End Sub

' Creates an Instance (Module, Class, Userform)
Sub AL_Include_CreateInstance(ByVal InstanceType, ByVal InstanceName As String)

    Select Case InstanceType
        Case 1, 2, 3, 11, 100:
            Dim vbProj As VBIDE.VBProject
            Dim vbComp As VBIDE.VBComponent

            Set vbProj = ThisWorkbook.VBProject
            Set vbComp = vbProj.VBComponents.Add(InstanceType)
            vbComp.Name = InstanceName
        Case Else
            AL_Error_Print 1, 2, InstanceType, InstanceName
            AL_Error_Show 1, 2, InstanceType, InstanceName
            Exit Sub
    End Select

End Sub