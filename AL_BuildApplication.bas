Public AL_ErrorInitialization As Boolean
Public AL_Error_Sheet As Worksheet
Public AL_Error_Range As Range
Public AL_ErrorLib_Range As Range
Public AL_Error_Index As Integer
Public Const AL_ErrorCategory_System As Integer = 1
Public Const AL_ErrorCategory_Workbook As Integer = 2
Public Const AL_ErrorCategory_Worksheet As Integer = 3
Public Const AL_ErrorCategory_Linker As Integer = 4
Public Const AL_ErrorCategory_Compiler As Integer = 5
Public Const AL_ErrorCategory_Module As Integer = 6
Public Const AL_ErrorCategory_Class As Integer = 7
Public Const AL_ErrorCategory_Userform As Integer = 8
Sub AL_ErrorCreateBasicError()

    AL_ErrorCategory_System_Sub
    AL_ErrorCategory_Workbook_Sub
    AL_ErrorCategory_Worksheet_Sub
    AL_ErrorCategory_Linker_Sub
    AL_ErrorCategory_Compiler_Sub
    AL_ErrorCategory_Module_Sub
    AL_ErrorCategory_Class_Sub
    AL_ErrorCategory_Userform_Sub

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
    AL_ErrorLib_Range.Offset(1, Distance + 3).Formula = "PLACEHOLDER"

End Sub


Function AL_CheckLong(ByVal LongValue As Long)

    Select Case LongValue
        Case 0
            AL_ErrorPrint 1, 3, LongValue
            AL_ErrorShow 1, 3, LongValue
            AL_CheckLong = False
    End Select
    AL_CheckLong = True

End Function
' ErrorCategory describes where the error comes from (Linker, Compiler, Module etc)
' ErrorType describes what error it is
' ErrorValue describes what caused the Error
Sub AL_ErrorShow(ByVal ErrorCategory As Integer, ByVal ErrorType As Integer, Optional ByVal ErrorValue1 As Variant = 0, Optional ByVal ErrorValue2 As Variant = 0)
    
    AL_ErrorInitialize
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
    ErrorCategoryString = AL_ErrorGetCategory(ErrorCategory)
    ErrorMessage = "( " & ErrorCategory & " ): ( " & ErrorCategoryString & " ) / ( " & ErrorType & " ): ( " & (AL_ErrorLib_Range.Offset(ErrorType, 3 + (ErrorCategory * 5 - 5)).Formula) & " ) / ( " & ErrorValue1 & " ) ( " & ErrorValue2 & " )"
    MsgBox (ErrorMessage)
    
End Sub
' ErrorCategory describes where the error comes from (Linker, Compiler, Module etc)
' ErrorType describes what error it is
' ErrorValue describes what caused the Error
Sub AL_ErrorPrint(ByVal ErrorCategory As Integer, ByVal ErrorType As Integer, Optional ByVal ErrorValue1 As Variant = 0, Optional ByVal ErrorValue2 As Variant = 0)
    
    AL_ErrorInitialize
    Dim ErrorCategoryString As String
    ErrorCategoryString = AL_ErrorGetCategory(ErrorCategory)
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
Sub AL_ErrorNew(ByVal ErrorCategory As Integer, ByVal ErrorMessage As String)

    Dim Distance As Integer
    Dim I As Integer

    Distance = ((ErrorCategory * 5) - 5)

    Do Until AL_ErrorLib_Range.Offset(I, Distance + 1).Formula = ""
        I = I + 1
    Loop
    AL_ErrorLib_Range.Offset(I, Distance + 1).Value = ErrorCategory
    AL_ErrorLib_Range.Offset(I, Distance + 2).Formula = AL_ErrorGetCategory(ErrorCategory)
    AL_ErrorLib_Range.Offset(I, Distance + 3).Value = I
    AL_ErrorLib_Range.Offset(I, Distance + 4).Formula = ErrorMessage

End Sub
    
    
Sub AL_ErrorInitialize()

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
Function AL_ErrorGetCategory(ByVal ErrorCategory As Integer) As String

    Select Case ErrorCategory
        Case AL_ErrorCategory_System: AL_ErrorGetCategory = "System"
        Case AL_ErrorCategory_Workbook: AL_ErrorGetCategory = "Workbook"
        Case AL_ErrorCategory_Worksheet: AL_ErrorGetCategory = "Worksheet"
        Case AL_ErrorCategory_Linker: AL_ErrorGetCategory = "Linker"
        Case AL_ErrorCategory_Compiler: AL_ErrorGetCategory = "Compiler"
        Case AL_ErrorCategory_Module: AL_ErrorGetCategory = "Module"
        Case AL_ErrorCategory_Class: AL_ErrorGetCategory = "Class"
        Case AL_ErrorCategory_Userform: AL_ErrorGetCategory = "Userform"
        Case Else
            AL_ErrorGetCategory = "UNKNOWN"
    End Select

End Function
Sub AL_ErrorCreate()

    Dim ws As Worksheet
    AL_ErrorInitialize
    If AL_CheckSheet("Error", True) = True Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Error"
        AL_ErrorInitialize
        AL_ErrorCreateBasicError
    End If

End Sub
Sub AL_ErrorClear()

    AL_ErrorInitialize
    Range(AL_Error_Range, AL_Error_Range.Offset(100000, 5)).ClearContents

End Sub
Function AL_CheckString(ByVal Text As String, ByVal CheckType As Integer) As Boolean

    Select Case CheckType
        Case 0
            If Text = "" Then
                AL_CheckString = False
                AL_ErrorPrint 1, 4, InstanceName
                AL_ErrorShow 1, 4, InstanceName
                Exit Function
            End If

    End Select
    AL_CheckString = True

End Function
Function AL_CheckSheet(ByVal InstanceName As String, ByVal InstanceExistence As Boolean) As Boolean

    ' InstanceExistence = True exists and throws error
    ' InstanceExistence = False doesnt exist and throws error
    Dim ws As Worksheet

    For Each ws In Worksheets
        If InstanceExistence = True Then
                If ws.Name = InstanceName Then
                    AL_ErrorPrint 3, 1, InstanceName
                    AL_ErrorShow 3, 1, InstanceName
                    AL_CheckSheet = False
                    Exit Function
                End If
            Else
                If ws.Name = InstanceName Then
                    AL_CheckSheet = True
                    Exit Function
                End If
        End If
    Next ws
    If InstanceExistence = True Then
        AL_CheckSheet = True
        Exit Function
    End If
    AL_ErrorPrint 3, 2, InstanceName
    AL_ErrorShow 3, 2, InstanceName
    AL_CheckSheet = False

End Function
Function AL_CheckInstance(ByVal InstanceName As String, ByVal InstanceExistence As Boolean) As Boolean

    ' InstanceExistence = True exists and throws error
    ' InstanceExistence = False doesnt exist and throws error
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If InstanceExistence = True Then
                If VBComp.Name = InstanceName Then
                    AL_ErrorPrint 2, 3, InstanceName
                    AL_ErrorShow 2, 3, InstanceName
                    AL_CheckInstance = False
                    Exit Function
                End If
            Else
                If VBComp.Name = InstanceName Then
                    AL_CheckInstance = True
                    Exit Function
                End If
        End If
    Next VBComp
    If InstanceExistence = True Then
        AL_CheckInstance = True
        Exit Function
    End If
    AL_ErrorPrint 2, 6, InstanceName
    AL_ErrorShow 2, 6, InstanceName
    AL_CheckInstance = False

End Function

' Creates an Instance (Module, Class, Userform)
Sub AL_Include_CreateInstance(ByVal InstanceType, ByVal InstanceName As String)

    Select Case InstanceType
        Case 1, 2, 3, 11, 100:
            Dim VBProj As VBIDE.VBProject
            Dim VBComp As VBIDE.VBComponent

            Set VBProj = ThisWorkbook.VBProject
            Set VBComp = VBProj.VBComponents.Add(InstanceType)
            VBComp.Name = InstanceName
        Case Else
            AL_ErrorShow 1, 2, InstanceType, InstanceName
            Exit Sub
    End Select

End Sub
' Adds Code to a given Instance
Sub AL_Include_AddCode(VBCodeModule As VBIDE.CodeModule, FilePath As String)

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
' Loops through Folders (Create Module per folder)
' Loop through Files and adds Code to Module
Sub AL_IncludeFolder(ByVal FolderPath As String)

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim InstanceName As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBCodeModule As VBIDE.CodeModule
    
    ' Create a FileSystemObject
    If FolderPath = "" Then
        AL_ErrorPrint 1, 3, FolderPath
        AL_ErrorShow 1, 3, FolderPath
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the Folder object
    Set Folder = fso.GetFolder(FolderPath)
    
    ' Checks if Instance already exists
    InstanceName = fso.GetBaseName(FolderPath)
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = InstanceName Then
                AL_ErrorPrint 2, 3, InstanceName
                AL_ErrorShow 2, 3, InstanceName
                Exit Sub
        End If
    Next VBComp

    ' Check if dependencies are included
    For Each File In Folder.Files
        If File.Path Like "*Dependencies" Then
                If AL_CheckDependencies(File.Path) = True Then
                        Exit Sub
                End If
        End If
    Next File

    ' Create an Instance (Module, Class, Userform)
    AL_Include_CreateInstance 1, InstanceName
    
    Set VBComp = VBProj.VBComponents(InstanceName)
    Set VBCodeModule = VBComp.CodeModule


    ' Loop through each File in the Folder
    For Each File In Folder.Files
        If File.Path Like "*Dependencies" Or File.Path Like "*README" Then
            Else
                AL_Include_AddCode VBCodeModule, File.Path
        End If
    Next File
    
    ' Recursively loop through each SubFolder
    For Each SubFolder In Folder.SubFolders
        AL_Include SubFolder.Path
    Next SubFolder
    
End Sub
Function AL_CheckDependencies(ByVal FilePath As String) As Boolean

    Dim FileLine As String
    Dim Index As Long
    Dim FileNumber As Integer
    Dim Included As Boolean
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ThisWorkbook.VBProject

    Index = 1
    FileNumber = FreeFile
    Open FilePath For Input As #FileNumber

    Do Until EOF(FileNumber)
        Line Input #FileNumber, FileLine
        For Each VBComp In VBProj.VBComponents
                If VBComp.Name = FileLine Then
                    Included = True
                End If
        Next VBComp
        If Included = False Then
                AL_ErrorPrint 2, 4, FileLine
                AL_ErrorShow 2, 4, FileLine
                AL_CheckDependencies = True
                Exit Function
            Else
                Included = False
        End If
        Index = Index + 1
    Loop
    Close #FileNumber
    AL_CheckDependencies = False

End Function
    
Sub AL_BuildApplication(ByVal FolderPath As String, ByVal OGFilePath As String, ByVal InstanceType As Integer)

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim InstanceName As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBCodeModule As VBIDE.CodeModule


    AL_ErrorCreate
    ' Create a FileSystemObject
    If FolderPath = "" Then
        AL_ErrorPrint 1, 3, FolderPath
        AL_ErrorShow 1, 3, FolderPath
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Get the Folder object
    Set Folder = fso.GetFolder(FolderPath)

    ' Checks if Instance already exists
    InstanceName = fso.GetBaseName(FolderPath)
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = InstanceName Then
                AL_ErrorPrint 2, 3, InstanceName
                AL_ErrorShow 2, 3, InstanceName
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
    MsgBox("Application Build")

End Sub

Sub Build()

    AL_ErrorCreate
    AL_BuildApplication "L:\RD\Automotive - Elastomere\Mitarbeiter\RD-AL\Projekte\AL_CustomLibrary", "L:\RD\Automotive - Elastomere\Mitarbeiter\RD-AL\Projekte\AL_CustomLibrary", 1

End Sub