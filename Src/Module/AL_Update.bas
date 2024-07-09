Attribute VB_Name = "AL_Update"


Public AL_UpdateInitialization As Boolean
Public AL_Update_Sheet As Worksheet
Public AL_Update_Range As Range
Public AL_UpdateLib_Range As Range

' Adds Code to existing Component
Public Function AL_Update_AddCode(VBCodeModule As VBIDE.CodeModule, Optional FilePath As String = Empty, Optional CodeString As String = Empty) As Boolean

    ' If from String
    If FilePath = Empty And CodeString <> Empty Then
        VBCodeModule.InsertLines Index, CodeString
    ' If from File
    ElseIf FilePath <> Empty And CodeString = Empty Then
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
    ' Handle wrong Function input
    ElseIf FilePath = Empty And CodeString = Empty
        AL_Error_ComVar AL_Error_System, 0012, "FilePath", "CodeString"
        AL_Update_AddCode = AL_IS_ERROR
    ' Handle wrong Function input
    Else
        AL_Error_ComVar AL_Error_System, 0013, "FilePath", "CodeString"
        AL_Update_AddCode = AL_IS_ERROR

    End If
    AL_Update_AddCode = AL_NO_Error
    
End Function

    ' Adds New Component to Project
Public Function AL_Update_AddComponent(VBProj As VBIDE.VBProject, ComponentType As Integer, ComponentName As String) As Boolean
    
    Dim VBComp As VBIDE.VBComponent
    ' Error Handling
        If AL_Error_Obj(VBProj) = AL_IS_ERROR Then:                                Exit Function: End If
        If AL_Error_Var(ComponentName) = AL_IS_ERROR Then:                         Exit Function: End If
        If AL_Check_Instance(ComponentName, True) = AL_IS_ERROR Then:              Exit Function: End If
        If AL_Error_Component(VBProj, AL_Exist, ComponentName) = AL_IS_ERROR Then: Exit Function: End If
    '
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner, vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_MSForm, vbext_ct_StdModule
            Set VBComp = vbProj.VBComponents.Add(ComponentType)
            VBComp.Name = ComponentName 
        Case Else
            AL_Error_Show AL_Error_System, 1, ComponentType
            Exit Sub
    End Select
    AL_Update_AddComponent = AL_NO_ERROR
    
End Sub

Sub AL_Update_Change(InstanceName As String, OldProcedure As String, NewProcedureFilePath As String)

    If AL_Check_String(InstanceName, 0) = No_Error And AL_Check_String(OldProcedure, 0) = No_Error And AL_Check_String(NewProcedureFilePath, 0) = No_Error Then
        If AL_Check_Instance(InstanceName, False) = No_Error Then 
            Dim VBProj As VBIDE.VBProject
            Dim VBComp As VBIDE.VBComponent
            Dim VBCodeMod As VBIDE.CodeModule

            Set VBProj = ThisWorkbook.VBProject
            Set VBComp = VBProj.VBComponents(InstanceName)
            Set VBCodeMod = VBComp.CodeModule
            AL_Update_Delete InstanceName, OldProcedure
            AL_Update_Add NewProcedureFilePath, 0, InstanceName
        End If
    End If
        
End Sub

Sub AL_Update_Create()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = AL_Update_Component Then
            Exit Sub
        End If
    Next VBComp
    If AL_Update_AddComponent(VBProj, vbext_ct_StdModule, AL_Update_Component) = AL_IS_ERROR Then
        Exit Sub
    End If
    
End Sub

Sub AL_Update_Delete(InstanceName As String, FunctionName As String)
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim StartLine As Long
    Dim EndLine As Long
    Dim Index As Long
    Dim FileLine As String
    Dim FoundStart As Boolean

    If AL_Check_String(FunctionName, 0) = Is_Error Then
        Exit Sub
    End If
    Set VBProj = ThisWorkbook.VBProject
    If AL_Check_Instance(InstanceName, False) = Is_Error Then
        Exit Sub
    End If
    Set VBComp = VBProj.VBComponents(InstanceName)
    Set CodeMod = VBComp.CodeModule
    
    ' Deletes a Function from declaration to End
    For Index = 1 To CodeMod.CountOfLines
        If CodeMod.Lines(Index, 1) Like "*" & FunctionName & "*" Then
            StartLine = Index
            FoundStart = True
        End If
        If CodeMod.Lines(Index, 1) = "End Function" Or CodeMod.Lines(Index, 1) = "End Sub" And FoundStart = True Then
            EndLine = Index
            Exit For
        End If
    Next Index
    If AL_Check_Long(StartLine) = True And AL_Check_Long(EndLine) = True Then
            CodeMod.DeleteLines StartLine, EndLine - StartLine + 1
    End If

End Sub

Function AL_Update_Get(UpdateName As String) As Boolean
    
    Dim i As Integer

    If AL_Check_String(UpdateName, 0) = Is_Error Then
        Exit Function
    End If
    Do Until AL_Update_Range.Offset(i, 0).Formula = ""
        If AL_Update_Range.Offset(i, 0).Formula = UpdateName Then
            Exit Function
        End If
        i = i + 1
    Loop
    AL_Error_Print 1, 5, UpdateName
    AL_Error_Show 1, 5, UpdateName

End Function


                    
Sub AL_Update_Initialize()

    If AL_UpdateInitialization <> True Then
        For Each ws In Worksheets
            If ws.Name = "Update" Then
                    Set AL_Update_Sheet = ThisWorkbook.Sheets("Update")
                    Set AL_Update_Range = AL_Update_Sheet.Range("A2")
                    AL_UpdateInitialization = True
            End If
        Next ws
    End If

End Sub