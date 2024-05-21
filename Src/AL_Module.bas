' The Functions here are meant to be used when building all files in a given Folder, nowhere else
' The Functions used in this File are implemented in another way with more security
' Name Module where this code is run to "AL_BuildModule" and replace "YOUR-FILEPATH" with your filepath

Option Explicit

Private Const BuildModule As String = "AL_BuildModule"
Private Const BuildDataFile As String = "BuildData"

Sub AL_BuildApplication()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent

    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name <> BuildModule Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
    AL_BuildFolder "YOUR_FOLDERPATH", Empty, Empty

End Sub

' Loops through a folder and Builds a single Component for each Folder and iterates through Subfolders
Private Sub AL_BuildFolder(FolderPath As String, ComponentType As Long, ComponentName As String)

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBCodeModule As VBIDE.CodeModule
    Dim FileArray() As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(FolderPath)
    Set FileArray = AL_GetBuildData(Folder)
    ' Get basic information from current folder and create new Component, if it doesnt already exist
    ComponentType = FileArray(0)
    If ComponentName <> Replace(FileArray(1), (FolderPath & "\"), "") Then
        ComponentName = Replace(FileArray(1), (FolderPath & "\"), "")
        For Each VBComp in VBProj.Component
            If VBComp.Name = ComponentName Then
                Exit For
            End If
        Next
        AL_BuildCreateVBComponent VBProj, FileArray(0), ComponentName
    End If
    Set VBComp = VBProj.VBComponents(FileArray(1))
    Set VBCodeModule = VBComp.CodeModule
    ' Add code in Stackform to Component
    For i = Ubound(FileArray) To 2 Step-1
        AL_BuildAddCode VBCodeModule, FileArray(1)
    Next i
    If ComponentType = vbext_ct_MSForm Then
        AL_Instant_Exe Folder, VBProj
    End If
    ' Loop through all Subfolders and Repeat
    For Each SubFolder In Folder.SubFolders
        AL_BuildFolder SubFolder.Path, ComponentType, ComponentName
    Next SubFolder

End Sub

' Adds Code to a given Component
Private Sub AL_BuildAddCode(VBCodeModule As VBIDE.CodeModule, FilePath As String)

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

' Creates an Component
Private Sub AL_BuildCreateVBComponent(VBProj As VBIDE.VBProject, ComponentType As Integer, ComponentName As String)

    Dim VBComp As VBIDE.VBComponent
    On Error GoTo Error
    Set VBComp = VBProj.VBComponents.Add(ComponentType)
    VBComp.Name = ComponentName
    Exit Sub
    Error:
    MsgBox("Component exists already:" & ComponentName & ", OR" "ComponentType doesnt exist:" ComponentType)

End Sub

' Gets all information needed to Include a Folder defined in BuildData
Private Function AL_GetBuildData(Folder As Object) As Variant()

    Dim File As Object
    
    Dim FileArray() As Variant
    For Each File In Folder.Files
        If File.Name = BuildDataFile Then
            Dim FileLine As String
            Dim Index As Long
            Dim FileNumber As Integer
            Dim i As Long

            Index = 1
            i = 0
            FileNumber = FreeFile
            Open File.Path For Input As #FileNumber
            ' Get Base Information
            Line Input #FileNumber, FileLine
            FileArray(i) = CLong(Replace(FileLine, "ComponentType = ", ""))
            Select Case FileArray(i)
                Case vbext_ct_ActiveXDesigner
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_MSForm
                    Index = Index + 1
                    i = i + 1
                    Line Input #FileNumber, FileLine
                    FileArray(i) = Replace(FileLine, "ComponentName = ", "")
                    i = i + 2
                    Index = Index + 1
            End Select
            ' Get all files in Order
            Do Until EOF(FileNumber)
                Line Input #FileNumber, FileLine
                ReDim Preserve FileArray(i)
                FileArray(i) = File.Path & "\" & FileLine
                i = i + 1
                Index = Index + 1
            Loop
            Close #FileNumber
            Exit For
        End If
    Next File
    Set AL_BuildFolder = FileArray

End Function

Private Sub AL_Instant_Exe(Folder As Object, VBProj As VBIDE.VBProject)

    Dim VBComp As VBIDE.VBComponent

    Set VBComp = VBProj.VBComponents.Add(ComponentType)
    VBComp.Name = "INSTANT"
    AL_BuildAddCode VBComp.CodeModule, (Folder.Path & "\InstantData")
    InstantData
    VBProj.VBComponents.Remove VBComp

End Sub