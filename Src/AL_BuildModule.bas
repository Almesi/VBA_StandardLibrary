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
        If VBComp.Name <> BuildModule And VBComp.Type <> vbext_ct_Document Then
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
    Dim i As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(FolderPath)
    Set VBProj = ThisWorkbook.VBProject
    FileArray = AL_GetBuildData(Folder)
    ' Get basic information from current folder and create new Component, if it doesnt already exist
    If FileArray(0) <> Empty Then
        ComponentType = FileArray(0)
        If ComponentName <> Replace(FileArray(1), (FolderPath & "\"), "") Then
            ComponentName = Replace(FileArray(1), (FolderPath & "\"), "")
            For Each VBComp in VBProj.VBComponents
                If VBComp.Name = ComponentName Then
                    GoTo SkipComp
                End If
            Next
            AL_BuildCreateVBComponent VBProj, FileArray(0), ComponentName
        End If
SkipComp:
        Set VBComp = VBProj.VBComponents(FileArray(1))
        Set VBCodeModule = VBComp.CodeModule
        ' Add code in Stackform to Component
        For i = Ubound(FileArray) To 2 Step-1
            ' AL_BuildAddCode VBCodeModule, FileArray(i)
        Next i
        If ComponentType = vbext_ct_MSForm Then
            AL_Instant_Exe Folder, VBProj
        End If
    End If
    ' Loop through all Subfolders and Repeat
    For Each SubFolder In Folder.SubFolders
        AL_BuildFolder SubFolder.Path, ComponentType, ComponentName
        Debug.Print "Success:" & SubFolder.Path
    Next SubFolder

End Sub

' Adds Code to a given Component
Private Sub AL_BuildAddCode(VBCodeModule As VBIDE.CodeModule, FilePath As Variant)

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
    Debug.Print "    Success Code:" & FilePath

End Sub

' Creates an Component
Private Sub AL_BuildCreateVBComponent(VBProj As VBIDE.VBProject, ComponentType As Variant, ComponentName As String)

    Dim VBComp As VBIDE.VBComponent
    On Error GoTo Error
    Set VBComp = VBProj.VBComponents.Add(ComponentType)
    VBComp.Name = ComponentName
    Debug.Print "        Success Comp:" & ComponentName
    Exit Sub
    Error:
    MsgBox("Component exists already:" & ComponentName & ", OR ComponentType doesnt exist:" & ComponentType)

End Sub

' Gets all information needed to Include a Folder defined in BuildData
Private Function AL_GetBuildData(Folder As Object) As Variant()

    Dim File As Object
    Dim FileArray As Variant
    Dim ReturnArray() As Variant

    ReDim ReturnArray(3) As Variant
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
                FileLine = Input(LOF(FileNumber), 1)
            Close #FileNumber

            FileArray = Split(FileLine, vbLf)
            ReDim ReturnArray(UBound(FileArray) - 1)

            ReturnArray(0) = CLng(Replace(FileArray(0), "ComponentType = ", ""))
            Select Case ReturnArray(0)
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_MSForm
                    ReturnArray(1) = Replace(FileArray(1), "ComponentName = ", "")
            End Select

            ' Get all files in Order
            For i = 2 To UBound(ReturnArray)
                ReturnArray(i) = Replace(Folder.Path, "\BuildData", "") & "\" & FileArray(i + 1)
            Next i
            Exit For
        End If
    Next File
    AL_GetBuildData = ReturnArray
    Debug.Print "            Success Data:" & Folder.Path

End Function

Private Sub AL_Instant_Exe(Folder As Object, VBProj As VBIDE.VBProject)

    Dim VBComp As VBIDE.VBComponent

    Set VBComp = VBProj.VBComponents.Add(1)
    VBComp.Name = "INSTANT"
    AL_BuildAddCode VBComp.CodeModule, (Replace(Folder.Path, "\BuildData", "") & "\InstantData")
    Application.Run "InstantData"
    VBProj.VBComponents.Remove VBComp
    Debug.Print "                Success Instant:" & Folder.Path

End Sub
