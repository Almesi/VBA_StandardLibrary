' The Functions here are meant to be used when building all files in a given Folder, nowhere else
' The Functions used in this File are implemented in another way with more security
' Name Module where this code is run to "AL_BuildModule" and replace "YOUR-FILEPATH" with your filepath
Sub AL_BuildApplication()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim Const BuildModule As String = "AL_BuildModule"

    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name <> BuildModule Then
            VBProj.VBComponents.Remove VBComp
        End If
    Next VBComp
    AL_BuildApplication "YOUR_FOLDERPATH", "YOUR_FOLDERPATH", 1

End Sub

' Loops through a folder and Builds a single Component for each Folder
Sub AL_BuildFolder(FolderPath As String, OGFilePath As String, ComponentType As Integer)

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
    Dim ComponentName As String
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim VBCodeModule As VBIDE.CodeModule

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set Folder = fso.GetFolder(FolderPath)

    ComponentName = fso.GetBaseName(FolderPath)
    Set VBProj = ThisWorkbook.VBProject
    For Each VBComp In VBProj.VBComponents
        If VBComp.Name = ComponentName Then
                AL_Error_Print 2, 3, ComponentName
                AL_Error_Show 2, 3, ComponentName
                Exit Sub
        End If
    Next VBComp

    Select Case ComponentName
        Case "Module"
            ComponentType = 1
        Case "Class"
            ComponentType = 2
        Case "Forms"
            ComponentType = 3
    End Select

    AL_BuildCreateVBComponent ComponentType, ComponentName
    Set VBComp = VBProj.VBComponents(ComponentName)
    Set VBCodeModule = VBComp.CodeModule

    For Each File In Folder.Files
        If File.Name Like "AL_*" Then
            AL_BuildAddCode VBCodeModule, File.Path
        End If
    Next File
    
    For Each SubFolder In Folder.SubFolders
        If SubFolder.Name Like "AL_*" Then
            AL_BuildApplication SubFolder.Path, FolderPath, ComponentType
        End If
    Next SubFolder

End Sub

' Adds Code to a given Component
Sub AL_BuildAddCode(VBCodeModule As VBIDE.CodeModule, FilePath As String)

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
Sub AL_BuildCreateVBComponent(VBProj As VBIDE.VBProject, ComponentType As Integer, ComponentName As String)

    Dim vbComp As VBIDE.VBComponent
    On Error GoTo Error
    Set vbComp = VBProj.VBComponents.Add(ComponentType)
    vbComp.Name = ComponentName
    Exit Sub
    Error:
    MsgBox("Component exists already:" & ComponentName & ", OR" "ComponentType doesnt exist:" ComponentType)

End Sub