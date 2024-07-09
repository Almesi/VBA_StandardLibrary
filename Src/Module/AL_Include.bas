Attribute VB_Name = "AL_Include"

' Adds Code to a given Component
Public Function AL_Include_AddCode(VBCodeModule As VBIDE.CodeModule, FilePath As String) As Boolean

    If AL_Error_Var(FilePath) = AL_IS_ERROR Then
        AL_Include_AddCode = AL_IS_ERROR
        Exit Function
    End If
    If AL_Error_Obj(VBCodeModule) = AL_IS_ERROR Then
        AL_Include_AddCode = AL_IS_ERROR
        Exit Function
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
    AL_Include_AddCode = AL_NO_ERROR

End Function

    ' Creates a Component (Module, Class, Userform)
Public Function AL_Include_AddComponent(VBProj As VBIDE.VBProject, ComponentType, ComponentName As String) As VBIDE.VBComponent

    Dim VBComp As VBIDE.VBComponent

    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner, vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_MSForm, vbext_ct_StdModule
            If AL_Error_Component(VBProj, AL_Exist, ComponentName) = AL_IS_ERROR Then
                AL_Error_Show AL_Error_Workbook, 0008, ComponentName
                Exit Function
            End If
            Set VBComp = VBProj.VBComponents.Add(ComponentType)
            VBComp.Name = ComponentName
        Case Else
            AL_Error_Show AL_Error_System, 0001, ComponentType
            Exit Function
    End Select
    Set AL_Include_AddComponent = VBComp

End Function

    Private Function AL_Include_CheckDependencies(FilePath As String) As Boolean

        If AL_Check_String(FilePath, 0) = Is_Error Then
            Exit Sub
        End If
        Dim FileLine As String
        Dim Index As Long
        Dim FileNumber As Integer
        Dim Included As Boolean
        Dim vbProj As VBIDE.VBProject
        Dim vbComp As VBIDE.VBComponent
    
        Set vbProj = ThisWorkbook.VBProject
    
        Index = 1
        FileNumber = FreeFile
        Open FilePath For Input As #FileNumber
    
        Do Until EOF(FileNumber)
            Line Input #FileNumber, FileLine
            For Each vbComp in vbProj.VBComponents
                    If vbComp.Name = FileLine Then
                        Included = True
                    End If
            Next vbComp
            If Included = False Then
                    AL_Error_Show AL_Error_Workbook, 0003, FileLine
                    AL_Include_CheckDependencies = True
                    Exit Function
                Else
                    Included = False
            End If
            Index = Index + 1
        Loop
        Close #FileNumber
        AL_Include_CheckDependencies = False
    
    End Function

        ' Loops through Folders (Create Module per folder)
' Loops through Files and adds Code to Module
Private Sub AL_Include_Folder(VBProj As VBIDE.VBProject, FolderPath As String, ComponentType As Long, ComponentName As String)

    Dim fso As Object
    Dim Folder As Object
    Dim SubFolder As Object
    Dim File As Object
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

    ' Gets all information needed to Include a Folder defined in BuildData
Private Function AL_Include_GetData(Folder As Object) As Variant()

    Dim File As Object
    
    Dim FileArray() As Variant
    For Each File In Folder.Files
        If File.Name = AL_Include_BuildData Then
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
            FileArray(i) = CLng(Replace(FileLine, AL_Include_ComponentType, ""))
            Select Case FileArray(i)
                Case vbext_ct_ActiveXDesigner
                Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document, vbext_ct_MSForm
                    Index = Index + 1
                    i = i + 1
                    Line Input #FileNumber, FileLine
                    FileArray(i) = Replace(FileLine, AL_Include_ComponentName, "")
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