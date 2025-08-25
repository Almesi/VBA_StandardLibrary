VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_VBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private VBProj As VBIDE.VBProject
Private TempIncludes() As String
Private p_Handler As std_ErrorHandler

Private Template As String
Private CurrentIdentity As String
Private ErrorValue As Variant

Public Enum VBProjPasteType
    NormalInclude = 0
    NormalReplace = 1
    ForceInclude = 2
End Enum

Public Enum VBProjReferenceType
    FullPath = 0
    Name = 1
    Major = 2
    Minor = 3
    Description = 4
    Typee = 5
    BuiltIn = 6
    GUID = 7
    IsBroken = 8
    All = 9
End Enum


' Public Procedures

    Public Property Let Handler(n_Handler As std_ErrorHandler)
        Set p_Handler = n_Handler
    End Property
    Public Property Get Handler() As std_ErrorHandler
        Set Handler = p_Handler
    End Property

    Public Property Let Project(Project As VBIDE.VBProject)
        Set VBProj = Project
    End Property
    Public Property Get Project() As VBIDE.VBProject
        Set Project = VBProj
    End Property

    Public Function Create(n_Project As VBIDE.VBProject, n_Handler As std_ErrorHandler) As std_VBProject
        Set Create = New std_VBProject
        Create.Project = n_Project
        Create.Handler = n_Handler
    End Function

    Public Function Include(FilePath As String, Optional PasteType As VBProjPasteType = NormalInclude, Optional Temporary As Boolean = False, Optional TryAsInclude As Boolean = False, Optional ComponentName As String = Empty, Optional ComponentType As vbext_ComponentType = vbext_ct_StdModule) As Boolean
    
        Dim Temp() As String
        Dim CompName As String

        If EmptyString(FilePath, "FilePath", "Include") = Handler.IS_ERROR Then Exit Function
        If ValidPasteType(PasteType, "Include") = Handler.IS_ERROR Then Exit Function

        FilePath = ValidateFilePath(FilePath)
        Temp = Split(FilePath, "/")
        CompName = GetCompName(FilePath)

        Select Case True
            Case CheckFileFormat(FilePath) And CheckClassType(FilePath)
                Include = Import(FilePath, CompName, PasteType, Temporary)
            Case FilePath Like "*.frx" And CheckClassType(FilePath) = False
            Case CheckClassType(FilePath) = False
                If EmptyString(ComponentName, "ComponentName", "Include") = Handler.IS_ERROR Then Exit Function
                If ValidComponentType(ComponentType, "Include") = Handler.IS_ERROR Then Exit Function
                If DeleteCode(CompName) <> Handler.IS_ERROR Then
                    Include = AddCodeFromFile(FilePath, CompName, 1, 10, -1)
                End If
            Case TryAsInclude
                If EmptyString(ComponentName, "ComponentName", "Include") = Handler.IS_ERROR Then Exit Function
                If ValidComponentType(ComponentType, "Include") = Handler.IS_ERROR Then Exit Function
                If AddComponent(ComponentName, ComponentType, PasteType) <> Handler.IS_ERROR Then
                    Include = AddCodeFromFile(FilePath, ComponentName, 1, 0, -1)
                End If
                Call AddTemporary(Temporary, CompName)
            Case Else
                Template = "wrong file format. It needs to be .cls, .bas or .frm: VAR0"
                ErrorValue =  FilePath
                GoTo Error
        End Select
        Exit Function

        Error:
        CurrentIdentity = "Include"
        Include = GetError(Handler.IS_ERROR)
    End Function
    
    Public Function Declude(CompName As String) As Boolean

        Dim VBComp As VBIDE.VBComponent
        Dim Found As Boolean

        On Error GoTo Error
        If EmptyString(CompName, "CompName", "Declude") = Handler.IS_ERROR Then Exit Function
        Set VBComp = VBProj.VBComponents(CompName) 
        VBProj.VBComponents.Remove VBComp
        Exit Function

        Error:
        CurrentIdentity = "Declude"
        Template = "could not declude: VAR0"
        If Exists(CompName, VBProj) = False Then Template = "Component doesnt exist: VAR0"
        Declude = GetError(Handler.IS_ERROR)
    End Function
    
    Public Function AddComponent(Name As String, ComponentType As vbext_ComponentType, Optional PasteType As VBProjPasteType = NormalInclude) As Boolean
    
        Dim VBComp As VBComponent
        Dim i As Long

        If ValidComponentType(ComponentType, "AddComponent") = Handler.IS_ERROR Then Exit Function
        If ValidPasteType(PasteType, "AddComponent") = Handler.IS_ERROR Then Exit Function

        If Exists(Name, VBProj) Then
            If PasteType = 2 Then
                On Error GoTo Repeat
                Set VBComp = VBProj.VBComponents.Add(ComponentType)
                Repeat:
                i = i + 1
                VBComp.Name = Name & i
                Exit Function
            ElseIf PasteType = 1 Then
                Exit Function
            Else
                Template = "could not create Component: VAR0"
                ErrorValue = Name
                GoTo Error
            End If
        Else
            Set VBComp = VBProj.VBComponents.Add(ComponentType)
            VBComp.Name = Name
        End If
        Exit Function

        Error:
        CurrentIdentity = "AddComponent"
        AddComponent = GetError(Handler.IS_ERROR)
    End Function
    
    Public Function AddCodeFromFile(FilePath As String, CompName As String, Optional StartLine As Long = 1, Optional FileStart As Long = 0, Optional FileEnd As Long = -1) As Boolean
    
        Dim CodeMod As CodeModule
        Dim Line As String
        Dim FileNumber As Integer
        Dim i As Long

        On Error GoTo Error
        Template = "Component doesnt exist: VAR0"
        ErrorValue = CompName
        Set CodeMod = VbProj.VBComponents(CompName).CodeModule
        FileNumber = FreeFile
        Template = "could not open file: VAR0"
        ErrorValue = FilePath
        Open FilePath For Input As #FileNumber
        Template = "could not add line of code: VAR0"
        Do Until EOF(FileNumber) Or i = FileEnd
            ErrorValue = Line
            Line Input #FileNumber, Line
            If i >= FileStart Then CodeMod.InsertLines StartLine, Line
            StartLine = StartLine + 1
            i = i + 1
        Loop
        Close #FileNumber
        Exit Function
        
        Error:
        CurrentIdentity = "AddCodeFromFile"
        AddCodeFromFile = GetError(Handler.IS_ERROR)
    End Function

    Public Function AddCodeFromComponent(FromComponent As VBComponent, ToComponent As VBComponent, Optional StartLine As Long = 0) As Boolean
    
        On Error GoTo Error
        If StartLine = 0 Then
            Call ToComponent.CodeModule.AddFromString(FromComponent.CodeModule.Lines(1, FromComponent.CodeModule.CountOfLines))
        Else
            Call ToComponent.CodeModule.InsertLines(StartLine, FromComponent.CodeModule.Lines(1, FromComponent.CodeModule.CountOfLines))
        End If
        Exit Function

        Error:
        CurrentIdentity = "AddCodeFromComponent"
        AddCodeFromComponent = GetError(Handler.IS_ERROR)
    End Function

    Public Function AddCodeFromString(CompName As String, Text As String, Optional StartLine As Long = 0) As Boolean

        Dim VBComp As VBComponent
        Set VBComp = Exists(CompName, VBProj, True, False)
        If Not VBComp Is Nothing Then
            If StartLine = 0 Then
                Call VBComp.CodeModule.InsertLines(StartLine, Text)
            Else
                VBComp.CodeModule.AddFromString(Text)
            End If
        Else
            CurrentIdentity = "AddCodeFromString"
            Template = "Component doesnt exist: VAR0"
            ErrorValue = CompName
            AddCodeFromString = GetError(Handler.IS_ERROR)
        End If
    End Function

    Public Function DeleteCode(CompName As String, Optional StartLine As Long = 1, Optional EndLine As Long = 0) As Boolean

        Dim CodeModule As CodeModule

        On Error GoTo Error
        If EmptyString(CompName, "CompName", "DeleteCode") = Handler.IS_ERROR Then Exit Function
        Set CodeModule = VBProj.VBComponents(CompName).CodeModule
        If EndLine = 0 Then EndLine = CodeModule.CountOfLines
        Template = "could not delete code from module: VAR0, VAR1, VAR2"
        CodeModule.DeleteLines StartLine, EndLine
        Exit Function

        Error:
        CurrentIdentity = "DeleteCode"
        DeleteCode = GetError(Handler.IS_ERROR)
    End Function
    
    Public Function IncludeFolder(FolderPath As String, PasteType As VBProjPasteType, Temporary As Boolean, TryAsInclude As Boolean, ParamArray Ignore() As Variant) As Boolean
        Dim Ignores As Variant
        Ignores = Ignore
        IncludeFolder = IncludeFolderArr(FolderPath, PasteType, Temporary, TryAsInclude, Ignores)
    End Function

    Public Function IncludeFolderArr(FolderPath As String, PasteType As VBProjPasteType, Temporary As Boolean, TryAsInclude As Boolean, Ignore As Variant) As Boolean
    
        Dim fso As Object
        Dim Folder As Object
    
        On Error GoTo Error
        If EmptyString(FolderPath, "FolderPath", "InlcudeFolderArr") = Handler.IS_ERROR Then Exit Function
        If ValidPasteType(PasteType, "InlcudeFolderArr") = Handler.IS_ERROR Then Exit Function
        Template = "could not include folder: VAR0"
        ErrorValue = FolderPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set Folder = fso.GetFolder(FolderPath)
        IncludeFolderArr = RecursiveInclude(Folder, PasteType, Temporary, Ignore, TryAsInclude)
        Exit Function

        Error:
        CurrentIdentity = "IncludeFolderArr"
        IncludeFolderArr = GetError(Handler.IS_ERROR)
    End Function
    
    Public Function DecludeFolder(FolderPath As String) As Boolean
    
        Dim VBComp As VBComponent
        Dim fso As Object
        Dim Folder As Object
    
        If EmptyString(FolderPath, "FolderPath", "InlcudeFolderArr") = Handler.IS_ERROR Then Exit Function
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set Folder = fso.GetFolder(FolderPath)
        For Each File In Folder.Files
            For Each VBComp In VBProj.VBComponents
                If GetCompName(File.Path) = VBComp.Name Then DecludeFolder = Declude(VBComp.Name)
            Next
        Next    
        For Each SubFolder In Folder.SubFolders
            DecludeFolder = DecludeFolder(SubFolder.Path)
        Next
        Exit Function

        Error:
        CurrentIdentity = "DecludeFolder"
        DecludeFolder = GetError(Handler.IS_ERROR)
    End Function

    Public Function AddReference(Optional AddFromFile As String = "", Optional AddFromGuid As String = "", Optional Major As Long = 0, Optional Minor As Long = 0) As Boolean
        Dim Ref As VBIDE.Reference

        On Error GoTo Error
        If AddFromFile = "" Then
            If AddFromGuid = "" Then
                Template = "String empty: VAR0, VAR1, VAR2, VAR3"
            ElseIf Major = 0 Or Minor = 0 Then
                Template = "invalid value: VAR0, VAR1, VAR2, VAR3"
            Else
                Template = "could not add reference via AddFromGui: VAR0, VAR1, VAR2, VAR3"
                Set Ref = VBProj.References.AddFromGuid(AddFromGuid, Major, Minor)
            End If
        Else
            Template = "could not add reference via AddFromFile: VAR0, VAR1, VAR2, VAR3"
            Set Ref = VBProj.References.AddFromFile(AddFromFile)
        End If
        Exit Function

        Error:
        CurrentIdentity = "AddReference"
        AddReference = GetError(Handler.IS_ERROR)
    End Function

    Public Function RemoveReference(Value As String, AccordingTo As String) As Boolean
        
        Dim Ref As VBIDE.Reference
        
        On Error GoTo Error
        For Each Ref in VBProj.References
            Select Case UCase(AccordingTo)
                Case "FULLPATH"   : If Ref.FullPath    = Value Then VBProj.References.Remove Ref
                Case "NAME"       : If Ref.Name        = Value Then VBProj.References.Remove Ref
                Case "MAJOR"      : If Ref.Major       = Value Then VBProj.References.Remove Ref
                Case "MINOR"      : If Ref.Minor       = Value Then VBProj.References.Remove Ref
                Case "DESCRIPTION": If Ref.Description = Value Then VBProj.References.Remove Ref
                Case "TYPE"       : If Ref.Type        = Value Then VBProj.References.Remove Ref
                Case "BUILTIN"    : If Ref.BuiltIn     = Value Then VBProj.References.Remove Ref
                Case "GUID"       : If Ref.GUID        = Value Then VBProj.References.Remove Ref
                Case "ISBROKEN"   : If Ref.IsBroken    = Value Then VBProj.References.Remove Ref
            End Select
        Next Ref
        Exit Function

        Error:
        CurrentIdentity = "RemoveReference"
        RemoveReference = GetError(Handler.IS_ERROR)
    End Function

    Public Function GetReferenceProperties(Ref As VBIDE.Reference, Optional Property As VBProjReferenceType = ALL) As String

        If Ref Is Nothing Then GetReferenceProperties = "": Exit Function
        If Property = VBProjReferenceType.ALL Then
            GetReferenceProperties  = Ref.FullPath    & vbcrlf & _ 
                                      Ref.Name        & vbcrlf & _ 
                                      Ref.Major       & vbcrlf & _ 
                                      Ref.Minor       & vbcrlf & _ 
                                      Ref.Description & vbcrlf & _ 
                                      Ref.Type        & vbcrlf & _ 
                                      Ref.BuiltIn     & vbcrlf & _ 
                                      Ref.GUID        & vbcrlf & _ 
                                      Ref.IsBroken    & vbcrlf
        Else
            Select Case UCase(Property)
                Case VBProjReferenceType.FullPath    : GetReferenceProperties = Ref.FullPath
                Case VBProjReferenceType.Name        : GetReferenceProperties = Ref.Name
                Case VBProjReferenceType.Major       : GetReferenceProperties = Ref.Major
                Case VBProjReferenceType.Minor       : GetReferenceProperties = Ref.Minor
                Case VBProjReferenceType.Description : GetReferenceProperties = Ref.Description
                Case VBProjReferenceType.Typee       : GetReferenceProperties = Ref.Type
                Case VBProjReferenceType.BuiltIn     : GetReferenceProperties = Ref.BuiltIn
                Case VBProjReferenceType.GUID        : GetReferenceProperties = Ref.GUID
                Case VBProjReferenceType.IsBroken    : GetReferenceProperties = Ref.IsBroken
            End Select
        End If

    End Function

    Public Function GetAllComponentNames(Optional WB As Workbook = Nothing) As Variant()
        
        Dim WorkProj As VBProject
        Dim VBComp As VBComponent
        Dim ReturnArray() As Variant
        Dim i As Long
        If WB Is Nothing Then
            Set WorkProj = VBProj
        Else
            Set WorkProj = WB.VBPRoject
        End If

        If WorkProj.VBComponents.Count > 0 Then
            ReDim ReturnArray(WorkProj.VBComponents.Count - 1)
            i = 0
            For Each VBComp In WorkProj.VBComponents
                ReturnArray(i) = VBComp.Name
                i = i + 1
            Next VBComp
        Else
            GetAllComponentNames = Empty
        End If
        GetAllComponentNames = ReturnArray

    End Function

    Public Function IncludeWorkbook(WB As Workbook, PasteType As VBProjPasteType, Temporary As Boolean, ParamArray Ignore() As Variant) As Boolean
        Dim Ignores As Variant
        Ignores = Ignore
        IncludeWorkbook = IncludeWorkbookArr(WB, PasteType, Temporary, TryAsInclude, Ignores)
    End Function

    Public Function IncludeWorkbookArr(WB As Workbook, PasteType As VBProjPasteType, Temporary As Boolean, Ignore() As Variant) As Boolean
        Dim WBProject As VBProject
        Dim VBComp As VBComponent
        Dim Existing As Boolean
    
        Set WBProject = WB.VBProject
        For Each VBComp In WBProject.VBComponents
            If Exists(VBComp.Name, Ignore) = False Then
                If AddComponent(VBComp.Name, VBComp.Type, PasteType, False) <> Handler.IS_ERROR Then
                    If PasteType = 1 Then Call DeleteCode(VBComp.Name)
                    Call AddCodeFromComponent(VBComp, VBProj.VBComponents(VBComp.Name))
                    Call AddTemporary(Temporary, VBComp.Name)
                End If
            End If
        Next VBComp
    End Function

    Public Function IncludeWorkbookSingle(WB As Workbook, CompName As String, PasteType As VBProjPasteType, Optional Temporary As Boolean = False) As Boolean
        Dim VBComp As VBComponent
 
        Set VBComp = Exists(CompName, WB.VBProject, True, False)
        If Not VBComp Is Nothing Then
            If AddComponent(VBComp.Name, VBComp.Type, PasteType) <> Handler.IS_ERROR Then
                If PasteType = 1 Then Call DeleteCode(VBComp.Name)
                Call AddCodeFromComponent(VBComp, VBProj.VBComponents(VBComp.Name))
                Call AddTemporary(Temporary, VBComp.Name)
            End If
        Else
            CurrentIdentity = "IncludeWorkbookSingle"
            Template = "could not include component: VAR0"
            ErrorValue = CompName
            IncludeWorkbookSingle = GetError(Handler.IS_ERROR)
        End If
    End Function

    Public Function Exists(Search As Variant, Optional SearchIn As Variant, Optional ShouldExist As Boolean = True, Optional ReturnAsBoolean As Boolean = True) As Variant
        
        Dim Obj As Object
        Dim Found As Boolean
        Dim Text As String
        Dim ReturnVariant As Variant

        If IsMissing(SearchIn) Then Set SearchIn = VBProj
        Select Case True
            Case TypeName(SearchIn) Like "VBProject"
                For Each Obj In SearchIn.VBComponents
                    If Obj.Name = Search Then
                        Found = True
                        Set ReturnVariant = Obj
                        Exit For
                    End If
                Next
            Case TypeName(SearchIn) Like "VBComponent"
            Case TypeName(SearchIn) Like "CodeModule"
                Text = SearchIn.Lines(1, SearchIn.CoundOfLines)
                Found = Text Like "*" & Search & "*"
                ReturnVariant = InStr(1, Text, Search)
            Case IsArray(SearchIn)
                Dim i As Long
                For i = 0 To UBound(SearchIn)
                    If SearchIn(i) = Search Then
                        Found = True
                        ReturnVariant = SearchIn(i)
                        Exit For
                    End If
                Next
            Case Else
        End Select

        If ReturnAsBoolean Then
            Exists = Not (Found Xor ShouldExist)
        Else
            If IsObject(ReturnVariant) Then
                Set Exists = ReturnVariant
            Else
                Exists = ReturnVariant
            End If
        End If

    End Function

    Public Function GetFileString(FilePath As String) As String
        Dim Line As String
        Dim FileNumber As Integer

        On Error GoTo Error
        FileNumber = FreeFile
        Template = "could not open file: VAR0"
        Open FilePath For Input As #FileNumber
        Template = "could not extract line: VAR0"
        Do Until EOF(FileNumber)
            Line Input #FileNumber, Line
            ErrorValue = Line
        Loop
        Close #FileNumber
        GetFileString = Line
        Exit Function
        Error:
        CurrentIdentity = ""
        Call GetError(Handler.IS_ERROR)
        GetFileString = Empty
    End Function
'

' Private Procedures
    Private Function Import(FilePath As String, ComponentName As String, Optional PasteType As VBProjPasteType = NormalInclude, Optional Temporary As Boolean = False) As Boolean

        Dim VBComp As VBIDE.VBComponent
        Dim Found As Boolean
        Dim Temp As Boolean

        On Error GoTo Error
        Template = "Could not import: VAR0"
        ErrorValue = FilePath
        Found = Exists(ComponentName)
        Select Case PasteType
            Case 0
                If Found Then
                    Template = "Component already exists: VAR0"
                    ErrorValue = ComponentName
                    GoTo Error
                Else
                    VBProj.VBComponents.Import FilePath
                End If
            Case 1
                If Found Then
                    Set VBComp = Exists(ComponentName,,,False)
                    VBComp.Name = ComponentName & "_TODELETE"
                    If Declude(ComponentName & "_TODELETE") <> Handler.IS_ERROR Then
                        VBProj.VBComponents.Import FilePath
                    End If
                Else
                    VBProj.VBComponents.Import FilePath
                End If
            Case 2
                VBProj.VBComponents.Import FilePath
        End Select
        If Temporary Then
            ReDim Preserve TempIncludes(Ubound(TempIncludes) + 1)
            TempIncludes(Ubound(TempIncludes)) = ComponentName
        End If
        Exit Function

        Error:
        CurrentIdentity = "Import"
        Import = GetError(Handler.IS_ERROR)
    End Function

    Private Function RecursiveInclude(Folder As Object, PasteType As VBProjPasteType, Temporary As Boolean, Ignore As Variant, Optional TryAsInclude As Boolean = False) As Boolean

        Dim File As Object
        Dim SubFolder As Object
        Dim Element As Variant
        Dim CompName As String
        Dim CompType As Long
        Dim Skip As Boolean

        If Ubound(Ignore) < 0 Then ReDim Ignore(0)
        For Each SubFolder In Folder.SubFolders        
            For Each Element In Ignore
                If SubFolder.Name = Cstr(Element) Then Skip = True: Exit For
            Next
        Next
        For Each File In Folder.Files
            For Each Element In Ignore
                If File.Name = Cstr(Element) Then Skip = True: Exit For
            Next
            If Skip Then
                Skip = False
            Else
                CompName = GetCompName(File.Path)
                CompType = vbext_ComponentType.vbext_ct_StdModule
                If Include(File.Path, PasteType, Temporary, TryAsInclude, CompName, CompType) = Handler.IS_ERROR Then RecursiveInclude = Handler.IS_ERROR
                CompName = ""
                CompType = 0
            End If
        Next
        Skip = False 
        For Each SubFolder In Folder.SubFolders
            For Each Element In Ignore
                If SubFolder.Name = Cstr(Element) Then Skip = True: Exit For
            Next
            If Skip Then
                Skip = False
            Else
            RecursiveInclude = RecursiveInclude(SubFolder, PasteType, Temporary, Ignore, TryAsInclude)
            End If
        Next

    End Function

    Private Function GetCompName(FilePath As String) As String
        Dim Temp() As String
        Dim Name As String

        FilePath = ValidateFilePath(FilePath)
        Temp = Split(FilePath, "/")
        Name = Temp(Ubound(Temp))
        Name = Replace(Replace(Replace(Name, ".bas", ""), ".cls", ""), ".frm", "")
        GetCompName = Name
    End Function

    Private Function CheckFileFormat(FilePath As String) As Boolean
        If FilePath Like "*.cls" Or _
           FilePath Like "*.bas" Or _
           FilePath Like "*.frm"    _
           Then CheckFileFormat = True
    End Function

    Private Function CheckClassType(FilePath As String) As Boolean
        If Not FilePath Like "*Sheet#.cls"   And _
           Not FilePath Like "*Sheet##.cls"  And _
           Not FilePath Like "*Sheet###.cls" And _
           Not FilePath Like "*.frx"         And _
           Not FilePath Like "*ThisWorkbook.cls" Then CheckClassType = True
    End Function

    Private Function ValidateFilePath(FilePath As String) As String
        ValidateFilePath = Replace(FilePath, "\", "/")
    End Function

    Public Sub AddTemporary(Temporary As Boolean, CompName As String)
        If Temporary Then
                ReDim Preserve TempIncludes(Ubound(TempIncludes) + 1)
                TempIncludes(Ubound(TempIncludes)) = CompName
        End If
    End Sub

    Private Function EmptyString(Text As String, VariableName As String, Identity As String) As Boolean
        CurrentIdentity = Identity
        ErrorValue = VariableName
        Template = "Variable [VAR0] is Empty"
        EmptyString = GetError(Text = Empty)
    End Function

    Private Function ValidPasteType(PasteType As VBProjPasteType, Identity As String) As Boolean
        Dim Condition As Boolean
        Condition = (PasteType < 0 Or PasteType > 2)
        CurrentIdentity = Identity
        ErrorValue = PasteType
        Template = "PasteType [VAR0] is not a valid option"
        ValidPasteType = GetError(Condition)
    End Function

    Private Function ValidComponentType(ComponentType As vbext_ComponentType, Identity As String) As Boolean
        Dim Condition As Boolean
        Condition = (ComponentType <> 1 And ComponentType <> 2 And ComponentType <> 3 And ComponentType <> 11 And ComponentType <> 100)
        CurrentIdentity = Identity
        ErrorValue = ComponentType
        Template = "ComponentType [VAR0] is not a valid option"
        ValidComponentType = GetError(Condition)
    End Function

    Private Function GetError(Condition As Boolean) As Boolean
        Dim CurrentError As std_Error
        Set CurrentError = std_Error.Create("std_VBProject", "major", CurrentIdentity, Template, Empty, ErrorValue)
        GetError = Handler.Handle(Condition, CurrentError)
    End Function

    Private Sub Class_Initialize()
        ReDim TempIncludes(0)
    End Sub

    Private Sub Class_Terminate()
        Dim Temp As Boolean
        Dim i As Long
        For i = 1 To Ubound(TempIncludes)
            Temp = Declude(TempIncludes(i))
        Next
    End Sub