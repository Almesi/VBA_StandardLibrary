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
Private ErrorCatalog(1, 99) As Variant
Private TempIncludes() As String
Private Handler As New std_Error

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

    Public Property Set SetHandler(n_Handler As std_Error)
        Set Handler = n_Handler
    End Function

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = Handler.IS_ERROR
    End Property

    Public Function Create(Project As VBIDE.VBProject) As std_VBProject
        Set Create = New std_VBProject
        Set Create.Project = Project
    End Function

    Public Property Set Project(Project As VBIDE.VBProject)
        Set VBProj = Project
    End Property

    Public Function Include(FilePath As String, Optional PasteType As VBProjPasteType = NormalInclude, Optional Temporary As Boolean = False, Optional TryAsInclude As Boolean = False, Optional ComponentName As String = Empty, Optional ComponentType As vbext_ComponentType = vbext_ct_StdModule, Optional ThrowError As Boolean = True) As Boolean
    
        Dim Temp() As String
        Dim CompName As String
        Dim ErrorCode As Long
        Dim ErrorValue As Variant

        If FilePath = Empty Then ErrorCode = 5: ErrorValue = FilePath: GoTo Error
        If PasteType <> 0 And PasteType <> 1 And PasteType <> 2 Then ErrorCode = 25: ErrorValue = PasteType: GoTo Error

        FilePath = ValidateFilePath(FilePath)
        Temp = Split(FilePath, "/")
        CompName = GetCompName(FilePath)

        If CheckFileFormat(FilePath) Then
                Include = Import(FilePath, CompName, PasteType, Temporary, ThrowError)
        ElseIf FilePath Like "*.frx" Then
        ElseIf TryAsInclude Then
            If ComponentName = Empty Then ErrorCode = 18: ErrorValue = ComponentName: GoTo Error
            If ComponentType = 0     Then ErrorCode = 20: ErrorValue = ComponentType: GoTo Error
            If AddComponent(ComponentName, ComponentType, PasteType, ThrowError) <> Handler.IS_Error Then
                Include = AddCodeFromFile(FilePath, ComponentName, 1, ThrowError)
            End If
            Call AddTemporary(Temporary, CompName)
        Else
            ErrorCode = 16: ErrorValue = FilePath: GoTo Error
        End If
        Exit Function

        Error:
        Include = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, ComponentType)
    
    End Function
    
    Public Function Declude(CompName As String, Optional ThrowError As Boolean = True) As Boolean

        Dim VBComp As VBIDE.VBComponent
        Dim Found As Boolean
        Dim ErrorCode As Long

        On Error GoTo Error
        If CompName = Empty           Then ErrorCode = 05: GoTo Error
        If CompName = "std_VBProject" Then ErrorCode = 17: GoTo Error
        If CompName = "Handler"       Then ErrorCode = 17: GoTo Error
        Set VBComp = VBProj.VBComponents(CompName) 
        VBProj.VBComponents.Remove VBComp
        Exit Function

        Error:
        ErrorCode = 14
        If Exists(CompName, VBProj) = False Then ErrorCode = 12
        Declude = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, CompName)
    
    End Function
    
    Public Function AddComponent(Name As String, ComponentType As vbext_ComponentType, Optional PasteType As VBProjPasteType = NormalInclude, Optional ThrowError As Boolean = True) As Boolean
    
        Dim VBComp As VBComponent
        Dim i As Long
        Dim ErrorValue As Variant
        Dim ErrorCode As Long

        If ComponentType <> 1 And ComponentType <> 2 And ComponentType <> 3 And ComponentType <> 11 And ComponentType <> 100 Then ErrorCode = 20: ErrorValue = Name: GoTo Error
        If PasteType <> 0 And PasteType <> 1 And PasteType <> 2 Then ErrorCode = 25: ErrorValue = PasteType: GoTo Error 

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
                If ThrowError Then
                    AddComponent = ErrorCode = 19: ErrorValue = Name: GoTo Error
                Else
                    AddComponent = Handler.IS_ERROR
                End If
            End If
        Else
            Set VBComp = VBProj.VBComponents.Add(ComponentType)
            VBComp.Name = Name
        End If
        Exit Function

        Error:
        AddComponent = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, ErrorValue)
    
    End Function
    
    Public Function AddCodeFromFile(FilePath As String, CompName As String, Optional StartLine As Long = 1, Optional ThrowError As Boolean = True) As Boolean
    
        Dim CodeMod As CodeModule
        Dim Line As String
        Dim FileNumber As Integer
        Dim ErrorCode As Long

        On Error GoTo Error
        ErrorCode = 12
        Set CodeMod = VbProj.VBComponents(CompName).CodeModule
        FileNumber = FreeFile
        ErrorCode = 21
        Open FilePath For Input As #FileNumber
        ErrorCode = 22
        Do Until EOF(FileNumber)
            Line Input #FileNumber, Line
            CodeMod.InsertLines StartLine, Line
            StartLine = StartLine + 1
        Loop
        Close #FileNumber
        Exit Function
        Error:
        AddCodeFromFile = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, CompName)

    End Function

    Public Function AddCodeFromComponent(FromComponent As VBComponent, ToComponent As VBComponent, Optional StartLine As Long = 0, Optional ThrowError As Boolean = True) As Boolean
    
        On Error GoTo Error
        If StartLine = 0 Then
            Call ToComponent.CodeModule.AddFromString(FromComponent.CodeModule.Lines(1, FromComponent.CodeModule.CountOfLines))
        Else
            Call ToComponent.CodeModule.InsertLines(StartLine, FromComponent.CodeModule.Lines(1, FromComponent.CodeModule.CountOfLines))
        End If
        Exit Function

        Error:
        AddCodeFromComponent = Handler.Handle(ErrorCatalog, 28, ThrowError, FromComponent.Name)

    End Function

    Public Function AddCodeFromString(CompName As String, Text As String, Optional StartLine As Long = 0, Optional ThrowError As Boolean = True) As Boolean

        Dim VBComp As VBComponent
        Set VBComp = Exists(CompName, VBProj, True, False)
        If Not VBComp Is Nothing Then
            If StartLine = 0 Then
                Call VBComp.CodeModule.InsertLines(StartLine, Text)
            Else
                VBComp.CodeModule.AddFromString(Text)
            End If
        Else
            AddCodeFromString = Handler.Handle(ErrorCatalog, 12, ThrowError, CompName, Text)
        End If

    End Function

    Public Function DeleteCode(CompName As String, Optional StartLine As Long = 1, Optional EndLine As Long = 0, Optional ThrowError As Boolean = True) As Boolean

        Dim CodeModule As CodeModule
        Dim ErrorCode As Long

        On Error GoTo Error
        If CompName = Empty Then DeleteCode = ErrorCode = 5: GoTo Error
        Set CodeModule = VBProj.VBComponents(CompName).CodeModule
        If EndLine = 0 Then EndLine = CodeModule.CountOfLines
        ErrorCode = 27
        CodeModule.DeleteLines StartLine, EndLine
        Exit Function

        Error:
        DeleteCode = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, CompName, StartLine, EndLine)

    End Function
    
    Public Function IncludeFolder(FolderPath As String, PasteType As VBProjPasteType, Temporary As Boolean, TryAsInclude As Boolean, ThrowError As Boolean, ParamArray Ignore() As Variant) As Boolean
        Dim Ignores As Variant
        Ignores = Ignore
        IncludeFolder = IncludeFolderArr(FolderPath, PasteType, Temporary, TryAsInclude, ThrowError, Ignores)
    End Function

    Public Function IncludeFolderArr(FolderPath As String, PasteType As VBProjPasteType, Temporary As Boolean, TryAsInclude As Boolean, ThrowError As Boolean, Ignore As Variant) As Boolean
    
        Dim fso As Object
        Dim Folder As Object
        Dim ErrorCode As Long
        Dim ErrorValue As Variant
    
        On Error GoTo Error
        If FolderPath = Empty Then ErrorCode = 5: ErrorValue = FolderPath: GoTo Error
        If PasteType <> 0 And PasteType <> 1 And PasteType <> 2 Then ErrorCode = 25: ErrorValue = PasteType: GoTo Error
        ErrorCode = 29
        ErrorValue = FolderPath
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set Folder = fso.GetFolder(FolderPath)
        IncludeFolderArr = RecursiveInclude(Folder, PasteType, Temporary, Ignore, TryAsInclude, ThrowError)
        Exit Function

        Error:
        IncludeFolderArr = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, ErrorValue): Exit Function
    
    End Function
    
    Public Function DecludeFolder(FolderPath As String, Optional ThrowError As Boolean = True) As Boolean
    
        Dim VBComp As VBComponent
        Dim fso As Object
        Dim Folder As Object
        Dim ErrorCode As Long
    
        If FolderPath = Empty Then ErrorCode = 5: GoTo Error
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set Folder = fso.GetFolder(FolderPath)
        For Each File In Folder.Files
            For Each VBComp In VBProj.VBComponents
                If GetCompName(File.Path) = VBComp.Name Then DecludeFolder = Declude(VBComp.Name)
            Next
        Next    
        For Each SubFolder In Folder.SubFolders
            DecludeFolder = DecludeFolder(SubFolder.Path, ThrowError)
        Next
        Exit Function

        Error:
        DecludeFolder = Handler.Handle(ErrorCatalog, 5, ThrowError, FolderPath)
    
    End Function

    Public Function AddReference(Optional AddFromFile As String = "", Optional AddFromGuid As String = "", Optional Major As Long = 0, Optional Minor As Long = 0, Optional ThrowError As Boolean = True) As Boolean

        Dim ErrorCode As Long
        Dim Ref As VBIDE.Reference

        On Error GoTo Error
        If AddFromFile = "" Then
            If AddFromGuid = "" Then
                ErrorCode = 5
            ElseIf Major = 0 Or Minor = 0 Then
                ErrorCode = 25
            Else
                ErrorCode = 24
                Set Ref = VBProj.References.AddFromGuid(AddFromGuid, Major, Minor)
            End If
        Else
            ErrorCode = 23
            Set Ref = VBProj.References.AddFromFile(AddFromFile)
        End If
        Exit Function

        Error:
        AddReference = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, AddFromFile, AddFromGuid, Major, Minor)

    End Function

    Public Function RemoveReference(Value As String, AccordingTo As String, Optional ThrowError As Boolean = True) As Boolean
        
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
        RemoveReference = Handler.Handle(ErrorCatalog, 26, ThrowError, Ref.FullPath, Ref.Name, Ref.Major, Ref.Minor, Ref.Description, Ref.Type, Ref.BuiltIn, Ref.GUID, Ref.IsBroken)
    
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

    Public Function IncludeWorkbook(WB As Workbook, PasteType As VBProjPasteType, Temporary As Boolean, ThrowError As Boolean, ParamArray Ignore() As Variant) As Boolean
        Dim Ignores As Variant
        Ignores = Ignore
        IncludeWorkbook = IncludeWorkbookArr(WB, PasteType, Temporary, TryAsInclude, ThrowError, Ignores)
    End Function

    Public Function IncludeWorkbookArr(WB As Workbook, PasteType As VBProjPasteType, Temporary As Boolean, ThrowError As Boolean, Ignore() As Variant) As Boolean
        Dim WBProject As VBProject
        Dim VBComp As VBComponent
        Dim Existing As Boolean
    
        Set WBProject = WB.VBProject
        For Each VBComp In WBProject.VBComponents
            If Exists(VBComp.Name, Ignore) = False Then
                If AddComponent(VBComp.Name, VBComp.Type, PasteType, False, ThrowError) <> Handler.IS_ERROR Then
                    If PasteType = 1 Then Call DeleteCode(VBComp.Name, ThrowError)
                    Call AddCodeFromComponent(VBComp, VBProj.VBComponents(VBComp.Name), ThrowError)
                    Call AddTemporary(Temporary, VBComp.Name, ThrowError)
                End If
            End If
        Next VBComp
    End Function

    Public Function IncludeWorkbookSingle(WB As Workbook, CompName As String, PasteType As VBProjPasteType, Optional Temporary As Boolean = False, Optional ThrowError As Boolean = True) As Boolean
        Dim VBComp As VBComponent

        Set VBComp = Exists(CompName, WB.VBProject, True, False)
        If Not VBComp Is Nothing Then
            If AddComponent(VBComp.Name, VBComp.Type, PasteType, False, ThrowError) <> Handler.IS_ERROR Then
                If PasteType = 1 Then Call DeleteCode(VBComp.Name, ThrowError)
                Call AddCodeFromComponent(VBComp, VBProj.VBComponents(VBComp.Name), ThrowError)
                Call AddTemporary(Temporary, VBComp.Name, ThrowError)
            End If
        Else
            IncludeWorkbookSingle = Handler.Handle(ErrorCatalog, 28, ThrowError, WB.Name, CompName, PasteType, Temporary)
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
'

' Private Procedures
    Private Function Import(FilePath As String, ComponentName As String, Optional PasteType As VBProjPasteType = NormalInclude, Optional Temporary As Boolean = False, Optional ThrowError As Boolean = True) As Boolean

        Dim VBComp As VBIDE.VBComponent
        Dim Found As Boolean
        Dim Temp As Boolean
        Dim ErrorCode As Long
        Dim ErrorValue As Variant

        On Error GoTo Error
        ErrorCode = 1: ErrorValue = FilePath
        If CheckClassType(FilePath) = False Then: GoTo Error
        Found = Exists(ComponentName)
        Select Case PasteType
            Case 0
                If Found Then
                    ErrorCode = 11: ErrorValue = ComponentName: GoTo Error
                Else
                    VBProj.VBComponents.Import FilePath
                End If
            Case 1
                If Found Then
                    Set VBComp = Exists(ComponentName,,,False)
                    VBComp.Name = ComponentName & "_TODELETE"
                    If Declude(ComponentName & "_TODELETE", ThrowError) <> Handler.IS_ERROR Then
                        VBProj.VBComponents.Import FilePath
                    End If
                Else
                    VBProj.VBComponents.Import FilePath
                End If
            Case 2
                VBProj.VBComponents.Import FilePath
            Case Else
                ErrorCode = 25: ErrorValue = PasteType: GoTo Error
        End Select
        If Temporary Then
            ReDim Preserve TempIncludes(Ubound(TempIncludes) + 1)
            TempIncludes(Ubound(TempIncludes)) = ComponentName
        End If
        Exit Function

        Error:
        Import = Handler.Handle(ErrorCatalog, ErrorCode, ThrowError, ErrorValue)

    End Function

    Private Function RecursiveInclude(Folder As Object, PasteType As VBProjPasteType, Temporary As Boolean, Ignore As Variant, Optional TryAsInclude As Boolean = False, Optional ThrowError As Boolean = True) As Boolean

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
                If Include(File.Path, PasteType, Temporary, TryAsInclude, CompName, CompType, ThrowError) = Handler.IS_Error Then RecursiveInclude = Handler.IS_Error
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
            RecursiveInclude = RecursiveInclude(SubFolder, PasteType, Temporary, Ignore, TryAsInclude, ThrowError)
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

    Private Sub Class_Initialize()
        ProtInit
        ReDim TempIncludes(0)
    End Sub

    Private Sub Class_Terminate()
        Dim Temp As Boolean
        Dim i As Long
        For i = 1 To Ubound(TempIncludes)
            Temp = Declude(TempIncludes(i))
        Next
    End Sub

    Private Sub ProtInit()
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_VBProject"
        ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Could not import"
        ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Could import"
        ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Could not include file"
        ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Could include file"
        ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "String empty"
        ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "String invalid"
        ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "could not import temporary"
        ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "could import temporary"
        ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "could not include temporary file"
        ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "could include temporary file"
        ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "Component already exists"
        ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "Component doesnt exist"
        ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "Unknown error"
        ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "could not declude"
        ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "could not declude file"
        ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "wrong file format. It needs to be .cls, .bas or .frm"
        ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "This Component cant be removed, as this code needs it to run"
        ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "Could not include file, no ComponentName"
        ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "could not create Component"
        ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "componentType is invalid"
        ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "could not open file"
        ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "could not add line of code"
        ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "could not add reference via AddFromFile"
        ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "could not add reference via AddFromGui"
        ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "invalid value"
        ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "could not remove reference"
        ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "could not delete code from module"
        ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "could not include component"
        ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "could not include folder"
        ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "PLACEHOLDER"
        ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "PLACEHOLDER"
        ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
        ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
        ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
        ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "PLACEHOLDER"
        ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "PLACEHOLDER"
        ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "PLACEHOLDER"
        ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "PLACEHOLDER"
        ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "PLACEHOLDER"
        ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "PLACEHOLDER"
        ErrorCatalog(0, 0041) = 1000: ErrorCatalog(1, 0041) = "PLACEHOLDER"
        ErrorCatalog(0, 0042) = 1000: ErrorCatalog(1, 0042) = "PLACEHOLDER"
        ErrorCatalog(0, 0043) = 1000: ErrorCatalog(1, 0043) = "PLACEHOLDER"
        ErrorCatalog(0, 0044) = 1000: ErrorCatalog(1, 0044) = "PLACEHOLDER"
        ErrorCatalog(0, 0045) = 1000: ErrorCatalog(1, 0045) = "PLACEHOLDER"
        ErrorCatalog(0, 0046) = 1000: ErrorCatalog(1, 0046) = "PLACEHOLDER"
        ErrorCatalog(0, 0047) = 1000: ErrorCatalog(1, 0047) = "PLACEHOLDER"
        ErrorCatalog(0, 0048) = 1000: ErrorCatalog(1, 0048) = "PLACEHOLDER"
        ErrorCatalog(0, 0049) = 1000: ErrorCatalog(1, 0049) = "PLACEHOLDER"
        ErrorCatalog(0, 0050) = 1000: ErrorCatalog(1, 0050) = "PLACEHOLDER"
        ErrorCatalog(0, 0051) = 1000: ErrorCatalog(1, 0051) = "PLACEHOLDER"
        ErrorCatalog(0, 0052) = 1000: ErrorCatalog(1, 0052) = "PLACEHOLDER"
        ErrorCatalog(0, 0053) = 1000: ErrorCatalog(1, 0053) = "PLACEHOLDER"
        ErrorCatalog(0, 0054) = 1000: ErrorCatalog(1, 0054) = "PLACEHOLDER"
        ErrorCatalog(0, 0055) = 1000: ErrorCatalog(1, 0055) = "PLACEHOLDER"
        ErrorCatalog(0, 0056) = 1000: ErrorCatalog(1, 0056) = "PLACEHOLDER"
        ErrorCatalog(0, 0057) = 1000: ErrorCatalog(1, 0057) = "PLACEHOLDER"
        ErrorCatalog(0, 0058) = 1000: ErrorCatalog(1, 0058) = "PLACEHOLDER"
        ErrorCatalog(0, 0059) = 1000: ErrorCatalog(1, 0059) = "PLACEHOLDER"
        ErrorCatalog(0, 0060) = 1000: ErrorCatalog(1, 0060) = "PLACEHOLDER"
        ErrorCatalog(0, 0061) = 1000: ErrorCatalog(1, 0061) = "PLACEHOLDER"
        ErrorCatalog(0, 0062) = 1000: ErrorCatalog(1, 0062) = "PLACEHOLDER"
        ErrorCatalog(0, 0063) = 1000: ErrorCatalog(1, 0063) = "PLACEHOLDER"
        ErrorCatalog(0, 0064) = 1000: ErrorCatalog(1, 0064) = "PLACEHOLDER"
        ErrorCatalog(0, 0065) = 1000: ErrorCatalog(1, 0065) = "PLACEHOLDER"
        ErrorCatalog(0, 0066) = 1000: ErrorCatalog(1, 0066) = "PLACEHOLDER"
        ErrorCatalog(0, 0067) = 1000: ErrorCatalog(1, 0067) = "PLACEHOLDER"
        ErrorCatalog(0, 0068) = 1000: ErrorCatalog(1, 0068) = "PLACEHOLDER"
        ErrorCatalog(0, 0069) = 1000: ErrorCatalog(1, 0069) = "PLACEHOLDER"
        ErrorCatalog(0, 0070) = 1000: ErrorCatalog(1, 0070) = "PLACEHOLDER"
        ErrorCatalog(0, 0071) = 1000: ErrorCatalog(1, 0071) = "PLACEHOLDER"
        ErrorCatalog(0, 0072) = 1000: ErrorCatalog(1, 0072) = "PLACEHOLDER"
        ErrorCatalog(0, 0073) = 1000: ErrorCatalog(1, 0073) = "PLACEHOLDER"
        ErrorCatalog(0, 0074) = 1000: ErrorCatalog(1, 0074) = "PLACEHOLDER"
        ErrorCatalog(0, 0075) = 1000: ErrorCatalog(1, 0075) = "PLACEHOLDER"
        ErrorCatalog(0, 0076) = 1000: ErrorCatalog(1, 0076) = "PLACEHOLDER"
        ErrorCatalog(0, 0077) = 1000: ErrorCatalog(1, 0077) = "PLACEHOLDER"
        ErrorCatalog(0, 0078) = 1000: ErrorCatalog(1, 0078) = "PLACEHOLDER"
        ErrorCatalog(0, 0079) = 1000: ErrorCatalog(1, 0079) = "PLACEHOLDER"
        ErrorCatalog(0, 0080) = 1000: ErrorCatalog(1, 0080) = "PLACEHOLDER"
        ErrorCatalog(0, 0081) = 1000: ErrorCatalog(1, 0081) = "PLACEHOLDER"
        ErrorCatalog(0, 0082) = 1000: ErrorCatalog(1, 0082) = "PLACEHOLDER"
        ErrorCatalog(0, 0083) = 1000: ErrorCatalog(1, 0083) = "PLACEHOLDER"
        ErrorCatalog(0, 0084) = 1000: ErrorCatalog(1, 0084) = "PLACEHOLDER"
        ErrorCatalog(0, 0085) = 1000: ErrorCatalog(1, 0085) = "PLACEHOLDER"
        ErrorCatalog(0, 0086) = 1000: ErrorCatalog(1, 0086) = "PLACEHOLDER"
        ErrorCatalog(0, 0087) = 1000: ErrorCatalog(1, 0087) = "PLACEHOLDER"
        ErrorCatalog(0, 0088) = 1000: ErrorCatalog(1, 0088) = "PLACEHOLDER"
        ErrorCatalog(0, 0089) = 1000: ErrorCatalog(1, 0089) = "PLACEHOLDER"
        ErrorCatalog(0, 0090) = 1000: ErrorCatalog(1, 0090) = "PLACEHOLDER"
        ErrorCatalog(0, 0091) = 1000: ErrorCatalog(1, 0091) = "PLACEHOLDER"
        ErrorCatalog(0, 0092) = 1000: ErrorCatalog(1, 0092) = "PLACEHOLDER"
        ErrorCatalog(0, 0093) = 1000: ErrorCatalog(1, 0093) = "PLACEHOLDER"
        ErrorCatalog(0, 0094) = 1000: ErrorCatalog(1, 0094) = "PLACEHOLDER"
        ErrorCatalog(0, 0095) = 1000: ErrorCatalog(1, 0095) = "PLACEHOLDER"
        ErrorCatalog(0, 0096) = 1000: ErrorCatalog(1, 0096) = "PLACEHOLDER"
        ErrorCatalog(0, 0097) = 1000: ErrorCatalog(1, 0097) = "PLACEHOLDER"
        ErrorCatalog(0, 0098) = 1000: ErrorCatalog(1, 0098) = "PLACEHOLDER"
        ErrorCatalog(0, 0099) = 1000: ErrorCatalog(1, 0099) = "PLACEHOLDER"
    End Sub
'