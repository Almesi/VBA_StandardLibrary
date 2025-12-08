VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Include"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit


Private p_Destination  As VBIDE.VBProject
Private p_Handler      As std_ErrorHandler
Private p_Recursive    As Boolean ' Will loop through all subfolder to get every file.
Private p_Log          As Boolean ' Will print every component included from its source
Private p_Overwrite    As Boolean ' Will change Code of found Component
Private p_Increment    As Boolean ' Will ignore that it found Component and will include again
Private p_Temporary    As Boolean ' Will remove the added component once the object is destroyed

Private Formats() As IIncludeFormat

Private ToBuildFormats() As IIncludeFormat
Private ToBuildSources() As std_IncludeSource

Private TempIncludes() As VBIDE.VBComponent


Public Enum VBProjReferenceType
    VBProjReferenceTypeFullPath = 0
    VBProjReferenceTypeName = 1
    VBProjReferenceTypeMajor = 2
    VBProjReferenceTypeMinor = 3
    VBProjReferenceTypeDescription = 4
    VBProjReferenceTypeTypee = 5
    VBProjReferenceTypeBuiltIn = 6
    VBProjReferenceTypeGUID = 7
    VBProjReferenceTypeIsBroken = 8
    VBProjReferenceTypeAll = 9
End Enum


' Factory

    Public Property Let Destination(n_Destination As VBIDE.VBProject)  : Set p_Destination = n_Destination : End Property
    Public Property Let Handler(n_Handler         As std_ErrorHandler) : Set p_Handler     = n_Handler     : End Property
    Public Property Let Recursive(n_Recursive     As Boolean)          : Let p_Recursive   = n_Recursive   : End Property
    Public Property Let Log(n_Log                 As Boolean)          : Let p_Log         = n_Log         : End Property
    Public Property Let Overwrite(n_Overwrite     As Boolean)          : Let p_Overwrite   = n_Overwrite   : End Property
    Public Property Let Increment(n_Increment     As Boolean)          : Let p_Increment   = n_Increment   : End Property
    Public Property Let Temporary(n_Temporary     As Boolean)          : Let p_Temporary   = n_Temporary   : End Property

    Public Property Get Destination()             As VBIDE.VBProject   : Set Destination   = p_Destination : End Property
    Public Property Get Handler()                 As std_ErrorHandler  : Set Handler       = p_Handler     : End Property
    Public Property Get Recursive()               As Boolean           : Let Recursive     = p_Recursive   : End Property
    Public Property Get Log()                     As Boolean           : Let Log           = p_Log         : End Property
    Public Property Get Overwrite()               As Boolean           : Let Overwrite     = p_Overwrite   : End Property
    Public Property Get Increment()               As Boolean           : Let Increment     = p_Increment   : End Property
    Public Property Get Temporary()               As Boolean           : Let Temporary     = p_Temporary   : End Property

    Public Function Create(ByVal n_Project As VBIDE.VBProject, ByVal n_Handler As std_ErrorHandler) As std_Include
        Set Create = New std_Include
        With Create
            .Destination = n_Project
            .Handler     = n_Handler
        End With
    End Function

    Public Sub AddFormat(ByVal Format As IIncludeFormat)
        Call Push(Formats, Format)
    End Sub

    Public Function Build() As Boolean
        Dim i As Long
        For i = 0 To USize(ToBuildFormats)
            Build = BuildAll(ToBuildFormats(i), ToBuildSources(i))
            If Build = Handler.IS_ERROR Then Exit Function
        Next i
    End Function

    Public Function AddToQueue(ByVal Formater As IIncludeFormat, ByVal Source As std_IncludeSource) As Boolean
        Call Push(ToBuildFormats, Formater)
        Call Push(ToBuildSources, Source)
        AddToQueue = True
    End Function
'


'Handle Files
    'Ignore As String
    Public Function IncludeFolder(ByVal FolderPath As String, ParamArray Ignore() As Variant) As Boolean
        On Error GoTo Error
        Dim Size As Long
        Size = Ubound(Ignore)
        Dim Ignores() As String
        If Size <> -1 Then
            ReDim Ignores(Size)
            Dim i As Long
            For i = 0 To Size
                Ignores(i) = CStr(Ignore(i))
            Next i
        End If
        IncludeFolder = IncludeFolderArr(FolderPath, Ignores)
        Exit Function

        Error:
        ' Ignore is not string
    End Function

    Public Function IncludeFolderArr(ByVal FolderPath As String, ByRef Ignore() As String) As Boolean
        If Handler.Handle(FolderPath = Empty, std_Error.Create("std_Include", "severe", "IncludeFolderArr", "Variable [FolderPath] is Empty", Empty)) = Handler.IS_ERROR Then Exit Function
        If USize(Ignore) < 0 Then ReDim Ignore(0)

        Dim fso    As Object : Set fso    = CreateObject("Scripting.FileSystemObject")
        Dim Folder As Object : Set Folder = fso.GetFolder(FolderPath)
        IncludeFolderArr = RecursiveInclude(Folder, Ignore)
    End Function

    Private Function RecursiveInclude(ByVal Folder As Object, ByRef Ignore() As String) As Boolean
        Dim File As Object
        Dim Skip As Boolean
        Dim i As Long

        For Each File In Folder.Files
            For i = 0 To USize(Ignore)
                If File.Name = Cstr(Ignore(i)) Then
                    Skip = True
                    Exit For
                End If
            Next i
            If Skip Then
                Skip = False
            Else
                RecursiveInclude = IncludeFile(File.Path)
            End If
        Next
        Skip = False
        If Recursive Then
            Dim SubFolder As Object
            For Each SubFolder In Folder.SubFolders
                For i = 0 To USize(Ignore)
                    If SubFolder.Name = Cstr(Ignore(i)) Then
                        Skip = True
                        Exit For
                    End If
                Next i
                If Skip Then
                    Skip = False
                Else
                    RecursiveInclude = RecursiveInclude(SubFolder, Ignore)
                End If
            Next
        End If

    End Function

    Public Function IncludeFile(ByVal FilePath As String) As Boolean
        Dim Source As std_IncludeSource
        Set Source = std_IncludeSource.Create(FilePath)
        IncludeFile = AddQueue(Source)
    End Function
'


' Handle Declusion and replacement
    Public Function Declude(ByVal Component As VBIDE.VBComponent) As Boolean
        On Error GoTo Error
        Dim Message As String
        Message = Component.Name
        Call Destination.VBComponents.Remove(Component)
        If Log Then Call SingleLog("Declusion successful: ", Message)
        Exit Function

        Error:
        Call Handler.Raise(std_Error.Create("std_Include", "severe", "Declude", "Could not declude component, if there is no name then Component was Nothing: Name: VAR0", Empty, Message))
        Declude = Handler.IS_ERROR
    End Function

    Public Function OverWriteComponent(ByVal Component As VBIDE.VBComponent, ByVal NewCode As String) As VBIDE.VBComponent
        Dim CodeMod As VBIDE.CodeModule
        Set CodeMod = Component.CodeModule

        If CodeMod.CountOfLines > 0 Then
            Call CodeMod.DeleteLines(1, CodeMod.CountOfLines)
        End If
        Call CodeMod.AddFromString(NewCode)
        Set OverWriteComponent = Component
    End Function
'

    
' Handle Reference
    Public Function AddReference(Optional ByVal AddFromFile As String = Empty, Optional ByVal AddFromGuid As String = Empty, Optional ByVal Major As Long = 0, Optional ByVal Minor As Long = 0) As Boolean
        Dim Ref As VBIDE.Reference

        On Error GoTo Error
        If AddFromFile = Empty Then
            If AddFromGuid = Empty Then
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
        Call Handler.Raise("std_Include", "severe", "RemoveReference", "Could not add according to variables: AddFromFile: VAR0, AddFromGuid: VAR1, Major: VAR2, Minor: Var3", Empty, AddFromFile, AddFromGuid, Major, Minor)
    End Function

    Public Function RemoveReference(ByVal Value As String, ByVal AccordingTo As VBProjReferenceType) As Boolean
        On Error GoTo Error
        
        Dim Ref As VBIDE.Reference
        For Each Ref in VBProj.References
            Select Case AccordingTo
                Case VBProjReferenceType.FullPath    : If Ref.FullPath    = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.Name        : If Ref.Name        = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.Major       : If Ref.Major       = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.Minor       : If Ref.Minor       = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.Description : If Ref.Description = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.Typee       : If Ref.Type        = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.BuiltIn     : If Ref.BuiltIn     = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.GUID        : If Ref.GUID        = Value Then Call VBProj.References.Remove(Ref)
                Case VBProjReferenceType.IsBroken    : If Ref.IsBroken    = Value Then Call VBProj.References.Remove(Ref)
            End Select
        Next Ref
        Exit Function

        Error:
        Call Handler.Raise("std_Include", "severe", "RemoveReference", "Could not remove reference VAR0 according to VAR1", Empty, Value, AccordingTo)
    End Function
'


' Handle Other Projects
    ' Ignore As String
    Public Function IncludeProject(ByVal Proj As VBProject, ParamArray Ignore() As Variant) As Boolean
        On Error GoTo Error
        Dim Size As Long
        Size = USize(Ignore)
        Dim Ignores() As String
        If Size <> -1 Then
            ReDim Ignores(Size)
            Dim i As Long
            For i = 0 To Size
                Ignores(i) = CStr(Ignore(i))
            Next i
        End If
        IncludeProject = IncludeProjectArr(Proj, Ignores)
        Exit Function

        Error:
        ' Ignore is not string
    End Function

    Public Function IncludeProjectArr(ByVal Proj As VBProject, ByRef Ignore() As Variant) As Boolean
        Dim Component As VBComponent
        Dim Existing As Boolean
    
        Set Proj = WB.VBProject
        For Each Component In Proj.VBComponents
            If InArray(Component.Name, Ignore) = -1 Then
                Dim Source As std_IncludeSource
                Set Source = std_IncludeSource.Create(Component)
                Call AddQueue(Source)
            End If
        Next Component
    End Function
'


' Private Procedures

    Private Function AddQueue(ByVal Source As std_IncludeSource) As Boolean
        Dim i As Long
        Dim Result As Boolean

        If InQueue(Source) = True Then Exit Function
        For i = 0 To USize(Formats)
            Dim Formater As IIncludeFormat
            Set Formater = Formats(i)
            If Formater.CanHandle(Source) Then
                Result = Formater.AddQueue(Me, Source)
                Exit For
            End If
        Next i
        AddQueue = Result
    End Function

    Private Function BuildAll(ByVal Formater As IIncludeFormat, ByVal Source As std_IncludeSource) As Boolean
        Dim AddToTemporary As Boolean
        AddToTemporary = Temporary

        Dim Result As VBIDE.VBComponent
        Set Result = HandleExistingComponent(Formater, Source)
        If Result Is Nothing Then
            Set Result = Formater.Include(Me, Source)
            BuildAll = Handler.Handle(Result Is Nothing, std_Error.Create("std_Include", "severe", "AddQueue", "Could not include Source: VAR0", Empty, Source.Value))
            If BuildAll = Handler.IS_ERROR Then Exit Function
        Else
            AddToTemporary = False
        End If

        Dim Message As IErrorInterpret
        Set Message = std_ErrorInterpret.Create(Source.Value)
        If Log Then Call SingleLog("Inclusion successful: ", Message.Interpret)

        If AddToTemporary Then Call Push(TempIncludes, Result)
    End Function

    Private Function HandleExistingComponent(ByVal Formater As IIncludeFormat, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
        Dim Result As VBIDE.VBComponent

        Dim Name As String
        Name = Formater.Name(Source)
        If Exists(Name, Destination) Then
            Dim Found As VBIDE.VBComponent
            Set Found = Destination.VBComponents(Name)

            If OverWrite Then
                Dim Code As String
                Code = Formater.Code(Source)
                Set Result = OverWriteComponent(Found, Code)
            Else
                If Increment = False Then
                    Set Result = Found
                End If
            End If
        End If
        Set HandleExistingComponent = Result
    End Function

    Private Function Exists(ByVal Search As String, ByVal SearchIn As VBIDE.VBProject) As Boolean
        Dim Component As Object

        For Each Component In SearchIn.VBComponents
            If Component.Name = Search Then
                Exists = True
                Exit For
            End If
        Next
    End Function

    Private Function InArray(ByVal Search As String, ByRef Arr() As String) As Long
        Dim i As Long
        InArray = -1
        For i = 0 To USize(Arr)
            If Search Like Arr(i) Then
                InArray = Arr(i)
                Exit Function
            End If
        Next i
    End Function

    Private Function Push(ByRef Arr As Variant, ByVal Value As Variant) As Long
        Dim Size As Long
        Size = USize(Arr) + 1
        ReDim Preserve Arr(Size)
        Set Arr(Size) = Value
        Push = Size
    End Function

    Private Sub SingleLog(ByVal Message As String, ByVal Text As String)
        Dim Temp As Variant
        If Log Then
            Call Handler.Logger.Add(Message & Text, Temp)
        End If
    End Sub

    Private Function InQueue(ByVal Source As std_IncludeSource) As Boolean
        Dim i As Long
        For i = 0 To USize(ToBuildSources)
            If Compare(Source.Value, ToBuildSources(i).Value) Then
                InQueue = True
                Exit Function
            End If
        Next i
    End Function

    Private Function Compare(ByVal Arg1 As Variant, ByVal Arg2 As Variant) As Boolean
        If IsObject(Arg1) And IsObject(Arg2) Then
            Compare = Arg1 Is Arg2
        Else
            If Arg1 = Arg2 Then Compare = True
        End If
    End Function

    Private Function USize(ByRef Arr As Variant) As Long
        On Error Resume Next
        Dim Size As Long
        Size = -1
        Size = Ubound(Arr)
        USize = Size
    End Function
'

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Dim Temp As Boolean
    Dim i As Long
    For i = 0 To USize(TempIncludes)
        Temp = Declude(TempIncludes(i))
    Next
End Sub