VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIncludeFormatDependencies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IIncludeFormat

Private Reader As New std_IncludeFileReader

Public Function Create() As IIncludeFormatDependencies
    Set Create = New IIncludeFormatDependencies
End Function


Private Function IIncludeFormat_Include(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
    Set IIncludeFormat_Include = Nothing
End Function

Private Function IIncludeFormat_AddQueue(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As Boolean
    Dim Code    As String : Code = Reader.ReadFile(Source.Value)
    Dim Lines() As String : Lines = Split(Code, vbCrLf)

    Dim MeObj As IIncludeFormat
    Set MeObj = Me

    Dim i As Long
    For i = 0 To Ubound(Lines)
        Dim BasePath As String : BasePath = FileFolderPath(Source.Value)
        Dim NextPath As String : NextPath = ResolveRelativePath(BasePath, Lines(i))
        If IsFile(NextPath) Then
            Call Includer.IncludeFile(NextPath)
        Else
            Call Includer.IncludeFolder(NextPath)
        End If
    Next i
    IIncludeFormat_AddQueue = True
End Function

Private Function IIncludeFormat_CanHandle(ByVal Source As std_IncludeSource) As Boolean
    If Source.Value       = Empty    Then Exit Function
    If Source.SourceType <> "String" Then Exit Function

    IIncludeFormat_CanHandle = UCase(Reader.GetFileName(Source.Value)) = "DEPENDENCIES"
End Function

Private Function IIncludeFormat_Code(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Code = Reader.ReadFileCompilable(Source.Value)
End Function

Private Function IIncludeFormat_Name(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Name = Reader.GetFileName(Source.Value)
End Function


Private Function ResolveRelativePath(ByVal BasePath As String, ByVal RelativePath As String) As String
    Dim BaseParts() As String
    Dim RelParts() As String
    Dim Stack As Collection
    Dim Part As Variant
    Dim Result As String
    Dim i As Long
    
    BasePath     = Replace(BasePath, "/", "\")
    RelativePath = Replace(RelativePath, "/", "\")
    
    ' Split paths
    BaseParts = Split(BasePath, "\")
    RelParts = Split(RelativePath, "\")
    
    Set Stack = New Collection
    
    ' Push base path parts into Stack
    For i = 0 To UBound(BaseParts)
        Call Stack.Add(BaseParts(i))
    Next i
    
    ' Process relative path directives
    For i = 0 To UBound(RelParts)
        Part = RelParts(i)
        
        Select Case Part
            Case Empty, "."
            Case ".."
                If Stack.Count > 0 Then Call Stack.Remove(Stack.Count)
            Case Else
                Call Stack.Add(Part)
        End Select
    Next i
    
    ' Rebuild the path
    Result = Empty
    For i = 1 To Stack.Count
        Result = Result & Stack(i) & "\"
    Next i
    
    ' Trim trailing backslash
    If Right(Result, 1) = "\" Then
        Result = Mid(Result, 1, Len(Result) - 1)
    End If
    
    ResolveRelativePath = Result
End Function

Private Function FileFolderPath(ByVal FilePath As String) As String
    Dim Parts() As String : Parts = Split(FilePath, "\")
    Dim Size    As Long   : Size  = Ubound(Parts)
    If IsFile(FilePath) Then
        Dim i As Long
        Dim Result As String
        For i = 0 To Size - 1
            Result = Result & Parts(i) & "\"
        Next i
        FileFolderPath = Mid(Result, 1, Len(Result) - 1)
    Else
        FileFolderPath = FilePath
    End If
End Function

Private Function IsFile(ByVal Path As String) As Boolean
    On Error GoTo Error
    If Len(Dir(Path, vbNormal)) > 0 Then
        IsFile = True
        Exit Function
    End If
    
    If Len(Dir(Path, vbDirectory)) > 0 Then
        IsFile = False
        Exit Function
    End If
    
    Error:
    IsFile = False
End Function