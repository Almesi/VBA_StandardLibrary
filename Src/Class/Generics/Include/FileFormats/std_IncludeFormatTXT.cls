VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIncludeFormatTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IIncludeFormat

Private Reader As New std_IncludeFileReader

Public Function Create() As IIncludeFormatTXT
    Set Create = New IIncludeFormatTXT
End Function


Private Function IIncludeFormat_Include(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
    Dim Name   As String            : Let Name   = Reader.GetFileName(Source.Value)
    Dim Code   As String            : Let Code   = Reader.ReadFileCompilable(Source.Value)
    Dim Result As VBIDE.VBComponent : Set Result = Includer.Destination.VBComponents.Add(Reader.GetComponentType(Source.Value))

    Result.Name = Name
    Call Result.CodeModule.AddFromString(Code)
    Set IIncludeFormat_Include = Result
End Function

Private Function IIncludeFormat_AddQueue(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As Boolean
    IIncludeFormat_AddQueue = Includer.AddToQueue(Me, Source)
End Function

Private Function IIncludeFormat_CanHandle(ByVal Source As std_IncludeSource) As Boolean
    If Source.Value       = Empty    Then Exit Function
    If Source.SourceType <> "String" Then Exit Function

    Dim Format As String
    Format = Reader.FileFormat(Source.Value)
    IIncludeFormat_CanHandle = Format Like "txt"
End Function

Private Function IIncludeFormat_Code(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Code = Reader.ReadFileCompilable(Source.Value)
End Function

Private Function IIncludeFormat_Name(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Name = Reader.GetFileName(Source.Value)
End Function