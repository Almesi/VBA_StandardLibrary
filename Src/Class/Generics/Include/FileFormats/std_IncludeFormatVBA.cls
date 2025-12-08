VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIncludeFormatVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Implements IIncludeFormat


Private Reader As New std_IncludeFileReader

Public Function Create() As IIncludeFormatVBA
    Set Create = New IIncludeFormatVBA
End Function


Private Function IIncludeFormat_Include(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
    Set IIncludeFormat_Include = Includer.Destination.VBComponents.Import(Source.Value)
End Function

Private Function IIncludeFormat_AddQueue(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As Boolean
    IIncludeFormat_AddQueue = Includer.AddToQueue(Me, Source)
End Function

Private Function IIncludeFormat_CanHandle(ByVal Source As std_IncludeSource) As Boolean
    If Source.Value       = Empty    Then Exit Function
    If Source.SourceType <> "String" Then Exit Function

    Dim Format As String
    Format = Reader.FileFormat(Source.Value)
    IIncludeFormat_CanHandle = Format Like "bas" Or _
                               Format Like "cls" Or _
                               Format Like "frm"
End Function

Private Function IIncludeFormat_Code(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Code = Reader.ReadFileCompilable(Source.Value)
End Function

Private Function IIncludeFormat_Name(ByVal Source As std_IncludeSource) As String
    IIncludeFormat_Name = Reader.GetFileName(Source.Value)
End Function