VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IIncludeFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Public Function Include(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As VBIDE.VBComponent
End Function

Public Function AddQueue(ByVal Includer As std_Include, ByVal Source As std_IncludeSource) As Boolean
End Function

Public Function CanHandle(ByVal Source As std_IncludeSource) As Boolean
End Function

Public Function Code(ByVal Source As std_IncludeSource) As String
End Function

Public Function Name(ByVal Source As std_IncludeSource) As String
End Function