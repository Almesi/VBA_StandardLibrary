VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConsoleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
 
Public Function Run(Model As Object) As Boolean
End Function

Public Sub Load(Optional n_ColorDef As ColorDefinition = Nothing, Optional ScreenHeight As Long = 5000, Optional ScreenWidth As Long = 3000)
End Sub

Public Function CheckPredeclaredAnswer(ReturnVariable As Variant, Message As Variant, AllowedValues As ArrayValueModel) As Long
End Function

Public Sub PrintEnter(Text As String, Colors() As Long, ColorLength() As Long)
End Sub

Public Sub PrintConsole(Text As String, Colors() As Long, ColorLength() As Long)
End Sub

Public Sub ColorConsole(StartPoint As Long, Colors() As Long, ColorLength() As Long)
End Sub