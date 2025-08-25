VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Destination"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private p_Destination As Variant

Public Event BeforeAdd(Text As String)
Public Event AfterAdd(Text As String)

Public Property Let Destination(n_Value As Variant)
    If IsObject(n_Value) Then
        Set p_Destination = n_Value
    Else
        p_Destination = n_Value
    End If
End Property

Public Property Get Destination() As Variant
    If IsObject(p_Destination) Then
        Set Destination = p_Destination
    Else
        Destination = p_Destination
    End If
End Property

Public Function Create(n_Destination As Variant) As std_Destination
    Set Create = New std_Destination
    Create.Destination = n_Destination
End Function

Public Function Add(Text As String) As Long
    Dim i As Long
    Dim fso As Object
    Dim TextStream As Object
    On Error Resume Next ' To avoid empty TextStream file
    RaiseEvent BeforeAdd(Text)
    Select Case True
        Case TypeName(p_Destination) = "Console"
            Add = LoggingConsole(Text)
        Case TypeName(p_Destination) = "Range"
            Do Until p_Destination.Offset(i, 0).Formula = ""
                i = i + 1
            Loop
            p_Destination.Offset(i, 0).Formula = Text
            Add = i
        Case TypeName(p_Destination) = "File", TypeName(p_Destination) = "String"
            If TypeName(p_Destination) = "String" Then
                Set TextStream = fso.GetFile(p_Destination).OpenAsTextStream(1, -2) ' For Reading, TriStateUseDefault
                Text = TextStream.ReadAll & Text
                Set TextStream = fso.GetFile(p_Destination).OpenAsTextStream(2, -2) ' For Writing, TriStateUseDefault
            Else
                Set TextStream = p_Destination.OpenAsTextStream(1, -2)
                Text = TextStream.ReadAll & Text
                Set TextStream = p_Destination.OpenAsTextStream(2, -2)
            End If
            TextStream.WriteLine Text
            Add = TextStream.Line
            TextStream.Close
        Case (p_Destination = Empty)
            MsgBox (Text)
            Add = -1
        Case (p_Destination Is Nothing)
            Debug.Print Text
            Add = -1
        Case Else
    End Select
    RaiseEvent AfterAdd(Text)
End Function


Private Function LoggingConsole(Text As String) As Long
    Call p_Destination.PrintEnter(Text, (&H100000FF - &H10000000))
    LoggingConsole =  p_Destination.LineCount()
End Function