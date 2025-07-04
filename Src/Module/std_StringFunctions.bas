Attribute VB_Name = "std_StringFunctions"

Option Explicit

Public Function MidP(Text As String, StartPoint As Long, EndPoint As Long) As String
    MidP = Mid(Text, StartPoint, (EndPoint - StartPoint) + 1)
End Function

Public Function InvertString(Text As String) As String
    Dim i As Long
    For i = Len(Text) To 1 Step -1
        InvertString = InvertString & Mid(Text, i, 1)
    Next i
End Function

Public Function RemoveNonNumericChars(Text As String) As String
    Dim i As Long
    For i = 1 To Len(Text)
        If IsNumeric(Mid(Text, i, 1)) Then RemoveNonNumericChars = RemoveNonNumericChars & Mid(Text, i, 1)
    Next i
End Function

Public Function IncrementString(Text As String) As String
    Dim LastChar As String
    LastChar = Mid(Text, Len(Text))
    If IsNumeric(LastChar) Then
        IncrementString = Mid(Text, Len(Text) - 1) & (CLng(LastChar) + 1)
    Else
        IncrementString = Text & 0
    End If
End Function

Public Function InStrAll(Text As String, Char As String) As Long()
    Dim Index As Long
    Dim Found As Long
    Dim i As Long
    Dim Temp() As Long

    Index = 1
    Found = InStr(Index, Text, Char)
    Do While Found <> 0
        Found = InStr(Index, Text, Char)
        If Found <> 0 Then
            ReDim Preserve Temp(i)
            Temp(i) = Found
            i = i + 1
            Index = Found + 1
            If Index > Len(Text) Then Exit Do
        End If
    Loop
    InStrAll = Temp
End Function

Public Function InString(Text As String, StartPoint As Long, EndPoint As Long) As Boolean
    Dim Quotes() As Long
    Dim i As Long
    Dim EndIndex As Long
    Quotes = InStrAll(Text, Chr(34))
    Select Case ArraySize(Quotes)
        Case 0
            If Quotes(0) = 0 Then
                InString = False
            ElseIf Quotes(0) =< StartPoint Then
                InString = True
            End If
        Case >0
            If (ArraySize(Quotes) + 1) Mod 2 = 0 Then
                EndIndex = ArraySize(Quotes)
            Else
                EndIndex = ArraySize(Quotes) - 1
                If Quotes(ArraySize(Quotes)) =< StartPoint Then InString = True: Exit Function
            End If
            For i = 0 To EndIndex Step+2
                If Quotes(i) =< StartPoint And Quotes(i + 1) >= EndPoint Then InString = True
            Next i
        Case Else
    End Select
End Function

Public Function GetParanthesesText(Line As String) As String
    Dim OpenPos() As Long: OpenPos = InStrAll(Line, "(") 
    Dim ClosePos() As Long: ClosePos = InStrAll(Line, ")")
    Dim StartPoint As Long
    Dim EndPoint As Long
    If ArraySize(OpenPos) = ArraySize(ClosePos) Then
        StartPoint = OpenPos(0) + 1
        EndPoint = ClosePos(ArraySize(ClosePos)) - 1
    End If
    If ArraySize(OpenPos) = 0 And OpenPos(0) = 0 Then StartPoint = Len(Line) + 1
    If ArraySize(ClosePos) = 0 And ClosePos(0) = 0 Then EndPoint = Len(Line)
    GetParanthesesText = MidP(Line, StartPoint, EndPoint)
End Function

Public Function GetProcedureName(Line As String) As String
    Dim i As Long
    i = InStr(1, Line, "(")
    If i = 0 Then
        GetProcedureName = Line
    Else
        GetProcedureName = MidP(Line, 1, i -1)
    End If
End Function

Private Function ArraySize(Arr As Variant, Optional Dimension As Long = 1)
    On Error Resume Next
    ArraySize = -1
    ArraySize = Ubound(Arr, Dimension)
End Function