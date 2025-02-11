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
    Found = InStr(Text, Index, Len(Text))
    Do While Found <> 0
        Found = InStr(Text, Index, Len(Text))
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