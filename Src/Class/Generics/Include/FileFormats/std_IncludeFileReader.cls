VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_IncludeFileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

' Helper Class for std_Include and consort

Public Function FileFormat(ByVal FilePath As String) As String
    Dim LastPoint As Long
    LastPoint = InStrRev(FilePath, ".")
    FileFormat = Mid(FilePath, LastPoint + 1, Len(FilePath))
End Function

Public Function GetFileName(ByVal FilePath As String)
    Dim FileText As String : FileText = ReadFile(FilePath)
    Dim Lines()  As String : Lines    = Split(FileText, vbCrLf)

    Dim i As Long
    For i = 0 To Ubound(Lines)
        Dim Line As String
        Line = Lines(i)
        If UCase(Line) Like "ATTRIBUTE*VB_NAME*=*" Then
            Dim StartPoint As Long   : StartPoint = InStr(1, Line, Chr(34))          + 1 '= "
            Dim EndPoint   As Long   : EndPoint   = InStr(StartPoint, Line, Chr(34)) - 1 '= "
            Dim Found      As String : Found      = MidP(Line, StartPoint, EndPoint)
            Exit For
        End If
    Next i

    If Found = Empty Then
        Found = FileName(FilePath)
    End If
    GetFileName = Found
End Function

Public Function ReadFile(ByVal FilePath As String) As String
    On Error GoTo Error
    Dim FileNum As Integer
    Dim FileContent As String
    
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    FileContent = Input$(LOF(FileNum), FileNum)
    Close #FileNum
    
    ReadFile = FileContent
    Exit Function

    Error:
    ReadFile = ReadBinaryFile(FilePath)
End Function

Public Function ReadBinaryFile(ByVal FilePath As String) As String
    On Error GoTo Error
    Dim Bytes() As Byte
    Dim FileNum As Integer
    FileNum = FreeFile
    
    Open FilePath For Binary Access Read As #FileNum
        ReDim Bytes(0 To LOF(FileNum) - 1)
        Get #FileNum, , Bytes
    Close #FileNum
    
    Dim Result As String
    Dim i As Long
    For i = 0 To Ubound(Bytes)
        Result = Result & Chr(Bytes(i))
    Next i
    ReadBinaryFile = Result

    Error:
End Function


Public Function FileName(ByVal FilePath As String) As String
    Dim LastPoint As Long : LastPoint = InStrRev(FilePath, "\")
    If LastPoint = 0 Then   LastPoint = InStrRev(FilePath, "/")
    If LastPoint = 0 Then   LastPoint = 1
    Dim EndPoint  As Long : EndPoint  = InStr(LastPoint, FilePath, ".")
    If EndPoint  = 0 Then   EndPoint  = Len(FilePath) + 1
    FileName = MidP(FilePath, LastPoint + 1, EndPoint - 1)
End Function

Public Function ReadFileCompilable(ByVal FilePath As String) As String
    Dim Text         As String : Text = ReadFile(FilePath)
    Dim Lines()      As String : Lines = Split(Text, vbCrLf)
    Dim ValidLines() As String

    Dim i As Long
    For i = 0 To Ubound(Lines)
        If ValidLine(UCase(Lines(i))) Then
            Call Push(ValidLines, Lines(i))
        End If
    Next i
    ReadFileCompilable = Join(ValidLines, vbCrLf)
End Function

Public Function GetComponentType(ByVal FilePath As String) As vbext_ComponentType
    Dim Code    As String              : Code   = ReadFile(FilePath)
    Dim Lines() As String              : Lines  = Split(Code, vbCrLf)
    Dim Result  As vbext_ComponentType : Result = vbext_ct_StdModule

    If UCase(GetFileName(FilePath)) Like "SHEET*"       Then Result = vbext_ct_Document
    If UCase(GetFileName(FilePath)) Like "THISWORKBOOK" Then Result = vbext_ct_Document
    
    Dim i As Long
    For i = 0 To UBound(Lines)
        Dim Line As String
        Line = Lines(i)
        
        If Line <> "" And Not Line Like "'*" Then
            If UCase(Line) Like "BEGIN {*" Then
                Result = vbext_ct_MSForm
                Exit For
            End If
            
            If UCase(Line) Like "VERSION*CLASS" Then
                Result = vbext_ct_ClassModule
                Exit For
            End If
        End If
    Next i
    GetComponentType = Result
End Function



Private Function MidP(ByVal Text As String, ByVal StartPoint As Long, ByVal EndPoint As Long) As String
    MidP = Mid(Text, StartPoint, (EndPoint - StartPoint) + 1)
End Function

Private Function Push(ByRef Arr() As String, ByVal Value As String) As Long
    On Error Resume Next
    Dim Size As Long
    Size = -1
    Size = Ubound(Arr)
    Size = Size + 1
    ReDim Preserve Arr(Size)
    Arr(Size) = Value
    Push = Size
End Function

Private Function ValidLine(ByVal Line As String) As Boolean
    ValidLine = Not (Line Like "VERSION*"        Or _
                       Line Like "*MULTIUSE = *" Or _
                       Line Like "BEGIN*"        Or _
                       Line Like "END"           Or _
                       Line Like "ATTRIBUTE*")
End Function