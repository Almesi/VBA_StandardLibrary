
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mName          As String
Private mReturn        As QueryReturnType
Private mArguments     As Collection
Private mBasicCommand  As String
Private mArgumentCount As Long


'===============================================================================
' CREATION
'===============================================================================

Public Function Create(ByVal nName As String, ByVal nReturn As QueryReturnType, ByVal nArgumentCount As Long) As QueryDefinition
    Set Create = New QueryDefinition
    Call Create.Initialize(nName, nReturn, nArgumentCount)
End Function

Friend Sub Initialize(ByVal nName As String, ByVal nReturn As QueryReturnType, ByVal nArgumentCount As Long)
    Let mName          = nName
    Set mReturn        = nReturn
    Set mArguments     = New Collection
    Let mBasicCommand  = vbNullString
    Let mArgumentCount = nArgumentCount

End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Friend Sub SetArgumentCount(ByVal nArgumentCount As Long)
    mArgumentCount = nArgumentCount
End Sub

Friend Sub SetReturnType(ByVal nReturn As QueryReturnType)
    Set mReturn = nReturn
End Sub

Friend Sub SetBasicCommand(ByVal nBasicCommand As String)
    mBasicCommand = nBasicCommand
End Sub


Public Property Get Name() As String
    Name = mName
End Property

Public Property Get ReturnType() As QueryReturnType
    Set ReturnType = mReturn
End Property

Public Property Get BasicCommand() As String
    BasicCommand = mBasicCommand
End Property

Public Property Get ArgumentCount() As Long
    ArgumentCount = mArgumentCount
End Property

Public Property Get ArgumentLinkCount() As Long
    Dim Argument As QueryArgument
    Dim Result As Long

    For Each Argument In mArguments
        If Not Argument.Link Is Nothing Then
            Result = Result + 1
        End If
    Next Argument

    ArgumentLinkCount = Result
End Property


'===============================================================================
' ARGUMENTS
'===============================================================================

Public Sub AddArgument(ByVal nArgument As QueryArgument)
    Call mArguments.Add(nArgument)
End Sub

Public Property Get Arguments() As Collection
    Set Arguments = mArguments
End Property


Public Property Get Argument(ByVal nIndex As Long) As QueryArgument
    If nIndex < 1 Or nIndex > mArguments.Count Then
        Err.Raise vbObjectError + 1003, "QueryDefinition", "Invalid argument index: " & CStr(nIndex)
    End If

    Set Argument = mArguments(nIndex)
End Property


'===============================================================================
' COPYING
'===============================================================================

Public Function Copy() As QueryDefinition
    Dim Result As QueryDefinition
    Dim Original As QueryArgument
    Dim ArgumentCopy As QueryArgument
    Dim i As Long

    Set Result = QueryDefinition.Create(mName, mReturn, mArgumentCount)

    Call Result.SetBasicCommand(mBasicCommand)

    For i = 1 To mArguments.Count

        Set Original = mArguments(i)
        Set ArgumentCopy = CopyArgument(Original)

        Call Result.AddArgument(ArgumentCopy)

    Next i

    Set Copy = Result
End Function


Private Function CopyArgument(ByVal nArgument As QueryArgument) As QueryArgument
    Dim Result As QueryArgument
    Dim Source As QueryDefinition

    Set Result = QueryArgument.Create(nArgument.ArgumentReturn)

    Result.Value = nArgument.Value

    If Not nArgument.Source Is Nothing Then
        Set Source = nArgument.Source.Copy()
        Set Result.Source = Source
    End If

    If Not nArgument.Link Is Nothing Then
        Set Result.Link = QueryArgumentLink.Create(nArgument.Link.InputIndex, nArgument.Link.OutputIndex)
    End If

    Set CopyArgument = Result
End Function