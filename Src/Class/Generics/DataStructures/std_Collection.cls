VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ==================================================
' ====================[Properties]==================
' ==================================================

Private mCollection As New Collection
Private mClassName  As String


' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function CreateFromValue(ByVal Value As Variant, Optional ByVal Key As Variant, Optional ByVal Before As Variant, Optional ByVal After As Variant) As std_Collection
    Set CreateFromValue = CreateFromVariant(Value)
    Call CreateFromValue.Add(Value, Key, Before, After)
End Function

Public Function CreateFromVariant(ByVal Value As Variant) As std_Collection
    Set CreateFromVariant = Create(TypeName(Value))
End Function

Public Function Create(ByVal ClassName As String) As std_Collection
    Set Create = New std_Collection
    Call Create.Initialize(ClassName)
End Function

Public Function Clone() As std_Collection
    Set Clone = Create(mClassName)
    Call Clone.AddFromstd_Collection(Me)
End Function

Friend Sub Initialize(ByVal ClassName As String)
    If ClassName = Empty Then Err.Raise vbObjectError + 1001, "std_Collection.Initialize", "ClassName cannot be empty."
    mClassName = ClassName
End Sub


' ==================================================
' ====================[Getters]=====================
' ==================================================

Public Property Get Count() As Long
    Count = mCollection.Count
End Property

Public Property Get ClassName() As String
    ClassName = mClassName
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mCollection.Count = 0)
End Property

Public Property Get First() As Variant
    If IsEmpty Then Err.Raise vbObjectError + 1010, "std_Collection.First", "Collection is empty."
    Call Assign(First, mCollection.Item(1))
End Property

Public Property Get Last() As Variant
    If IsEmpty Then Err.Raise vbObjectError + 1011, "std_Collection.Last", "Collection is empty."
    Call Assign(Last, mCollection.Item(mCollection.Count))
End Property

Public Property Get Item(ByVal IndexOrKey As Variant) As Variant
    Call Assign(Item, mCollection.Item(IndexOrKey))
End Property

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCollection.[_NewEnum]
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Sub Add(ByVal Value As Variant, Optional ByVal Key As Variant, Optional ByVal Before As Variant, Optional ByVal After As Variant)
    Call ValidateItem(Value)

    If Not IsMissing(Before) And Not IsMissing(After) Then
        Err.Raise vbObjectError + 1005, "std_Collection.Add", "Before and After cannot both be specified."
    End If

    If IsMissing(Key) Then
        If IsMissing(Before) Then
            If IsMissing(After) Then
                Call mCollection.Add(Value)
            Else
                Call mCollection.Add(Value, , , After)
            End If
        Else
            Call mCollection.Add(Value, , Before)
        End If
    Else
        If IsMissing(Before) Then
            If IsMissing(After) Then
                Call mCollection.Add(Value, CStr(Key))
            Else
                Call mCollection.Add(Value, CStr(Key), , After)
            End If
        Else
            Call mCollection.Add(Value, CStr(Key), Before)
        End If
    End If
End Sub

Public Sub AddFromstd_Collection(ByVal NewCollection As std_Collection)
    Dim i As Long
    Dim ItemCount As Long

    If NewCollection Is Nothing Then Exit Sub
    If StrComp(NewCollection.ClassName, mClassName, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 1012, "std_Collection.AddFromstd_Collection", "Collection type mismatch."

    ItemCount = NewCollection.Count

    For i = 1 To ItemCount
        Call Me.Add(NewCollection.Item(i))
    Next i
End Sub

Public Sub Insert(ByVal Value As Variant, ByVal Index As Long)
    If Index < 1 Or Index > Count + 1 Then
        Err.Raise vbObjectError + 1013, "std_Collection.Insert", "Index is out of range."
    End If

    If Index = Count + 1 Then
        Call Me.Add(Value)
    Else
        Call Me.Add(Value, , Index)
    End If
End Sub

Public Sub Remove(ByVal IndexOrKey As Variant)
    Call mCollection.Remove(IndexOrKey)
End Sub

Public Sub Clear()
    Set mCollection = New Collection
End Sub

Public Function Contains(ByVal Value As Variant) As Boolean
    Dim Result As Long
    Result = IndexOf(Value)
    Contains = (Result <> 0)
End Function

Public Function Exists(ByVal IndexOrKey As Variant) As Boolean
    Dim Result As Variant

    On Error Resume Next
    Err.Clear
    Call Assign(Result, mCollection.Item(IndexOrKey))
    Exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Public Function IndexOf(ByVal Value As Variant) As Long
    Dim i As Long

    Call ValidateItem(Value)

    For i = 1 To Count
        If EqualsAt(i, Value) Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function ToArray() As Variant
    Dim Result() As Variant
    Dim i As Long

    If IsEmpty Then
        ToArray = Array()
        Exit Function
    End If

    ReDim Result(0 To Count - 1)

    For i = 1 To Count
        Call Assign(Result(i - 1), mCollection.Item(i))
    Next i

    ToArray = Result
End Function

Public Sub Reverse()
    Dim Result As New Collection
    Dim i As Long

    For i = Count To 1 Step -1
        Call Result.Add(mCollection.Item(i))
    Next i

    Set mCollection = Result
End Sub


' ==================================================
' ====================[Helpers]=====================
' ==================================================

Private Sub ValidateItem(ByVal Value As Variant)
    If StrComp(TypeName(Value), mClassName, vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 1009, "std_Collection.ValidateItem", "Invalid object type. Expected '" & mClassName & "', received '" & TypeName(Value) & "'."
End Sub

Private Function EqualsAt(ByVal Index As Long, ByVal Value As Variant) As Boolean
    Dim Current As Variant

    Call Assign(Current, mCollection.Item(Index))

    If IsObject(Current) Then
        If Current Is Value Then EqualsAt = True
    Else
        If Current = Value Then EqualsAt = True
    End If
End Function