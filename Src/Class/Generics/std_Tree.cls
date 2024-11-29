    VERSION 1.0 CLASS
    BEGIN
    MultiUse = -1  'True
    END
    Attribute VB_Name = "std_Tree"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = False
    Attribute VB_PredeclaredId = True
    Attribute VB_Exposed = False


Option Explicit

' Private Variables
    Private Type Node
        Value As Variant
        Branches() As Long
    End Type


    Private p_Tree() As Node
    Private p_Width As Long
    Private p_Depth As Long
'

' Events
    Public Event BeforeAdd(Index As Long, Value As Variant)
    Public Event AfterAdd(Index As Long, Value As Variant)
    Public Event BeforeDelete(Index As Long)
    Public Event AfterDelete(Index As Long)
'

' Let Set Get

    Public Property Let Value(Index As Long, Variable As Variant)
        If IsObject(Variable) Then
            Set p_Tree(Index).Value = Variable
        Else
            p_Tree(Index).Value = Variable
        End If
    End Property

    Public Property Get Value(Index As Long) As Variant
        If IsObject(p_Tree(Index).Value) Then
            Set Value = p_Tree(Index).Value
        Else
            Value = p_Tree(Index).Value
        End If
    End Function

    Public Property Let TreeData(ByVal n_Tree As std_Tree)
        Dim Temp() As Node
        Temp = p_Tree
        Me.Tree = n_Tree.Tree
        p_Width = n_Tree.Width
        p_Depth = n_Tree.Depth
    End Property

    Public Property Get Tree() As Variant()
        Dim BiggestSize As Long: BiggestSize = -1
        Dim i As Long, j As Long
        Dim Temp() As Variant
        ReDim Temp(UboundT(p_Tree), 0)
        For i = 0 To UboundT(p_Tree)
            Temp(i, 0) = p_Tree(i).Value
            If BiggestSize < UboundK(p_Tree(i).Branches) Then BiggestSize = UboundK(p_Tree(i).Branches)
        Next
        If BiggestSize <> -1 Then ReDim Preserve Temp(UboundT(p_Tree), BiggestSize + 1)
        For i = 0 To UboundK(Temp, 1)
            For j = 1 To UboundK(Temp, 2)
                If j =< UboundK(p_Tree(i).Branches) + 1 Then
                    Temp(i, j) = p_Tree(i).Branches(j - 1)
                End If
            Next j
        Next i
        Tree = Temp
    End Property

    Public Property Let Tree(Data() As Variant)
        Dim i As Long, j As Long, k As Long
        Dim Temp() As Node
        ReDim Temp(UboundK(Data, 1))
        For i = 0 To UboundK(Data, 1)
            Temp(i).Value = Data(i, 0)
            j = 1
            For j = 1 To UboundK(Data, 2)
                If Data(i, j) = Empty Then Exit For
            Next
            j = j - 2 ' One for 1 and one for Empty
            If j >= 0 Then ReDim Temp(i).Branches(j)
            For k = 0 To j
                Temp(i).Branches(k) = Data(i, k + 1)
            Next
        Next i
        p_Tree = Temp
    End Property

    Public Property Get Width() As Long
        Width = p_Width
    End Property

    Public Property Get Depth() As Long
        Depth = p_Depth
    End Property
'

' Public Functions
    Public Function Create(Optional Branches As Long = -1, Optional Depth As Long = 0) As std_Tree
        Call Init(Branches, Depth)
        If Branches < -1 Or Depth < 0 Then
            Set Create = Nothing
        Else
            Set Create = New std_Tree
            Call CreateTreeRecursion(0, Branches, Depth)
            Create.TreeData = Me
        End If
        Call Init(0, 0)
    End Function

    Public Function Add(Index As Long, Value As Variant) As Long
        RaiseEvent BeforeAdd(Index, Value)
        If Width <> -1 Then
            ReDim Preserve p_Tree(UBoundT(p_Tree) + 1)
            If IsObject(Value) Then
                Set p_Tree(UBoundT(p_Tree)).Value = Value
            Else
                p_Tree(UBoundT(p_Tree)).Value = Value
            End If

            ReDim Preserve p_Tree(Index).Branches(UBoundK(p_Tree(Index).Branches) + 1)
            p_Tree(Index).Branches(UBoundK(p_Tree(Index).Branches)) = UBoundT(p_Tree)
            Add = UBoundT(p_Tree)
        Else
            Add = -1
        End If
        RaiseEvent AfterAdd(Index, Value)
    End Function

    Public Function Delete(Index As Long) As Long
        Dim i As Long, j As Long
        Dim Temp() As Node
        Dim ToDelete() As Long
        Dim Skip As Boolean
        RaiseEvent BeforeDelete(Index)
        Temp = p_Tree
        ReDim ToDelete(0)
        ToDelete(0) = Index
        Call GetIndexAll(ToDelete, Index)
        ReDim p_Tree(UBoundT(p_Tree) - (UBoundK(ToDelete) + 1))
        For i = 0 To UBoundT(p_Tree)
            For j = 0 To UBoundK(ToDelete)
                If ToDelete(j) = i Then Skip = True: Exit For
            Next
            If Skip = False Then p_Tree(i) = Temp(i)
            Skip = False
        Next i

        For i = 0 To UboundT(p_Tree)
            For j = 0 To UboundK(ToDelete)
                p_Tree(i).Branches = ReRollBranches(p_Tree(i).Branches, ToDelete(j))
            Next j
        Next i
        RaiseEvent AfterDelete(Index)
    End Function

    Public Sub GetIndexAll(ByRef Variable() As Long, ByVal Index As Long)
        Dim i As Long
        Dim NewSize As Long
        Dim Temp As Long

        Temp = UBoundK(Variable)
        NewSize = UboundK(p_Tree(Index).Branches)
        If NewSize <> -1 Then
            ReDim Preserve Variable(Temp + NewSize + 1)
            For i = 0 To NewSize
                Variable(i + Temp  + 1) = p_Tree(Index).Branches(i)
                Call GetIndexAll(Variable, p_Tree(Index).Branches(i))
            Next i
        End If
    End Sub

    Public Function GetIndex(Path() As Long) As Long
        Dim i As Long
        GetIndex = Path(0)
        For i = 1 To UBoundK(Path)
            GetIndex = p_Tree(GetIndex).Branches(Path(i))
        Next
    End Function

    Public Function GetIndexParam(ParamArray Path() As Variant)
        Dim Temp() As Long
        Dim TempVar As Variant
        Dim i As Long
        TempVar = Path
        ReDim Temp(UBoundK(TempVar))
        For i = 0 To UBoundK(TempVar)
            Temp(i) = Clng(TempVar(i))
        Next i
        GetIndexParam = GetIndex(Temp)
    End Function

    Public Sub Sort(Optional Ascending As Boolean = True)
        Dim Arr() As Long
        Call SortIndex(Arr, 0)
        Call SortBranch(Arr, Ascending)
    End Sub

    Public Function Find(Search As Variant, Index As Long, SearchDepth As Long) As Long
        Dim i As Long
        Find = -1
        If SearchDepth = -1 Then
            Exit Function
        End If
        If IsObject(p_Tree(Index).Value) Then
            If p_Tree(Index).Value Is Search Then Find = Index: Exit Function
        Else
            If p_Tree(Index).Value = Search Then Find = Index: Exit Function
        End If
        For i = 0 To UBoundK(p_Tree(Index).Branches)
            Find = Find(Search, p_Tree(Index).Branches(i), SearchDepth - 1)
            If Find <> -1 Then Exit Function
        Next i
    End Function

    Public Function Search(Value As Variant, Index As Long, SearchDepth As Long) As Long
        Dim i As Long
        Search = -1
        If SearchDepth = -1 Then
            Exit Function
        End If
        If p_Tree(Index).Value = Value Then
            Search = Index
        ElseIf Value < p_Tree(Index).Value Then
            For i = 0 To MiddleBranch(Index)
                If Search <> -1 Then Exit Function
                Search = Search(Value, p_Tree(Index).Branches(i), SearchDepth - 1)
            Next
        Else
            For i = MiddleBranch(Index) + 1 To UboundK(p_Tree(Index).Branches)
                If Search <> -1 Then Exit Function
                Search = Search(Value, p_Tree(Index).Branches(i), SearchDepth - 1)
            Next
        End If
    End Function

    Public Sub Display()
        Dim Temp() As String
        Dim i As Long
        Call DisplayX(-1, 0, Temp)
        For i = 0 To UboundK(Temp)
            Debug.Print Temp(i)
        Next i
    End Sub

    Public Sub Clear()
        Dim Temp() As Node
        p_Tree = Temp
    End Sub
'

' Privates
    Private Sub Init(Width As Long, Depth As Long)
        ReDim p_Tree(0)
        p_Width = Width
        p_Depth = Depth
    End Sub

    Private Function MiddleBranch(Index As Long) As Long
        MiddleBranch = Int(UBoundK(p_Tree(Index).Branches) / 2)
    End Function

    Private Function SetTo4Hex(Value As Variant) As String
        If IsObject(Value) Then
            SetTo4Hex = " Obj"
        Else
            SetTo4Hex = CStr(Hex(Value))
        End If
        Select Case Len(SetTo4Hex)
            Case 1: SetTo4Hex = "000" & SetTo4Hex
            Case 2: SetTo4Hex = "00"  & SetTo4Hex
            Case 3: SetTo4Hex = "0"   & SetTo4Hex
        End Select
    End Function

    Private Sub SortBranch(Arr() As Long, Ascending As Boolean)
        Dim i As Long, j As Long
        Dim Temp As Variant

        If Ascending Then
            For i = 0 To UboundK(Arr) - 1
                For j = 0 To UBoundK(Arr) - i - 1
                    If p_Tree(Arr(j)).Value > p_Tree(Arr(j + 1)).Value Then
                        Temp = p_Tree(Arr(j)).Value
                        p_Tree(Arr(j)).Value = p_Tree(Arr(j + 1)).Value
                        p_Tree(Arr(j + 1)).Value = Temp
                    End If
                Next j
            Next i
        Else
        For i = UboundK(Arr) To 0 + 1 Step-1
                For j = UboundK(Arr) To 0 + i + 1 Step-1
                    If p_Tree(Arr(j)).Value > p_Tree(Arr(j - 1)).Value Then
                        Temp = p_Tree(Arr(j)).Value
                        p_Tree(Arr(j)).Value = p_Tree(Arr(j - 1)).Value
                        p_Tree(Arr(j - 1)).Value = Temp
                    End If
                Next j
            Next i
        End If
    End Sub
    
    Private Sub SortIndex(ByRef Arr() As Long, Index As Long)
        Dim i As Long
        Dim BranchSize As Long
        Dim Pushed As Boolean
        BranchSize = UboundK(p_Tree(Index).Branches)
        If BranchSize <> -1 Then
            For i = 0 To BranchSize
                Call SortIndex(Arr, p_Tree(Index).Branches(i))
                If i = MiddleBranch(Index) Then
                    Call Push(Arr, Index)
                    Pushed = True
                End If
            Next i
        End If
        If Pushed = False Then Call Push(Arr, Index)
    End Sub

    Private Sub Push(ByRef Arr() As Long, Value As Long)
        ReDim Preserve Arr(UboundK(Arr) + 1)
        Arr(UboundK(Arr)) = Value
    End Sub

    Private Sub DisplayX(Index As Long, Depth As Long, Messages() As String)
        Dim i As Long
        Dim CurrentMessage As String
        Dim BranchSize As Long

        If Index = -1 Then 
            ReDim Messages (0)
            Messages(0) = SetTo4Hex(p_Tree(0).Value)
            Index = 0
            Depth = 1
        End If
        BranchSize = UboundK(p_Tree(Index).Branches)

        If Depth > UboundK(Messages) And BranchSize <> -1 Then ReDim Preserve Messages(Depth)
        If BranchSize <> -1 Then
            For i = 0 To BranchSize
                CurrentMessage = CurrentMessage & SetTo4Hex(p_Tree(p_Tree(Index).Branches(i)).Value) & " "
                Call DisplayX(p_Tree(Index).Branches(i), Depth + 1, Messages)
            Next i
            Messages(Depth) = Messages(Depth) & CurrentMessage
        End If

    End Sub

    Private Sub CreateTreeRecursion(ByVal CurrentNode As Long, ByVal Width As Long, ByVal Depth As Long)
        Dim i As Long
        Dim Temp As Long
        If Depth > -1 Then
            If Width = -1 Then Width = 0
            Depth = Depth - 1
            For i = 0 To Width
                Temp = Add(CurrentNode, 0)
                Call CreateTreeRecursion(Temp, Width, Depth)
            Next
        End If
    End Sub

    Private Function ReRollBranches(Data() As Long, CompareValue As Long) As Long()
        Dim Temp() As Long
        Dim Index As Long
        Dim i As Long
        Dim j As Long

        Index = - 1
        If UboundK(Data) = -1 Then Exit Function
        For i = 0 To UboundK(Data)
            If CompareValue = Data(i) Then
                Index = i
                Exit For
            End If
        Next i

        If Index <> -1 Then
            Redim Temp(UboundK(Data) - 1)
        Else
            Redim Temp(UboundK(Data))
        End If

        For i = 0 To UboundK(Data)
            If i <> Index Then
                If CompareValue < Data(i) Then
                    Temp(j) = Data(i) - 1
                Else
                    Temp(j) = Data(i)
                End If
                j = j + 1
            End If
        Next
        ReRollBranches = Temp
    End Function

    Private Function UBoundK(Arr As Variant, Optional Dimension As Long = 1) As Long
        On Error Resume Next
        UBoundK = -1
        UBoundK = UBound(Arr, Dimension)
    End Function

    Private Function UBoundT(Tree() As Node) As Long
        On Error Resume Next
        UBoundT = -1
        UBoundT = UBound(Tree)
    End Function
'