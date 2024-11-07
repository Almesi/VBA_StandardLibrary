VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Matrix_dbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private p_Matrix() As Double

Public Enum MatrixOperator
    Equal = 0
    NotEqual = 1
    Bigger = 2
    Smaller = 3
    BiggerEqual = 4
    SmallerEqual = 5
    IsObj = 6
    NotObj = 7
End Enum

' Let Get

    Public Property Let Item(ByVal Row As Long, ByVal Column As Long, ByVal Value As Double)
        p_Matrix(Row, Column) = Value
    End Property

    Public Property Get Item(ByVal Row As Long, ByVal Column As Long) As Double
        Item = p_Matrix(Row, Column)
    End Property

    Public Property Get RowCount() As Long
        RowCount = Ubound(p_Matrix, 1)
    End Property

    Public Property Get ColumnCount() As Long
        ColumnCount = Ubound(p_Matrix, 2)
    End Property

    Public Property Let Matrix(Data() As Double)
        p_Matrix = Data
    End Property

    Public Property Get Matrix() As Double()
        Matrix = p_Matrix
    End Property
'

' Creates
    Public Function Create(Optional RowCount As Long = 0, Optional ColumnCount As Long = 0) As std_Matrix_dbl
        Set Create = New std_Matrix_dbl
        Call SetMatrixSize(Create, RowCount, ColumnCount)
    End Function

    Public Function CreateFromArray_Column(Data As Variant, Count As Long) As std_Matrix_dbl
        Dim i As Long
        Dim Rest As Long     : Rest = Ubound(Data) Mod Count
        Dim RowCount As Long : RowCount = Int(Ubound(Data) / Count)
        If Rest <> 0 Then RowCount = RowCount + 1
        Set CreateFromArray_Column = Create(Count - 1, RowCount - 1)
        For i = 0 To Ubound(Data)
            CreateFromArray_Column.Item(Int(i / (Count)), i Mod (Count)) = Data(i)
        Next i
    End Function

    Public Function CreateFromArray_Row(Data As Variant, Count As Long) As std_Matrix_dbl
        Dim i As Long
        Dim Rest As Long     : Rest = Ubound(Data) Mod Count
        Dim RowCount As Long : RowCount = Int(Ubound(Data) / Count)
        If Rest <> 0 Then RowCount = RowCount + 1
        Set CreateFromArray_Row = Create(RowCount - 1, Count - 1)
        For i = 0 To Ubound(Data)
            CreateFromArray_Row.Item(i Mod (Count), Int(i / (Count))) = Data(i)
        Next i
    End Function

    Public Function CreateRotationX(Dimensions As Long, Theta As Double) As std_Matrix_dbl
        Dim i As Long
        Set CreateRotationX = New std_Matrix_dbl
        Call SetMatrixSize(CreateRotationX, Dimensions, Dimensions)
        For i = 0 To Dimensions
            CreateRotationX.Item(i, i) = 1
        Next
        If Dimensions >= 1 Then
            CreateRotationX.Item(1, 1) =  Cos(Theta)
        End If
        If Dimensions > 1 Then
            CreateRotationX.Item(1, 2) = -Sin(Theta)
            CreateRotationX.Item(2, 1) =  Sin(Theta)
            CreateRotationX.Item(2, 2) =  Cos(Theta)
        End If
    End Function

    Public Function CreateRotationY(Dimensions As Long, Theta As Double) As std_Matrix_dbl
        Dim i As Long
        Set CreateRotationY = Create(Dimensions, Dimensions)
        For i = 0 To Dimensions
            CreateRotationY.Item(i, i) = 1
        Next
        If Dimensions >= 1 Then
            CreateRotationY.Item(0, 0) =  Cos(Theta)
        End If
        If Dimensions > 1 Then
            CreateRotationY.Item(0, 2) =  Sin(Theta)
            CreateRotationY.Item(2, 0) = -Sin(Theta)
            CreateRotationY.Item(2, 2) =  Cos(Theta)
        End If
    End Function

    Public Function CreateRotationZ(Dimensions As Long, Theta As Double) As std_Matrix_dbl
        Dim i As Long
        Set CreateRotationZ = Create(Dimensions, Dimensions)
        For i = 0 To Dimensions
            CreateRotationZ.Item(i, i) = 1
        Next
        If Dimensions > 1 Then
            CreateRotationZ.Item(0, 0) =  Cos(Theta)
            CreateRotationZ.Item(0, 1) = -Sin(Theta)
            CreateRotationZ.Item(1, 0) =  Sin(Theta)
            CreateRotationZ.Item(1, 1) =  Cos(Theta)
        End If
    End Function

    Public Function CreateRotation(Dimensions As Long, ThetaX As Double, ThetaY As Double, ThetaZ As Double) As std_Matrix_dbl
        Dim x As std_Matrix_dbl    : Set x = CreateRotationX(Dimensions, ThetaX)
        Dim y As std_Matrix_dbl    : Set y = CreateRotationY(Dimensions, ThetaY)
        Dim z As std_Matrix_dbl    : Set z = CreateRotationZ(Dimensions, ThetaZ)
        Dim Temp As std_Matrix_dbl : Set Temp = Create(Dimensions, Dimensions)
        Set CreateRotation = Create(Dimensions, Dimensions)
        Call std_Matrix_dbl.Multiply(Temp, y, z)
        Call std_Matrix_dbl.Multiply(CreateRotation, x, Temp)
    End Function
'

' Functions
    Public Sub Add(Goal As std_Matrix_dbl, Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl)
        Dim i As Long, j As Long
        If CompareRow(Matrix1, Matrix2)    = False Then Exit Sub
        If CompareColumn(Matrix1, Matrix2) = False Then Exit Sub
        Call SetMatrixSize(Goal, Matrix1.RowCount, Matrix2.ColumnCount)
        For i = 0 To Goal.RowCount
            For j = 0 To Goal.ColumnCount
                Goal.Item(i, j) = Matrix1.Item(i, j) + Matrix2.Item(i, j)
            Next j
        Next i
    End Sub

    Public Sub Subtract(Goal As std_Matrix_dbl, Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl)
        Dim i As Long, j As Long
        If CompareRow(Matrix1, Matrix2)    = False Then Exit Sub
        If CompareColumn(Matrix1, Matrix2) = False Then Exit Sub
        Call SetMatrixSize(Goal, Matrix1.RowCount, Matrix2.ColumnCount)
        For i = 0 To Goal.RowCount
            For j = 0 To Goal.ColumnCount
                Goal.Item(i, j) = Matrix1.Item(i, j) - Matrix2.Item(i, j)
            Next j
        Next i
    End Sub

    Public Sub Multiply(Goal As std_Matrix_dbl, Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl)
        Dim i As Long, j As Long, k As Long
        Dim Temp As std_Matrix_dbl

        Set Temp = std_Matrix_dbl.Create(Matrix1.RowCount, Matrix2.ColumnCount)
        For i = 0 To Matrix1.RowCount
            For j = 0 To Matrix2.ColumnCount
                For k = 0 To Matrix1.ColumnCount
                    Temp.Item(i, j) = Temp.Item(i, j) + Matrix1.Item(i, k) * Matrix2.Item(k, j)
                Next k
            Next j
        Next i
        Goal.Matrix = Temp.Matrix
    End Sub

    Public Sub Scaling(Goal As std_Matrix_dbl, Value As Double)
        Dim i As Long, j As Long
        For i = 0 To Goal.RowCount
            For j = 0 To Goal.ColumnCount
                Goal.Item(i, j) = Goal.Item(i, j) * Value
            Next j
        Next i
    End Sub

    Public Sub Translate(Goal As std_Matrix_dbl, Value As Double)
        Dim i As Long, j As Long
        For i = 0 To Goal.RowCount
            For j = 0 To Goal.ColumnCount
                Goal.Item(i, j) = Goal.Item(i, j) + Value
            Next j
        Next i
    End Sub

    Public Sub Divide(Goal As std_Matrix_dbl, ByRef Matrix1 As std_Matrix_dbl, ByVal Matrix2 As std_Matrix_dbl)
        Dim i As Long, j As Long, k As Long
        Dim Temp As std_Matrix_dbl

        Set Temp = std_Matrix_dbl.Create(Matrix1.RowCount, Matrix2.ColumnCount)
        For i = 0 To Matrix1.RowCount
            For j = 0 To Matrix2.ColumnCount
                For k = 0 To Matrix1.ColumnCount
                    Temp.Item(i, j) = Matrix1.Item(i, k) / Matrix2.Item(k, j)
                Next k
            Next j
        Next i
        Goal.Matrix = Temp.Matrix
    End Sub

    Public Function Compare(Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl, Operator As MatrixOperator) As Boolean
        Dim i As Long, j As Long
        Select Case Operator
            Case IsObj  : Compare = Matrix1 Is Matrix2
            Case NotObj : Compare = Not Matrix1 Is Matrix2
            Case Else
            If CompareRow(Matrix1, Matrix2)    = False Then Exit Function
            If CompareColumn(Matrix1, Matrix2) = False Then Exit Function
            Compare = True
            For i = 0 To Matrix1.RowCount
                For j = 0 To Matrix2.ColumnCount
                    Select Case Operator
                        Case Equal        : If Matrix1.Item(i, j) <> Matrix2.Item(i, j) Then Compare = False: Exit Function
                        Case NotEqual     : If Matrix1.Item(i, j) =  Matrix2.Item(i, j) Then Compare = False: Exit Function
                        Case Bigger       : If Matrix1.Item(i, j) >= Matrix2.Item(i, j) Then Compare = False: Exit Function
                        Case Smaller      : If Matrix1.Item(i, j) =< Matrix2.Item(i, j) Then Compare = False: Exit Function
                        Case BiggerEqual  : If Matrix1.Item(i, j) >  Matrix2.Item(i, j) Then Compare = False: Exit Function
                        Case SmallerEqual : If Matrix1.Item(i, j) <  Matrix2.Item(i, j) Then Compare = False: Exit Function
                    End Select
                Next j
            Next i
        End Select
    End Function

    Public Sub Transpose(Goal As std_Matrix_dbl)
        Dim i As Long, j As Long
        Dim Temp As std_Matrix_dbl
        Set Temp = Create(Goal.ColumnCount, Goal.RowCount)
        For i = 0 To Goal.ColumnCount
            For j = 0 To Goal.RowCount
                Temp.Item(i, j) = Goal.Item(j, i)
            Next j
        Next i
        Goal.Matrix = Temp.Matrix
    End Sub

    Public Sub Inverse(Goal As std_Matrix_dbl)
        Dim i As Long, j As Long
        With Goal
            For i = 0 To RowCount
                For j = 0 To ColumnCount
                    .Matrix(i, j) = 1 / .Matrix(i, j)
                Next j
            Next i
        Next
    End Sub

    Public Sub Display()
        Dim i As Long, j As Long
        Dim Message As String
        For i = 0 To Ubound(Matrix, 1)
            For j = 0 To Ubound(Matrix, 2)
                Message = Message & p_Matrix(i, j) & "|"
            Next j
            Debug.Print Message
            Message = ""
        Next i
    End Sub
'

' Private
    Private Function CompareRow(Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl) As Boolean
        If Matrix1.RowCount    = Matrix2.RowCount    Then CompareRow = True
    End Function

    Private Function CompareColumn(Matrix1 As std_Matrix_dbl, Matrix2 As std_Matrix_dbl) As Boolean
        If Matrix1.ColumnCount = Matrix2.ColumnCount Then CompareColumn = True
    End Function

    Private Sub SetMatrixSize(Matr As std_Matrix_dbl, RowCount As Long, ColumnCount As Long)
        Dim Temp() As Double
        ReDim Temp(RowCount, ColumnCount)
        Matr.Matrix = Temp
    End Sub
'