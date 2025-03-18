VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Mesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)

Private Const pi As Single = 3.14159

Private p_Data() As Byte
Private p_Type As Long
Private p_Count As Long
Private p_Size As Long
Private p_VertexSize As Long

Public Enum std_MeshType
    NoPoint    = 00
    Point      = 01
    Line       = 02
    Triangle   = 03
    Rectangle  = 04
    Pentagon   = 05
    Hexagon    = 06
    Variable   = 07
    Pyramid    = 08
    Box        = 09
    Variable3D = 10
End Enum

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub


Public Property Let Data(n_Data() As Byte)
    p_Data = n_Data
End Property
Public Property Get Data() As Byte()
    Data = p_Data
End Property

Public Property Get Pointer() As LongPtr
    If UBound(p_Data) >= 0 Then Pointer = VarPtr(p_Data(0))
End Property

Public Property Get Typee()      As Long: Typee      = p_Type      : End Property
Public Property Get Count()      As Long: Count      = p_Count     : End Property
Public Property Get Size()       As Long: Size       = p_Size      : End Property
Public Property Get VertexSize() As Long: VertexSize = p_VertexSize: End Property 


Public Sub CreateStandardMesh(MeshType As std_MeshType, Optional VariableSize As Long = 0)
    Dim MeshData() As Single
    Dim i As Long
    Dim NewSize  As Long
    Dim PiFactor As Single
    Dim Radian   As Single
    NewSize  = VariableSize * 3
    PiFactor = 3.14 / 180
    Select Case MeshType
        Case std_MeshType.NoPoint
        Case std_MeshType.Point
            ReDim MeshData(2)
        Case std_MeshType.Line
            ReDim MeshData(5)
            MeshData(0) = 0: MeshData(1) = 0: MeshData(2) = 0
            MeshData(3) = 1: MeshData(4) = 0: MeshData(5) = 0
        Case std_MeshType.Triangle
            ReDim MeshData(8)
            MeshData(0) = 0: MeshData(1) = 0: MeshData(2) = 0
            MeshData(3) = 1: MeshData(4) = 0: MeshData(5) = 0
            MeshData(6) = 1: MeshData(7) = 1: MeshData(8) = 0
        Case std_MeshType.Rectangle
            ReDim MeshData(11)
            MeshData(00) = 0: MeshData(01) = 0: MeshData(02) = 0
            MeshData(03) = 1: MeshData(04) = 0: MeshData(05) = 0
            MeshData(06) = 1: MeshData(07) = 1: MeshData(08) = 0
            MeshData(09) = 0: MeshData(10) = 1: MeshData(11) = 0
        Case std_MeshType.Pentagon
            ReDim MeshData(14)
            MeshData(00) = +0.25: MeshData(01) = -1.00: MeshData(02) = +0.00
            MeshData(03) = +1.00: MeshData(04) = +0.00: MeshData(05) = +0.00
            MeshData(06) = +0.00: MeshData(07) = +1.00: MeshData(08) = +0.00
            MeshData(09) = -1.00: MeshData(10) = +0.00: MeshData(11) = +0.00
            MeshData(12) = -0.25: MeshData(13) = -1.00: MeshData(14) = +0.00
        Case std_MeshType.Hexagon
            ReDim MeshData(17)
            MeshData(00) = +0.00: MeshData(01) = -1.00: MeshData(02) = +0.00
            MeshData(03) = +1.00: MeshData(04) = -0.25: MeshData(05) = +0.00
            MeshData(06) = +1.00: MeshData(07) = +0.25: MeshData(08) = +0.00
            MeshData(09) = +0.00: MeshData(10) = +1.00: MeshData(11) = +0.00
            MeshData(12) = -1.00: MeshData(13) = +0.25: MeshData(14) = +0.00
            MeshData(15) = -1.00: MeshData(16) = -0.25: MeshData(17) = +0.00
        Case std_MeshType.Variable
            If VariableSize < 3 Then Exit Sub
            ReDim MeshData(NewSize - 1)
            For i = 1 To VariableSize
                Radian = i * (360 / VariableSize) * PiFactor
                MeshData(3 * i - 3    ) = CSng(Sin(Radian))
                MeshData(3 * i - 3 + 1) = CSng(Cos(Radian))
                MeshData(3 * i - 3 + 2) = 0
            Next i
        Case std_MeshType.Pyramid
            ReDim MeshData(14)
            MeshData(00) = +1.00: MeshData(01) = +0.00: MeshData(02) = -1.00
            MeshData(03) = -1.00: MeshData(04) = +0.00: MeshData(05) = -1.00
            MeshData(06) = +1.00: MeshData(07) = +0.00: MeshData(08) = +1.00
            MeshData(09) = +1.00: MeshData(10) = +0.00: MeshData(11) = +1.00
            MeshData(12) = +0.00: MeshData(13) = +1.00: MeshData(14) = +0.00
        Case std_MeshType.Box
            ReDim MeshData(23)
            MeshData(00) = -1.00: MeshData(01) = -1.00: MeshData(02) = -1.00
            MeshData(03) = +1.00: MeshData(04) = -1.00: MeshData(05) = -1.00
            MeshData(06) = +1.00: MeshData(07) = +1.00: MeshData(08) = -1.00
            MeshData(09) = -1.00: MeshData(10) = +1.00: MeshData(11) = -1.00
            MeshData(12) = -1.00: MeshData(13) = -1.00: MeshData(14) = +1.00
            MeshData(15) = +1.00: MeshData(16) = -1.00: MeshData(17) = +1.00
            MeshData(18) = +1.00: MeshData(19) = +1.00: MeshData(20) = +1.00
            MeshData(21) = -1.00: MeshData(22) = +1.00: MeshData(23) = +1.00
        Case std_MeshType.Variable3D
            If VariableSize < 6 Then Exit Sub
            'MeshData = SphereMesh(VariableSize)
    End Select
    If Ubound(MeshData) >= 0 Then
        Call AssignData(VarPtr(MeshData(0)), VarType(MeshData(0)), LenB(MeshData(0)), UBound(MeshData) + 1, 3)
    End If

End Sub

Public Sub CreateStandardMeshIndex(MeshType As std_MeshType, Optional VariableSize As Long = 0)
    Dim MeshData() As Single
    Dim i As Long
    Dim NewSize  As Long
    NewSize  = (VariableSize - 2) * 3
    Select Case MeshType
        Case std_MeshType.NoPoint
        Case std_MeshType.Point
            ReDim MeshData(0)
        Case std_MeshType.Line
            ReDim MeshData(1)
            MeshData(0) = 0: MeshData(1) = 1
        Case std_MeshType.Triangle
            ReDim MeshData(2)
            MeshData(0) = 0: MeshData(1) = 1: MeshData(2) = 2
        Case std_MeshType.Rectangle
            ReDim MeshData(5)
            MeshData(00) = 0: MeshData(01) = 3: MeshData(02) =  1
            MeshData(03) = 3: MeshData(04) = 2: MeshData(05) =  1
        Case std_MeshType.Pentagon
            ReDim MeshData(8)
            MeshData(00) = 0: MeshData(01) = 2: MeshData(02) =  1
            MeshData(03) = 0: MeshData(04) = 3: MeshData(05) =  2
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) =  3
        Case std_MeshType.Hexagon
            ReDim MeshData(11)
            MeshData(00) = 0: MeshData(01) = 2: MeshData(02) = 1
            MeshData(03) = 0: MeshData(04) = 3: MeshData(05) = 2
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) = 3
            MeshData(09) = 0: MeshData(10) = 5: MeshData(11) = 4
        Case std_MeshType.Variable
            If VariableSize = < 3 Then Exit Sub
            ReDim MeshData(NewSize - 1)
            For i = 0 To VariableSize - 1
                MeshData(3 * i    ) = 0
                MeshData(3 * i + 1) = i + 2
                MeshData(3 * i + 2) = i + 2 - 1
            Next i
        Case std_MeshType.Pyramid
            ReDim MeshData(17)
            MeshData(00) = 0: MeshData(01) = 3: MeshData(02) = 1
            MeshData(03) = 0: MeshData(04) = 2: MeshData(05) = 3
            MeshData(06) = 0: MeshData(07) = 4: MeshData(08) = 1
            MeshData(09) = 1: MeshData(10) = 4: MeshData(11) = 3
            MeshData(12) = 3: MeshData(13) = 4: MeshData(14) = 2
            MeshData(15) = 2: MeshData(16) = 4: MeshData(17) = 0
        Case std_MeshType.Box
            ReDim MeshData(35)
            MeshData(00) = 0: MeshData(01) =  2: MeshData(02) =  1
            MeshData(03) = 0: MeshData(04) =  3: MeshData(05) =  2
            MeshData(06) = 1: MeshData(07) =  2: MeshData(08) =  6
            MeshData(09) = 1: MeshData(10) =  6: MeshData(11) =  5
            MeshData(12) = 5: MeshData(13) =  6: MeshData(14) =  7
            MeshData(15) = 5: MeshData(16) =  7: MeshData(17) =  4
            MeshData(18) = 4: MeshData(19) =  7: MeshData(20) =  3
            MeshData(21) = 4: MeshData(22) =  3: MeshData(23) =  0
            MeshData(24) = 4: MeshData(25) =  0: MeshData(26) =  1
            MeshData(27) = 4: MeshData(28) =  1: MeshData(29) =  5
            MeshData(30) = 3: MeshData(31) =  7: MeshData(32) =  6
            MeshData(33) = 3: MeshData(34) =  6: MeshData(35) =  2
        Case std_MeshType.Variable3D
            If VariableSize < 6 Then Exit Sub
            ' MeshData = SphereMesh(VariableSize)
    End Select
    If Ubound(MeshData) >= 0 Then
        Call AssignData(VarPtr(MeshData(0)), VarType(MeshData(0)), LenB(MeshData(0)), UBound(MeshData) + 1, 3)
    End If
End Sub

Public Sub AssignData(DataPointer As LongPtr, DataType As Long, DataSize As Long, DataCount As Long, n_VertexSize As Long)
    Dim Temp() As Byte
    p_Type = DataType
    p_Count = DataCount
    p_Size = DataSize
    p_Data = Temp
    p_VertexSize = n_VertexSize
    ReDim p_Data(DataSize * DataCount - 1)
    Call CopyMemory(p_Data(0), DataPointer, DataSize * DataCount)
End Sub

' AttributeData has needs either as many elements as ItemsPerVertex, which will then be used for every input
' Or It has to have enough elements for each input as seperate data
Public Sub AddAttribute(ItemsPerVertex As Long, AttributeData() As Single)
    Dim TempData() As Single
    Dim VertexCount As Long
    Dim i As Long, j As Long
    Dim NewVertexCount As Long
    Dim Offset As Long
    Dim Index As Long

    VertexCount = p_Count / VertexSize
    NewVertexCount = p_VertexSize + ItemsPerVertex
    ReDim TempData(VertexCount * NewVertexCount - 1)
    For i = 0 To VertexCount - 1
        Offset = p_VertexSize * i
        Call CopyMemory(TempData(ItemsPerVertex * i + Offset    ), VarPtr(p_Data(Offset * Size + (Size * 0))), Size)
        Call CopyMemory(TempData(ItemsPerVertex * i + Offset + 1), VarPtr(p_Data(Offset * Size + (Size * 1))), Size)
        Call CopyMemory(TempData(ItemsPerVertex * i + Offset + 2), VarPtr(p_Data(Offset * Size + (Size * 2))), Size)
        If Ubound(AttributeData) + 1 = ItemsPerVertex Then
            For j = 1 To ItemsPerVertex
                TempData(ItemsPerVertex * i + Offset + 2 + j) = AttributeData(j - 1)
            Next j
        Else
            For j = 1 To ItemsPerVertex
                TempData(ItemsPerVertex * i + Offset + 2 + j) = AttributeData(Index)
                Index = Index + 1
            Next j
        End If
    Next i
    Call AssignData(VarPtr(TempData(0)), VarType(TempData(0)), LenB(TempData(0)), Ubound(TempData) + 1, NewVertexCount)
End Sub

Private Function SphereMesh(Slices As Long, Stacks As Long) As Single()

End Function