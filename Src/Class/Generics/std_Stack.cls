VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Private Functions
    Private p_Data() As Variant
    Private p_Size As Long
'

' Events
    Public Event BeforeAdd(Value As Variant)
    Public Event AfterAdd(Value As Variant)
    Public Event BeforeDelete()
    Public Event AfterDelete()
'

' Get Let
    Public Property Let Data(n_Data() As Variant)
        p_Data = n_Data
    End Property

    Public Property Get Data() As Variant()
        Data = p_Data
    End Property

    Private Property Let Size(n_Size As Long)
        p_Size = n_Size
    End Property

    Public Property Get Size() As Long
        Size = p_Size
    End Property

    Public Property Let Value(n_Value As Variant)
        If Size <> -1 Then
            If IsObject(n_Value) Then
                Set p_Data(Size) = n_Value
            Else
                p_Data(Size) = n_Value
            End If
        End If
    End Property

    Public Property Get Value() As Variant
        If Size <> -1 Then
            If IsObject(p_Data(Size)) Then
                Set Value = p_Data(Size)
            Else
                Value = p_Data(Size)
            End If
        Else
            Set Value = Nothing
        End If
    End Property

'

' Public Functions
    Public Function Create(Optional n_Value As Variant) As std_Stack
        Set Create = New std_Stack
        If IsMissing(n_Value) = False Then Call Create.Add(n_Value)
    End Function

    Public Function Add(n_Value As Variant) As Long
        RaiseEvent BeforeAdd(n_Value)
        Size = Size + 1
        ReDim Preserve p_Data(Size)
        Value = n_Value
        Add = Size
        RaiseEvent AfterAdd(n_Value)
    End Function

    Public Function Delete() As Long
        Dim i As Long
        Dim Temp() As Variant

        RaiseEvent BeforeDelete()
        Size = Size - 1
        If Size = -1 Then
            p_Data = Temp
        Else
            Temp = p_Data
            ReDim p_Data(Size)
            For i = 0 To Size
                If IsObject(Temp(i)) Then
                    Set p_Data(i) = Temp(i)
                Else
                    p_Data(i) = Temp(i)
                End If
            Next i
        End If
        Delete = Size
        RaiseEvent AfterDelete()
    End Function
'

' Private Functions
    Private Sub Class_Initialize()
        p_Size = -1
    End Sub
'