VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private ErrorCatalog(1, 99) As Variant
Private p_Handler As New std_Error

Private p_Destination As Variant

Public Event BeforeLet(Indicator As Variant, Value As Variant)
Public Event AfterLet(Indicator As Variant, Value As Variant)


' Get Let Set And Public Functions

    Public Function Create(Optional n_Destination As Variant = Nothing, Optional n_Handler As std_Error = Nothing) As std_Config
        Dim Arr() As Variant
        Set Create = New std_Config
        If n_Destination Is Nothing Then
            Create.Destination = Arr()
        Else
            Create.Destination = n_Destination
        End If
        If n_Handler Is Nothing Then
            Create.Handler = std_Error
        Else
            Create.Handler = n_Handler
        End If
    End Function

    Public Property Let Handler(n_Handler As std_Error)
        Set p_Handler = n_Handler
    End Property

    Public Property Get IS_ERROR() As Boolean
        IS_ERROR = p_Handler.IS_ERROR
    End Property

    Public Property Let Destination(n_Destination As Variant)
        If IsObject(n_Destination) Then
            Set p_Destination = n_Destination
        Else
            p_Destination = n_Destination
        End If
    End Property

    Public Property Get Destination() As Variant
        If IsObject(p_Destination) Then
            Set Destination = p_Destination
        Else
            Destination = p_Destination
        End If
    End Property

    Public Property Let Setting(Indicator As Variant, Value As Variant)
        RaiseEvent BeforeLet(Indicator, Value)
        Select Case True
            Case TypeName(p_Destination) =    "Collection"      : Call Setting_Collection(Indicator, Value, True)
            Case TypeName(p_Destination) =    "Workbook"        : Call Setting_Workbook(Indicator, Value, True)
            Case TypeName(p_Destination) =    "Worksheet"       : Call Setting_Worksheet(p_Destination,Indicator, Value, True)
            Case TypeName(p_Destination) =    "Range"           : Call Setting_Range(p_Destination, Indicator, Value, True)
            Case TypeName(p_Destination) Like "*()"             : Call Setting_Array(Indicator, Value, True)
            Case Else                                           : Call p_Handler.Handle(ErrorCatalog, 1, True, Indicator)
        End Select
        RaiseEvent AfterLet(Indicator, Value)
    End Property

    Public Property Get Setting(Indicator As Variant) As Variant
        Dim Value As Variant
        Select Case True
            Case TypeName(p_Destination) =    "Collection"      : Setting = Setting_Collection(Indicator, Value, False)
            Case TypeName(p_Destination) =    "Workbook"        : Setting = Setting_Workbook(Indicator, Value, False)
            Case TypeName(p_Destination) =    "Worksheet"       : Setting = Setting_Worksheet(p_Destination,Indicator, Value, False)
            Case TypeName(p_Destination) =    "Range"           : Setting = Setting_Range(p_Destination, Indicator, Value, False)
            Case TypeName(p_Destination) Like "*()"             : Setting = Setting_Array(Indicator, Value, False)
            Case Else                                           : Setting = p_Handler.Handle(ErrorCatalog, 1, True, Indicator)
        End Select
    End Property
'



' Privates

    Private Function Setting_Collection(Indicator As Variant, Value As Variant, SetValue As Boolean) As Variant
        Dim i As Long
        If p_Destination.Count Mod 2 = 1 Then
            Call p_Handler.Handle(ErrorCatalog, 4, True)
        Else
            For i = 0 To p_Destination.Count Step 2
                If p_Destination(i) = Indicator Then
                    p_Destination(i + 1) = Value
                    Setting_Collection = p_Destination(i + 1)
                    Exit Function
                End If
            Next i
            Call p_Handler.Handle(ErrorCatalog, 5, True, Indicator)
        End If
    End Function

    Private Function Setting_Workbook(Indicator As Variant, Value As Variant, SetValue As Boolean) As Variant
        Dim WS As Worksheet
        For Each WS In p_Destination.Worksheets
            If Not WS.Cells.Find(Indicator) Is Nothing Then Setting_Workbook = Setting_Worksheet(WS, Indicator, Value, SetValue)
        Next WS
    End Function

    Private Function Setting_Worksheet(a_Destination As Variant, Indicator As Variant, Value As Variant, SetValue As Boolean) As Variant
        Setting_Worksheet = Setting_Range(a_Destination.Cells, Indicator, Value, SetValue)
    End Function

    Private Function Setting_Range(a_Destination As Variant, Indicator As Variant, Value As Variant, SetValue As Boolean) As Variant
        Dim WorkingCell As Range
        Dim Allowed As Boolean
        Dim i As Long

        Set WorkingCell = a_Destination.Find(Indicator)
        If Not WorkingCell Is Nothing Then
            If SetValue Then
                i = 2
                Allowed = True
                Do Until WorkingCell.Offset(0, i) = Empty
                    Allowed = False
                    If WorkingCell.Offset(0, i) = Value Then
                        Allowed = True
                        Exit Do
                    End If
                    i = i + 1
                Loop
                If Allowed Then
                    WorkingCell.Offset(0, 1) = Value
                    Setting_Range = WorkingCell.Offset(0, 1)
                Else
                    Call p_Handler.Handle(ErrorCatalog, 9, True, Indicator)
                End If
            Else
                Setting_Range = WorkingCell.Offset(0, 1)
            End If
        Else
            Call p_Handler.Handle(ErrorCatalog, 5, True, Indicator)
        End If
    End Function

    Private Function Setting_Array(Indicator As Variant, Value As Variant, SetValue As Boolean) As Variant
        Dim Dimensions As Long
        Dim i As Long
        Dim j As Long

        On Error GoTo DimensionFound
        For i = 1 To 255
            If Ubound(p_Destination, i) <> -1 Then
                Dimensions = i
            End If
        Next i
        DimensionFound:

        Select Case Dimensions
            Case 0:
                Call p_Handler.Handle(ErrorCatalog, 3, True)
            Case 1
                If Ubound(p_Destination, 1) Mod 2 = 1 Then
                    Call p_Handler.Handle(ErrorCatalog, 4, True)
                Else
                    For i = 0 To Ubound(p_Destination, 1) Step 2
                        If p_Destination(i) = Indicator Then
                            p_Destination(i + 1) = Value
                            Setting_Array = p_Destination(i + 1)
                            Exit Function
                        End If
                    Next i
                    Call p_Handler.Handle(ErrorCatalog, 5, True, Indicator)
                End If
            Case 2
                If Ubound(p_Destination, 2) >= 1 Then
                    For i = 0 To Ubound(p_Destination, 1)
                        If p_Destination(i, 0) = Indicator Then
                            If SetValue Then
                                Dim Allowed As Boolean
                                Allowed = True
                                For j = 2 To Ubound(p_Destination, 2)
                                    Allowed = False
                                    If p_Destination(i, j) = Value Then
                                        Allowed = True
                                        Exit For
                                    End If                    
                                Next j
                                If Allowed Then
                                    p_Destination(i, 1) = Value
                                    Setting_Array = p_Destination(i, 1)
                                    Exit Function
                                End If
                                Allowed = False
                            Else
                                Setting_Array = p_Destination(i, 1)
                                Exit Function
                            End If
                        End If
                    Next i
                    Call p_Handler.Handle(ErrorCatalog, 5, True, Indicator)
                Else
                    Call p_Handler.Handle(ErrorCatalog, 4, True)
                End If
            Case Else
                Call p_Handler.Handle(ErrorCatalog, 7, True)
        End Select
        
    End Function

'


' Private Procedures
    Private Sub Class_Initialize()
        ProtInit
    End Sub 

    Private Sub Class_Terminate()
    End Sub


    Private Sub ProtInit()
        ErrorCatalog(0, 0000) = 0002: ErrorCatalog(1, 0000) = "std_Config"
        ErrorCatalog(0, 0001) = 1000: ErrorCatalog(1, 0001) = "Invalid Destination"
        ErrorCatalog(0, 0002) = 1000: ErrorCatalog(1, 0002) = "Invalid Value"
        ErrorCatalog(0, 0003) = 1000: ErrorCatalog(1, 0003) = "Destination as Array has no Dimension"
        ErrorCatalog(0, 0004) = 1000: ErrorCatalog(1, 0004) = "Destination as Array has odd size"
        ErrorCatalog(0, 0005) = 1000: ErrorCatalog(1, 0005) = "Could not find Indicator"
        ErrorCatalog(0, 0006) = 1000: ErrorCatalog(1, 0006) = "Can not Assign Object to non-Object Type"
        ErrorCatalog(0, 0007) = 1000: ErrorCatalog(1, 0007) = "Destination as Array has unallowed Dimension(s)"
        ErrorCatalog(0, 0008) = 1000: ErrorCatalog(1, 0008) = "Destination as Collection has odd size"
        ErrorCatalog(0, 0009) = 1000: ErrorCatalog(1, 0009) = "Value not in List of allowed Values"
        ErrorCatalog(0, 0010) = 1000: ErrorCatalog(1, 0010) = "PLACEHOLDER"
        ErrorCatalog(0, 0011) = 1000: ErrorCatalog(1, 0011) = "PLACEHOLDER"
        ErrorCatalog(0, 0012) = 1000: ErrorCatalog(1, 0012) = "PLACEHOLDER"
        ErrorCatalog(0, 0013) = 1000: ErrorCatalog(1, 0013) = "PLACEHOLDER"
        ErrorCatalog(0, 0014) = 1000: ErrorCatalog(1, 0014) = "PLACEHOLDER"
        ErrorCatalog(0, 0015) = 1000: ErrorCatalog(1, 0015) = "PLACEHOLDER"
        ErrorCatalog(0, 0016) = 1000: ErrorCatalog(1, 0016) = "PLACEHOLDER"
        ErrorCatalog(0, 0017) = 1000: ErrorCatalog(1, 0017) = "PLACEHOLDER"
        ErrorCatalog(0, 0018) = 1000: ErrorCatalog(1, 0018) = "PLACEHOLDER"
        ErrorCatalog(0, 0019) = 1000: ErrorCatalog(1, 0019) = "PLACEHOLDER"
        ErrorCatalog(0, 0020) = 1000: ErrorCatalog(1, 0020) = "PLACEHOLDER"
        ErrorCatalog(0, 0021) = 1000: ErrorCatalog(1, 0021) = "PLACEHOLDER"
        ErrorCatalog(0, 0022) = 1000: ErrorCatalog(1, 0022) = "PLACEHOLDER"
        ErrorCatalog(0, 0023) = 1000: ErrorCatalog(1, 0023) = "PLACEHOLDER"
        ErrorCatalog(0, 0024) = 1000: ErrorCatalog(1, 0024) = "PLACEHOLDER"
        ErrorCatalog(0, 0025) = 1000: ErrorCatalog(1, 0025) = "PLACEHOLDER"
        ErrorCatalog(0, 0026) = 1000: ErrorCatalog(1, 0026) = "PLACEHOLDER"
        ErrorCatalog(0, 0027) = 1000: ErrorCatalog(1, 0027) = "PLACEHOLDER"
        ErrorCatalog(0, 0028) = 1000: ErrorCatalog(1, 0028) = "PLACEHOLDER"
        ErrorCatalog(0, 0029) = 1000: ErrorCatalog(1, 0029) = "PLACEHOLDER"
        ErrorCatalog(0, 0030) = 1000: ErrorCatalog(1, 0030) = "PLACEHOLDER"
        ErrorCatalog(0, 0031) = 1000: ErrorCatalog(1, 0031) = "PLACEHOLDER"
        ErrorCatalog(0, 0032) = 1000: ErrorCatalog(1, 0032) = "PLACEHOLDER"
        ErrorCatalog(0, 0033) = 1000: ErrorCatalog(1, 0033) = "PLACEHOLDER"
        ErrorCatalog(0, 0034) = 1000: ErrorCatalog(1, 0034) = "PLACEHOLDER"
        ErrorCatalog(0, 0035) = 1000: ErrorCatalog(1, 0035) = "PLACEHOLDER"
        ErrorCatalog(0, 0036) = 1000: ErrorCatalog(1, 0036) = "PLACEHOLDER"
        ErrorCatalog(0, 0037) = 1000: ErrorCatalog(1, 0037) = "PLACEHOLDER"
        ErrorCatalog(0, 0038) = 1000: ErrorCatalog(1, 0038) = "PLACEHOLDER"
        ErrorCatalog(0, 0039) = 1000: ErrorCatalog(1, 0039) = "PLACEHOLDER"
        ErrorCatalog(0, 0040) = 1000: ErrorCatalog(1, 0040) = "PLACEHOLDER"
        ErrorCatalog(0, 0041) = 1000: ErrorCatalog(1, 0041) = "PLACEHOLDER"
        ErrorCatalog(0, 0042) = 1000: ErrorCatalog(1, 0042) = "PLACEHOLDER"
        ErrorCatalog(0, 0043) = 1000: ErrorCatalog(1, 0043) = "PLACEHOLDER"
        ErrorCatalog(0, 0044) = 1000: ErrorCatalog(1, 0044) = "PLACEHOLDER"
        ErrorCatalog(0, 0045) = 1000: ErrorCatalog(1, 0045) = "PLACEHOLDER"
        ErrorCatalog(0, 0046) = 1000: ErrorCatalog(1, 0046) = "PLACEHOLDER"
        ErrorCatalog(0, 0047) = 1000: ErrorCatalog(1, 0047) = "PLACEHOLDER"
        ErrorCatalog(0, 0048) = 1000: ErrorCatalog(1, 0048) = "PLACEHOLDER"
        ErrorCatalog(0, 0049) = 1000: ErrorCatalog(1, 0049) = "PLACEHOLDER"
        ErrorCatalog(0, 0050) = 1000: ErrorCatalog(1, 0050) = "PLACEHOLDER"
        ErrorCatalog(0, 0051) = 1000: ErrorCatalog(1, 0051) = "PLACEHOLDER"
        ErrorCatalog(0, 0052) = 1000: ErrorCatalog(1, 0052) = "PLACEHOLDER"
        ErrorCatalog(0, 0053) = 1000: ErrorCatalog(1, 0053) = "PLACEHOLDER"
        ErrorCatalog(0, 0054) = 1000: ErrorCatalog(1, 0054) = "PLACEHOLDER"
        ErrorCatalog(0, 0055) = 1000: ErrorCatalog(1, 0055) = "PLACEHOLDER"
        ErrorCatalog(0, 0056) = 1000: ErrorCatalog(1, 0056) = "PLACEHOLDER"
        ErrorCatalog(0, 0057) = 1000: ErrorCatalog(1, 0057) = "PLACEHOLDER"
        ErrorCatalog(0, 0058) = 1000: ErrorCatalog(1, 0058) = "PLACEHOLDER"
        ErrorCatalog(0, 0059) = 1000: ErrorCatalog(1, 0059) = "PLACEHOLDER"
        ErrorCatalog(0, 0060) = 1000: ErrorCatalog(1, 0060) = "PLACEHOLDER"
        ErrorCatalog(0, 0061) = 1000: ErrorCatalog(1, 0061) = "PLACEHOLDER"
        ErrorCatalog(0, 0062) = 1000: ErrorCatalog(1, 0062) = "PLACEHOLDER"
        ErrorCatalog(0, 0063) = 1000: ErrorCatalog(1, 0063) = "PLACEHOLDER"
        ErrorCatalog(0, 0064) = 1000: ErrorCatalog(1, 0064) = "PLACEHOLDER"
        ErrorCatalog(0, 0065) = 1000: ErrorCatalog(1, 0065) = "PLACEHOLDER"
        ErrorCatalog(0, 0066) = 1000: ErrorCatalog(1, 0066) = "PLACEHOLDER"
        ErrorCatalog(0, 0067) = 1000: ErrorCatalog(1, 0067) = "PLACEHOLDER"
        ErrorCatalog(0, 0068) = 1000: ErrorCatalog(1, 0068) = "PLACEHOLDER"
        ErrorCatalog(0, 0069) = 1000: ErrorCatalog(1, 0069) = "PLACEHOLDER"
        ErrorCatalog(0, 0070) = 1000: ErrorCatalog(1, 0070) = "PLACEHOLDER"
        ErrorCatalog(0, 0071) = 1000: ErrorCatalog(1, 0071) = "PLACEHOLDER"
        ErrorCatalog(0, 0072) = 1000: ErrorCatalog(1, 0072) = "PLACEHOLDER"
        ErrorCatalog(0, 0073) = 1000: ErrorCatalog(1, 0073) = "PLACEHOLDER"
        ErrorCatalog(0, 0074) = 1000: ErrorCatalog(1, 0074) = "PLACEHOLDER"
        ErrorCatalog(0, 0075) = 1000: ErrorCatalog(1, 0075) = "PLACEHOLDER"
        ErrorCatalog(0, 0076) = 1000: ErrorCatalog(1, 0076) = "PLACEHOLDER"
        ErrorCatalog(0, 0077) = 1000: ErrorCatalog(1, 0077) = "PLACEHOLDER"
        ErrorCatalog(0, 0078) = 1000: ErrorCatalog(1, 0078) = "PLACEHOLDER"
        ErrorCatalog(0, 0079) = 1000: ErrorCatalog(1, 0079) = "PLACEHOLDER"
        ErrorCatalog(0, 0080) = 1000: ErrorCatalog(1, 0080) = "PLACEHOLDER"
        ErrorCatalog(0, 0081) = 1000: ErrorCatalog(1, 0081) = "PLACEHOLDER"
        ErrorCatalog(0, 0082) = 1000: ErrorCatalog(1, 0082) = "PLACEHOLDER"
        ErrorCatalog(0, 0083) = 1000: ErrorCatalog(1, 0083) = "PLACEHOLDER"
        ErrorCatalog(0, 0084) = 1000: ErrorCatalog(1, 0084) = "PLACEHOLDER"
        ErrorCatalog(0, 0085) = 1000: ErrorCatalog(1, 0085) = "PLACEHOLDER"
        ErrorCatalog(0, 0086) = 1000: ErrorCatalog(1, 0086) = "PLACEHOLDER"
        ErrorCatalog(0, 0087) = 1000: ErrorCatalog(1, 0087) = "PLACEHOLDER"
        ErrorCatalog(0, 0088) = 1000: ErrorCatalog(1, 0088) = "PLACEHOLDER"
        ErrorCatalog(0, 0089) = 1000: ErrorCatalog(1, 0089) = "PLACEHOLDER"
        ErrorCatalog(0, 0090) = 1000: ErrorCatalog(1, 0090) = "PLACEHOLDER"
        ErrorCatalog(0, 0091) = 1000: ErrorCatalog(1, 0091) = "PLACEHOLDER"
        ErrorCatalog(0, 0092) = 1000: ErrorCatalog(1, 0092) = "PLACEHOLDER"
        ErrorCatalog(0, 0093) = 1000: ErrorCatalog(1, 0093) = "PLACEHOLDER"
        ErrorCatalog(0, 0094) = 1000: ErrorCatalog(1, 0094) = "PLACEHOLDER"
        ErrorCatalog(0, 0095) = 1000: ErrorCatalog(1, 0095) = "PLACEHOLDER"
        ErrorCatalog(0, 0096) = 1000: ErrorCatalog(1, 0096) = "PLACEHOLDER"
        ErrorCatalog(0, 0097) = 1000: ErrorCatalog(1, 0097) = "PLACEHOLDER"
        ErrorCatalog(0, 0098) = 1000: ErrorCatalog(1, 0098) = "PLACEHOLDER"
        ErrorCatalog(0, 0099) = 1000: ErrorCatalog(1, 0099) = "PLACEHOLDER"
    End Sub
'