VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ColorValueForm 
   Caption         =   "Change Value"
   ClientHeight    =   4665
   ClientWidth     =   8460.001
   OleObjectBlob   =   "ColorValueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ColorValueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model As ValueModel
End Type

Private This As TView
Private Updating As Boolean

Public Function Create() As ColorValueForm
    Set Create = New ColorValueForm
End Function


Private Sub OKButton_Click()
    This.Model.Value = ValueBox.Value
    This.ReturnValue = FormReturnOK
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    This.ReturnValue = FormReturnCancel
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.ReturnValue = FormReturnCancel
    Hide
End Sub


Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    This.ReturnValue = FormReturnCancel

   Call SetRGB(0, 0, 0)

    Me.Show
    IView_Show = This.ReturnValue
End Function

Private Property Let IView_Result(ByVal Value As FormReturn)
    This.ReturnValue = Value
End Property

Private Property Get IView_Form() As Object
    Set IView_ForM = Me
End Property


Private Sub RedValue_Change()            : RGBTextChanged : End Sub
Private Sub GreenValue_Change()          : RGBTextChanged : End Sub
Private Sub BlueValue_Change()           : RGBTextChanged : End Sub
Private Sub HueValue_Change()            : HSVTextChanged : End Sub
Private Sub SaturationValue_Change()     : HSVTextChanged : End Sub
Private Sub BrightnessValue_Change()     : HSVTextChanged : End Sub

Private Sub RedScrollBar_Change()
   If Not Updating Then Call SetRGB(RedScrollBar.Value, GreenScrollBar.Value, BlueScrollBar.Value)
End Sub
Private Sub GreenScrollBar_Change()
   If Not Updating Then Call SetRGB(RedScrollBar.Value, GreenScrollBar.Value, BlueScrollBar.Value)
End Sub
Private Sub BlueScrollBar_Change()
   If Not Updating Then Call SetRGB(RedScrollBar.Value, GreenScrollBar.Value, BlueScrollBar.Value)
End Sub
Private Sub HueScrollBar_Change()
   If Not Updating Then ApplyHSV
End Sub
Private Sub SaturationScrollBar_Change()
   If Not Updating Then ApplyHSV
End Sub
Private Sub BrightnessScrollBar_Change()
   If Not Updating Then ApplyHSV
End Sub


Private Sub RGBTextChanged()
    If Updating Then Exit Sub

    If IsByte(RedValue.Value) And IsByte(GreenValue.Value) And IsByte(BlueValue.Value) Then
        Call SetRGB(CLng(RedValue.Value), CLng(GreenValue.Value), CLng(BlueValue.Value))
    End If
End Sub

Private Sub HSVTextChanged()
    If Updating Then Exit Sub

    If IsNumeric(HueValue.Value) And _
       IsNumeric(SaturationValue.Value) And _
       IsNumeric(BrightnessValue.Value) Then

        If HueValue.Value        >= 0 And HueValue.Value        <= 360 And _
           SaturationValue.Value >= 0 And SaturationValue.Value <= 100 And _
           BrightnessValue.Value >= 0 And BrightnessValue.Value <= 100 Then

            Updating = True
            HueScrollBar.Value        = CLng(HueValue.Value)
            SaturationScrollBar.Value = CLng(SaturationValue.Value)
            BrightnessScrollBar.Value = CLng(BrightnessValue.Value)
            Updating = False
            ApplyHSV
        End If
    End If
End Sub

Private Sub ApplyHSV()
    Dim R As Long, G As Long, B As Long

    Call HSVToRGB(CDbl(HueScrollBar.Value), _
                  CDbl(SaturationScrollBar.Value) / 100, _
                  CDbl(BrightnessScrollBar.Value) / 100, R, G, B)
    Call SetRGB(R, G, B)
End Sub


Private Sub ValueBox_Change()
    Dim R As Long, G As Long, B As Long
    Dim HexValue As String

    If Updating Then Exit Sub

    HexValue = Trim(ValueBox.Value)
    If TryParseHex(HexValue, R, G, B) Then
        Call SetRGB(R, G, B)
    End If
End Sub

Private Function TryParseHex(ByVal Text As String, ByRef R As Long, ByRef G As Long, ByRef B As Long) As Boolean
    On Error GoTo InvalidHex

    If Len(Text) <> 6 Then Exit Function
    With CreateObject("vbscript.regexp")
         .Global = True
         .Pattern = "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]"
         If Not .Match(Text) Then Exit Function
    End With

    R = CLng("&H" & Mid(Text, 1, 2))
    G = CLng("&H" & Mid(Text, 3, 2))
    B = CLng("&H" & Mid(Text, 5, 2))
    TryParseHeX = True
    Exit Function

InvalidHex:
    TryParseHeX = False
End Function

Private Function RGBToHex(ByVal R As Long, ByVal G As Long, ByVal B As Long) As String
    RGBToHeX = Empty & Right("0" & Hex(R), 2) _
                     & Right("0" & Hex(G), 2) _
                     & Right("0" & Hex(B), 2)
End Function


Private Sub SetRGB(ByVal R As Long, ByVal G As Long, ByVal B As Long)
    Dim H As Double, S As Double, V As Double

    R = Clamp(r, 0, 255)
    G = Clamp(g, 0, 255)
    B = Clamp(b, 0, 255)

    On Error GoTo CleanUp
    Updating = True

    RedScrollBar.Value   = R
    GreenScrollBar.Value = G
    BlueScrollBar.Value  = B
    RedValue.Value       = R
    GreenValue.Value     = G
    BlueValue.Value      = B

    ValueBox.Value       = RGBToHex(R, G, B)
    ColorLabel.BackColoR = RGB(R, G, B)

    Call RGBToHSV(R, G, B, H, S, V)
    HueScrollBar.Value        = Clamp(CLng(H), 0, 360)
    SaturationScrollBar.Value = Clamp(CLng(S * 100), 0, 100)
    BrightnessScrollBar.Value = Clamp(CLng(V * 100), 0, 100)

    HueValue.Value        = HueScrollBar.Value
    SaturationValue.Value = SaturationScrollBar.Value
    BrightnessValue.Value = BrightnessScrollBar.Value

CleanUp:
    Updating = False
End Sub


Private Sub RGBToHSV(ByVal R As Long  , ByVal G As Long  , ByVal B As Long, _
                     ByRef H As Double, ByRef S As Double, ByRef V As Double)

    Dim RD  As Double, GD  As Double, BD    As Double
    Dim MaX As Double, Min As Double, Delta As Double

    RD = R / 255
    GD = G / 255
    BD = B / 255

    MaX = RD
    If GD > Max Then MaX = GD
    If BD > Max Then MaX = BD

    Min = RD
    If GD < Min Then Min = GD
    If BD < Min Then Min = BD

    Delta = Max - Min
    V = Max

    If MaX = 0 Then
        S = 0
    Else
        S = Delta / Max
    End If

    Select Case True
      Case Delta = 0  : H = 0
      Case Max   = RD : H = 60 * ((GD - BD) / Delta)       : If h < 0 Then H = h + 360
      Case Max   = GD : H = 60 * (((BD - RD) / Delta) + 2)
      Case Else       : H = 60 * (((RD - GD) / Delta) + 4)
    End Select
End Sub

Private Sub HSVToRGB(ByVal H As Double, ByVal S As Double, ByVal V As Double, _
                     ByRef R As Long  , ByRef G As Long  , ByRef B As Long)
                     
    Dim C As Double, X As Double, M As Double
    Dim RD As Double, GD As Double, BD As Double
    Dim Sector As Long

    H = H Mod 360
    S = ClampDouble(S, 0, 1)
    V = ClampDouble(V, 0, 1)

    C = v * S
    X = C * (1 - Abs((H / 60) - 2 * Int(H / 120) - 1))
    M = V - C
    Sector = Int(H / 60)

    Select Case Sector
        Case 0   : RD = C: GD = X: BD = 0
        Case 1   : RD = X: GD = C: BD = 0
        Case 2   : RD = 0: GD = C: BD = X
        Case 3   : RD = 0: GD = X: BD = C
        Case 4   : RD = X: GD = 0: BD = C
        Case Else: RD = C: GD = 0: BD = X
    End Select

    R = Clamp(CLng((RD + m) * 255 + 0.5), 0, 255)
    G = Clamp(CLng((GD + m) * 255 + 0.5), 0, 255)
    B = Clamp(CLng((BD + m) * 255 + 0.5), 0, 255)
End Sub

Private Function IsByte(ByVal Value As Variant) As Boolean
    If Not IsNumeric(Value)                 Then Exit Function
    If CDbl(Value) < 0 Or CDbl(Value) > 255 Then Exit Function
    If CDbl(Value) <> Fix(CDbl(Value))      Then Exit Function
    IsByte = True
End Function

Private Function Clamp(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    Clamp = CLng(ClampDouble(Value, Min, Max))
End Function

Private Function ClampDouble(ByVal Value As Double, ByVal Min As Double, ByVal MaX As Double) As Double
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
    ClampDouble = Value
End Function