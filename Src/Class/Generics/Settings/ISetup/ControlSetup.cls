VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "ControlSetup"

Option Explicit

Implements ISetup

Private mControl As Object

Public Function Create(ByVal TargetControl As Object) As ControlSetup
    Dim Result As New ControlSetup

    Set Result.mControl = TargetControl

    Set Create = Result
End Function

Private Sub ISetup_Apply(ByVal Settings As Settings)
    Dim Value As IGetValue

    On Error Resume Next
    If Settings.TryGetValue("Caption"                  , Value) Then mControl.Caption              = Value.Value
    If Settings.TryGetValue("Text"                     , Value) Then mControl.Text                 = Value.Value
    If Settings.TryGetValue("Value"                    , Value) Then mControl.Value                = Value.Value
    If Settings.TryGetValue("Visible"                  , Value) Then mControl.Visible              = Value.Value
    If Settings.TryGetValue("Enabled"                  , Value) Then mControl.Enabled              = Value.Value
    If Settings.TryGetValue("Locked"                   , Value) Then mControl.Locked               = Value.Value
    If Settings.TryGetValue("TabStop"                  , Value) Then mControl.TabStop              = Value.Value
    If Settings.TryGetValue("Left"                     , Value) Then mControl.Left                 = Value.Value
    If Settings.TryGetValue("Top"                      , Value) Then mControl.Top                  = Value.Value
    If Settings.TryGetValue("Width"                    , Value) Then mControl.Width                = Value.Value
    If Settings.TryGetValue("Height"                   , Value) Then mControl.Height               = Value.Value
    If Settings.TryGetValue("BackColor"                , Value) Then mControl.BackColor            = Value.Value
    If Settings.TryGetValue("ForeColor"                , Value) Then mControl.ForeColor            = Value.Value
    If Settings.TryGetValue("BorderColor"              , Value) Then mControl.BorderColor          = Value.Value
    If Settings.TryGetValue("BorderStyle"              , Value) Then mControl.BorderStyle          = Value.Value
    If Settings.TryGetValue("SpecialEffect"            , Value) Then mControl.SpecialEffect        = Value.Value
    If Settings.TryGetValue("FontName"                 , Value) Then mControl.Font.Name            = Value.Value
    If Settings.TryGetValue("FontSize"                 , Value) Then mControl.Font.Size            = Value.Value
    If Settings.TryGetValue("FontBold"                 , Value) Then mControl.Font.Bold            = Value.Value
    If Settings.TryGetValue("FontItalic"               , Value) Then mControl.Font.Italic          = Value.Value
    If Settings.TryGetValue("FontUnderline"            , Value) Then mControl.Font.Underline       = Value.Value
    If Settings.TryGetValue("TextAlign"                , Value) Then mControl.TextAlign            = Value.Value
    If Settings.TryGetValue("TextAlignHorizontal"      , Value) Then mControl.TextAlign            = Value.Value
    If Settings.TryGetValue("ListIndex"                , Value) Then mControl.ListIndex            = Value.Value
    If Settings.TryGetValue("ListRows"                 , Value) Then mControl.ListRows             = Value.Value
    If Settings.TryGetValue("BoundColumn"              , Value) Then mControl.BoundColumn          = Value.Value
    If Settings.TryGetValue("ColumnCount"              , Value) Then mControl.ColumnCount          = Value.Value
    If Settings.TryGetValue("ColumnWidths"             , Value) Then mControl.ColumnWidths         = Value.Value
    If Settings.TryGetValue("MatchRequired"            , Value) Then mControl.MatchRequired        = Value.Value
    If Settings.TryGetValue("Style"                    , Value) Then mControl.Style                = Value.Value
    If Settings.TryGetValue("MultiLine"                , Value) Then mControl.MultiLine            = Value.Value
    If Settings.TryGetValue("WordWrap"                 , Value) Then mControl.WordWrap             = Value.Value
    If Settings.TryGetValue("EnterKeyBehavior"         , Value) Then mControl.EnterKeyBehavior     = Value.Value
    If Settings.TryGetValue("ScrollBars"               , Value) Then mControl.ScrollBars           = Value.Value
    If Settings.TryGetValue("MaxLength"                , Value) Then mControl.MaxLength            = Value.Value
    If Settings.TryGetValue("PasswordChar"             , Value) Then mControl.PasswordChar         = Value.Value
    If Settings.TryGetValue("PictureSizeMode"          , Value) Then mControl.PictureSizeMode      = Value.Value
    If Settings.TryGetValue("PictureAlignment"         , Value) Then mControl.PictureAlignment     = Value.Value
    If Settings.TryGetValue("PictureTiling"            , Value) Then mControl.PictureTiling        = Value.Value
    If Settings.TryGetValue("ControlTipText"           , Value) Then mControl.ControlTipText       = Value.Value
    If Settings.TryGetValue("Tag"                      , Value) Then mControl.Tag                  = Value.Value
    If Settings.TryGetValue("TabIndex"                 , Value) Then mControl.TabIndex             = Value.Value

    On Error GoTo 0

End Sub