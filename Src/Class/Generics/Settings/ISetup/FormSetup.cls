VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "FormSetup"

Option Explicit

Implements ISetup

Private mForm As Object

Public Function Create(ByVal TargetForm As Object) As FormSetup
    Dim Result As New FormSetup
    Set Result.mForm = TargetForm
    Set Create = Result
End Function

Private Sub ISetup_Apply(ByVal Settings As Settings)
    Dim Value As IGetValue
    If Settings.TryGetValue("Caption"               , Value) Then mForm.Caption               = Value.Value
    If Settings.TryGetValue("Tag"                   , Value) Then mForm.Tag                   = Value.Value
    If Settings.TryGetValue("Width"                 , Value) Then mForm.Width                 = Value.Value
    If Settings.TryGetValue("Height"                , Value) Then mForm.Height                = Value.Value
    If Settings.TryGetValue("Left"                  , Value) Then mForm.Left                  = Value.Value
    If Settings.TryGetValue("Top"                   , Value) Then mForm.Top                   = Value.Value
    If Settings.TryGetValue("BackColor"             , Value) Then mForm.BackColor             = Value.Value
    If Settings.TryGetValue("BorderColor"           , Value) Then mForm.BorderColor           = Value.Value
    If Settings.TryGetValue("BorderStyle"           , Value) Then mForm.BorderStyle           = Value.Value
    If Settings.TryGetValue("DrawBuffer"            , Value) Then mForm.DrawBuffer            = Value.Value
    If Settings.TryGetValue("PictureAlignment"      , Value) Then mForm.PictureAlignment      = Value.Value
    If Settings.TryGetValue("PictureSizeMode"       , Value) Then mForm.PictureSizeMode       = Value.Value
    If Settings.TryGetValue("PictureTiling"         , Value) Then mForm.PictureTiling         = Value.Value
    If Settings.TryGetValue("SpecialEffect"         , Value) Then mForm.SpecialEffect         = Value.Value
    If Settings.TryGetValue("FontName"              , Value) Then mForm.Font.Name             = Value.Value
    If Settings.TryGetValue("FontSize"              , Value) Then mForm.Font.Size             = Value.Value
    If Settings.TryGetValue("FontBold"              , Value) Then mForm.Font.Bold             = Value.Value
    If Settings.TryGetValue("FontItalic"            , Value) Then mForm.Font.Italic           = Value.Value
    If Settings.TryGetValue("FontUnderline"         , Value) Then mForm.Font.Underline        = Value.Value
    If Settings.TryGetValue("Enabled"               , Value) Then mForm.Enabled               = Value.Value
    If Settings.TryGetValue("Visible"               , Value) Then mForm.Visible               = Value.Value
    If Settings.TryGetValue("ShowModal"             , Value) Then mForm.ShowModal             = Value.Value
    If Settings.TryGetValue("KeepScrollBarsVisible" , Value) Then mForm.KeepScrollBarsVisible = Value.Value
    If Settings.TryGetValue("ScrollBars"            , Value) Then mForm.ScrollBars            = Value.Value
    If Settings.TryGetValue("ScrollTop"             , Value) Then mForm.ScrollTop             = Value.Value
    If Settings.TryGetValue("ScrollLeft"            , Value) Then mForm.ScrollLeft            = Value.Value
    If Settings.TryGetValue("StartUpPosition"       , Value) Then mForm.StartUpPosition       = Value.Value
    If Settings.TryGetValue("Zoom"                  , Value) Then mForm.Zoom                  = Value.Value
End Sub