VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DateBox 
   Caption         =   "Select Date"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2865
   OleObjectBlob   =   "DateBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DateBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Implements IView

Private Type TView
    IsCancelled As Boolean
    Model As SingleInputModel
End Type

Private This As TView

Private FirstDayOfMonth As Date
Private FirstdayIndex As Long


Private Sub UserForm_Initialize()
    Dim i As Long
    Month.AddItem "Jan"
    Month.AddItem "Feb"
    Month.AddItem "Mar"
    Month.AddItem "Apr"
    Month.AddItem "May"
    Month.AddItem "Jun"
    Month.AddItem "Jul"
    Month.AddItem "Aug"
    Month.AddItem "Sep"
    Month.AddItem "Oct"
    Month.AddItem "Nov"
    Month.AddItem "Dec"
    Month.Value = Format(Date(), "mmm")
    For i = 0 To 9999
       Year.AddItem i
    Next i
    Year.Value = Format(Date(), "yyyy")
    Value.Text = FirstDayOfMonth
End Sub

Private Sub Month_Change()
    Dim CurrentDate As Date
    If Year.Value <> Empty Then
        CurrentDate = DateSerial(Year.Value, MonthIndex(Month.Value), 1)
        Call RefreshDates(CurrentDate)
    End If
End Sub

Private Sub Year_Change()
    Dim CurrentDate As Date
    If Month.Value <> Empty Then
        CurrentDate = DateSerial(Year.Value, MonthIndex(Month.Value), 1)
        Call RefreshDates(CurrentDate)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OkButton_Click()
    Hide
End Sub
 
Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Hide
End Sub

Private Sub CommandButton1_Click() : Call SetDate(01): End Sub
Private Sub CommandButton2_Click() : Call SetDate(02): End Sub
Private Sub CommandButton3_Click() : Call SetDate(03): End Sub
Private Sub CommandButton4_Click() : Call SetDate(04): End Sub
Private Sub CommandButton5_Click() : Call SetDate(05): End Sub
Private Sub CommandButton6_Click() : Call SetDate(06): End Sub
Private Sub CommandButton7_Click() : Call SetDate(07): End Sub
Private Sub CommandButton8_Click() : Call SetDate(08): End Sub
Private Sub CommandButton9_Click() : Call SetDate(09): End Sub
Private Sub CommandButton10_Click(): Call SetDate(10): End Sub
Private Sub CommandButton11_Click(): Call SetDate(11): End Sub
Private Sub CommandButton12_Click(): Call SetDate(12): End Sub
Private Sub CommandButton13_Click(): Call SetDate(13): End Sub
Private Sub CommandButton14_Click(): Call SetDate(14): End Sub
Private Sub CommandButton15_Click(): Call SetDate(15): End Sub
Private Sub CommandButton16_Click(): Call SetDate(16): End Sub
Private Sub CommandButton17_Click(): Call SetDate(17): End Sub
Private Sub CommandButton18_Click(): Call SetDate(18): End Sub
Private Sub CommandButton19_Click(): Call SetDate(19): End Sub
Private Sub CommandButton20_Click(): Call SetDate(20): End Sub
Private Sub CommandButton21_Click(): Call SetDate(21): End Sub
Private Sub CommandButton22_Click(): Call SetDate(22): End Sub
Private Sub CommandButton23_Click(): Call SetDate(23): End Sub
Private Sub CommandButton24_Click(): Call SetDate(24): End Sub
Private Sub CommandButton25_Click(): Call SetDate(25): End Sub
Private Sub CommandButton26_Click(): Call SetDate(26): End Sub
Private Sub CommandButton27_Click(): Call SetDate(27): End Sub
Private Sub CommandButton28_Click(): Call SetDate(28): End Sub
Private Sub CommandButton29_Click(): Call SetDate(29): End Sub
Private Sub CommandButton30_Click(): Call SetDate(30): End Sub
Private Sub CommandButton31_Click(): Call SetDate(31): End Sub
Private Sub CommandButton32_Click(): Call SetDate(32): End Sub
Private Sub CommandButton33_Click(): Call SetDate(33): End Sub
Private Sub CommandButton34_Click(): Call SetDate(34): End Sub
Private Sub CommandButton35_Click(): Call SetDate(35): End Sub

Private Sub RefreshDates(CurrentDate As Date)
    Dim CurrentMonth As Long
    Dim Firstday As Long
    Dim i As Long

    For i = 1 To 35
        Controls("CommandButton" & i).BackColor = RGB(255, 255, 255)
    Next i

    CurrentMonth = Format(CurrentDate, "m")
    FirstDayOfMonth = DateSerial(Year.Value, CurrentMonth, 1)
    Firstday = Weekday(FirstDayOfMonth)
    If Firstday = vbSunday Then
        FirstdayIndex = 7
    Else
        FirstdayIndex = Firstday - 1
    End If

    For i = 1 To 35
        Controls("CommandButton" & i).Caption = Format(FirstDayOfMonth + (i - FirstdayIndex), "dd")
        If Format(FirstDayOfMonth + (i - FirstdayIndex), "m") <> CurrentMonth Then
            Controls("CommandButton" & i).BackColor = RGB(200, 200, 200)
        End If
    Next i
End Sub

Private Function MonthIndex(Mon As String) As Long
    Select Case Mon
        Case "Jan": MonthIndex = 01
        Case "Feb": MonthIndex = 02
        Case "Mar": MonthIndex = 03
        Case "Apr": MonthIndex = 04
        Case "May": MonthIndex = 05
        Case "Jun": MonthIndex = 06
        Case "Jul": MonthIndex = 07
        Case "Aug": MonthIndex = 08
        Case "Sep": MonthIndex = 09
        Case "Oct": MonthIndex = 10
        Case "Nov": MonthIndex = 11
        Case "Dec": MonthIndex = 12
    End Select
End Function

Private Sub SetDate(ButtonIndex As Long)
    Value.Text = FirstDayOfMonth + (ButtonIndex - FirstdayIndex)
    This.Model.Value = Value.Text
End Sub

Public Function IView_Run(ByVal viewModel As Object, Optional LabelText As String = Empty) As Boolean
    Set This.Model = viewModel
    If LabelText <> Empty Then Me.Caption = LabelText
    Show
    IView_Run = Not This.IsCancelled
End Function