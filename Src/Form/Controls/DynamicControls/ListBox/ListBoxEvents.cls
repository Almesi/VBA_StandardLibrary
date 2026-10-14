VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ListBoxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True

Attribute VB_Exposed = False

Option Explicit

Public  WithEvents Control   As MSForms.ListBox
Private            Functions As FunctionDir

Public Function Create(ByVal nControl As MSForms.ListBox, ByVal nFunctions As FunctionDir) As ListBoxEvents
    Set Create = New ListBoxEvents
    Call Create.Initialize(nControl, nFunctions)
End Function

Friend Sub Initialize(ByVal nControl As MSForms.ListBox, ByVal nFunctions As FunctionDir)
    Set Control   = nControl
    Set Functions = nFunctions
End Sub

Private Sub Control_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer) : Call RunEvent(Functions, "BeforeDragOver", Cancel, Data, X, Y, DragState, Effect, Shift)                     : End Sub
Private Sub Control_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)    : Call RunEvent(Functions, "BeforeDropOrPaste", Cancel, Action, Data, X, Y, Effect, Shift)                     : End Sub
Private Sub Control_Change()                                                                                                                                                                                                                            : Call RunEvent(Functions, "Change")                                                                           : End Sub
Private Sub Control_Click()                                                                                                                                                                                                                             : Call RunEvent(Functions, "Click")                                                                            : End Sub
Private Sub Control_DblClick(ByVal Cancel As MSForms.ReturnBoolean)                                                                                                                                                                                     : Call RunEvent(Functions, "DblClick", Cancel)                                                                 : End Sub
Private Sub Control_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)           : Call RunEvent(Functions, "Error", Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)  : End Sub
Private Sub Control_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)                                                                                                                                                             : Call RunEvent(Functions, "KeyDown", KeyCode, Shift)                                                          : End Sub
Private Sub Control_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)                                                                                                                                                                                   : Call RunEvent(Functions, "KeyPress", KeyAscii)                                                               : End Sub
Private Sub Control_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)                                                                                                                                                               : Call RunEvent(Functions, "KeyUp", KeyCode, Shift)                                                            : End Sub
Private Sub Control_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)                                                                                                                                    : Call RunEvent(Functions, "MouseDown", Button, Shift, X, Y)                                                   : End Sub
Private Sub Control_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)                                                                                                                                    : Call RunEvent(Functions, "MouseMove", Button, Shift, X, Y)                                                   : End Sub
Private Sub Control_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)                                                                                                                                      : Call RunEvent(Functions, "MouseUp", Button, Shift, X, Y)                                                     : End Sub