VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditSmartTableRow 
   Caption         =   "Edit smart table row in dialog mode"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690.001
   OleObjectBlob   =   "EditSmartTableRow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditSmartTableRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ButtonPressed As Integer

Private Sub Cancel_Button_Click()
    DialogResult = 0
    Unload Me
End Sub

Private Sub OK_Button_Click()
    DialogResult = 1
    Me.Hide
End Sub

Public Property Let DialogResult(ByVal ButtonCode As Integer)
    ButtonPressed = ButtonCode
End Property
 
Public Property Get DialogResult() As Integer
    DialogResult = ButtonPressed
End Property

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call EditFieldDialog
End Sub

Private Sub EditFieldDialog()
    Dim Value
    Value = ListBox1.List(ListBox1.ListIndex, 1)
    InputStringDialog.Caption = "Editing a value in a list"
    InputStringDialog.DialogDescription.Caption = "Enter value"
    InputStringDialog.InputString = Value
    InputStringDialog.InputString.SetFocus
    InputStringDialog.InputString.SelStart = 0
    InputStringDialog.InputString.SelLength = Len(InputStringDialog.InputString.Text)
    Dim Result As Variant
    InputStringDialog.Show 1
    Result = InputStringDialog.DialogResult
    If Result = 0 Then
        Unload InputStringDialog
        Exit Sub
    End If
    Dim Result2 As Variant
    Result2 = InputStringDialog.InputString.Text
    If IsEmpty(Result2) Or Result2 = "" Then
        Unload InputStringDialog
        Exit Sub
    End If
    ListBox1.List(ListBox1.ListIndex, 1) = Result2
End Sub


Private Sub ListBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim intShiftDown As Integer, intAltDown As Integer
    Dim intCtrlDown As Integer
    Dim fmShiftMask, fmCtrlMask, fmAltMask
    fmShiftMask = 1       'SHIFT was pressed
    fmCtrlMask = 2    'CTRL was pressed
    fmAltMask = 4    'ALT was pressed
    ' Use bit masks to determine which key was pressed
    intShiftDown = (Shift And fmShiftMask) > 0
    intAltDown = (Shift And fmAltMask) > 0
    intCtrlDown = (Shift And fmCtrlMask) > 0
    If KeyCode = vbKeyReturn And intCtrlDown = -1 Then
        Call EditFieldDialog
    Else
        Exit Sub
    End If
End Sub
