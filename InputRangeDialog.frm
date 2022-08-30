VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputRangeDialog 
   Caption         =   "Input range dialog"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "InputRangeDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputRangeDialog"
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

Private Sub Range_Button_Click()
    On Error Resume Next
    Dim x As Range
    Me.Hide
    Set x = Application.InputBox(PROMPT:="Cells with row numbers", TITLE:="Select a range of cells", Type:=8)
    If Err.Number = 0 Then
        InputString.Text = x.Address(True, True, xlA1, True, True)
    End If
    Me.Show
End Sub
