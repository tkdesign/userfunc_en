VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputDelimiterDialog 
   Caption         =   "Input delimiter dialog"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "InputDelimiterDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputDelimiterDialog"
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

Private Sub CR_Button_Click()
    InputString.Text = CStr(vbCrLf)
    DialogResult = 1
    Me.Hide
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

Private Sub Tab_Button_Click()
    InputString.Text = CStr(vbTab)
    DialogResult = 1
    Me.Hide
End Sub
