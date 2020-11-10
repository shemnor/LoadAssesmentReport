VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SectionProperty 
   Caption         =   "Member Properties"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "SectionProperty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SectionProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ComboBox1_Change()

End Sub

Private Sub Delete_Click()

Dim ws As Worksheet
Dim UserForm As clsUserForm
Dim MsgInp As Integer

MsgInp = MsgBox("Would you like to delete this member?", vbYesNo)

If MsgInp = vbNo Then Exit Sub

Set ws = Sheet3

Call DeleteMember(ws)

Call UpdateUserform(ws)

End Sub

Private Sub GO_Click()

Call AddComment

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub UPDATE_Click()

Dim ws As Worksheet
Dim UserForm As clsUserForm

Set ws = Sheet3

Call UpdateSectionProperties(ws)

Call UpdateUserform(ws)
Me.status.Caption = "UPDATED!"

End Sub
Private Sub UserForm_Initialize()

If Me.Visible = True Then Unload Me

    Me.StartUpPosition = 0
    Me.Top = Application.Top + Application.Height - Me.Height - 230
    Me.Left = Application.Left + Application.Width - Me.Width - 700

End Sub
