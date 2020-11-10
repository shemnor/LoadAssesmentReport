VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private UserForm As UserForm
Private cControl As control
Private EditableTextbox As New Collection

Sub Class_Initialize()
    Set UserForm = SectionProperty
End Sub

Property Get GetTBoxNames() As Collection
'not in use YET
    EditableTextbox.Add "section_name"
    EditableTextbox.Add "centers"
    EditableTextbox.Add "length"

    Set GetTBoxNames = EditableTextbox

End Property

Sub UpdateAllTBox(Worksheet As Worksheet, lPropertyRow As Long)
'loops through texboxes in userform and changes the box value to corresponding value in WS

    For Each cControl In UserForm.Controls
        If TypeName(cControl) = "TextBox" Then
            cControl.Value = Worksheet.Cells(lPropertyRow, Get_lVariableColumn(cControl.name)).Value
        End If
    Next

    UserForm.SectionPropertyRow.Caption = lPropertyRow
    UserForm.status.Caption = ""

End Sub

Sub UpdateSectionProperties(Worksheet As Worksheet, lPropertyRow As Long)
'loops through texboxes in userform and changes the corresponding value in WS to box value

    For Each cControl In UserForm.Controls
        If TypeName(cControl) = "TextBox" Then
            Worksheet.Cells(lPropertyRow, Get_lVariableColumn(cControl.name)).Value = cControl.Value
        End If
    Next

End Sub

Sub ClearUserForm()
'loops through all textboxes

    For Each cControl In UserForm.Controls
        If TypeName(cControl) = "TextBox" Then
            cControl.Value = ""
        End If
    Next

    UserForm.SectionPropertyRow.Caption = ""

End Sub


