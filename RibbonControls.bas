Attribute VB_Name = "RibbonControls"
Option Explicit

Sub Editor(control As IRibbonControl)

'With SectionProperty
'    If .Visible = True Then Unload SectionProperty
'    .Show vbModeless
'    .StartUpPosition = 0
'    .Top = Application.Top + Application.Height - .Height
'    .Left = Application.Left + Application.Width - .Width
'End With

SectionProperty.Show vbModeless

    Call UpdateUserform(Application.ActiveWorkbook.ActiveSheet)
    
End Sub

Sub UpdateUserform(ws As Worksheet)
'updates UF by scraping information from WS
'triggered by selection change (by checking if UF is showing)

Dim oSelectedSection As clsSection
Dim Sections As clsSections
Dim UserForm As clsUserForm
'Dim t1 As Single
't1 = Timer

If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)
If UserForm Is Nothing Then Set UserForm = New clsUserForm

'Get selected section object. If nothing then no section selected
    Set oSelectedSection = Sections.GetSelectedSection(Worksheet:=ws)

'If Section selected then update UF, else clear UF
    If Not oSelectedSection Is Nothing Then
        UserForm.UpdateAllTBox Worksheet:=ws, lPropertyRow:=oSelectedSection.PropertyRow
    Else
        UserForm.ClearUserForm
    End If

'Debug.Print Timer - t1

End Sub

Sub UpdateSectionProperties(ws As Worksheet)
'writes from UF to WB
'Triggered by UPDATE button Click

Dim oSelectedSection As clsSection
Dim Sections As clsSections
Dim UserForm As clsUserForm

If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)
If UserForm Is Nothing Then Set UserForm = New clsUserForm

'Get selected section object. If nothing then no section selected
    Set oSelectedSection = Sections.GetSelectedSection(Worksheet:=ws)

'If Section selected then update WS, else EXIT
    If Not oSelectedSection Is Nothing Then
        UserForm.UpdateSectionProperties Worksheet:=ws, lPropertyRow:=oSelectedSection.PropertyRow
    Else
        Debug.Print ("Selected cell is not in a recognised section.")
    End If

End Sub

Sub AddComment()
'adds comments box to section
'triggered by button on UF

Dim ws As Worksheet
Dim oSections As clsSections
Dim oSection As clsSection
Dim lRow As Long

Set ws = Sheet3

'begin
Application.ScreenUpdating = False

If oSections Is Nothing Then Set oSections = factory.clsSections(Worksheet:=ws)

With oSections.GetSelectedSection(ws).Range
    lRow = .row + .Rows.Count - 1 'last row
End With

'insert 3 rows taking 1 = format from below
ws.Range(Get_sDocumentColumn(left_margin) & lRow).EntireRow.Resize(3).Insert , 1

' make 1 space and insert 2 rows of comments, 3 columns form right margin, merge all
ws.Range((Get_sDocumentColumn(Title) & lRow + 1), Cells(lRow + 2, Get_lDocumentColumn(right_margin) - 3)).Merge

'formatting
With ws.Range(Get_sDocumentColumn(Title) & lRow + 1)
    .VerticalAlignment = xlTop
    .HorizontalAlignment = xlLeft
    .WrapText = True
End With

'ending
Application.ScreenUpdating = True

End Sub

Sub ShowHideSectionInfo(control As IRibbonControl)

Dim ws As Worksheet
Dim ColRange As Range

Set ws = Sheet3
Set ColRange = ws.Range(Cells(1, Get_lDocumentColumn(right_margin) + 1), Cells(1, Get_lDocumentColumn(right_hiddenmargin))).EntireColumn

If ColRange.Hidden = True Then
    ColRange.Hidden = False
Else
    ColRange.Hidden = True
End If

End Sub



