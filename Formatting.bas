Attribute VB_Name = "Formatting"
Sub FormatReportViaRibbon(control As IRibbonControl)

Call FormatReport

End Sub

Sub AddSummary(control As IRibbonControl)

Dim ws As Worksheet
Dim oSections As clsSections
Dim oSummSection As clsSection
Dim oSection As clsSection
Dim oRng As Range
Dim row As Integer
Dim index As Long

Set ws = Sheet3

'begin
Application.ScreenUpdating = False

'add template
Call Operator_member("memsummary")

If oSections Is Nothing Then Set oSections = factory.clsSections(ws)

If Not oSections Is Nothing Then
    Set oSummSection = oSections.GetSelectedSection(ws)
Else
    Exit Sub
End If

row = oSummSection.LoadRange.row + 1
    ws.Range("A" & row).EntireRow.Resize(oSections.SectionsCollection.Count - 1).Insert xlShiftDown


For index = oSections.SectionsCollection.Count To 1 Step -1
    
    Set oSection = oSections.Item(index)
    
    If oSection.MemName = "memsummary" Then GoTo nextsection
    
    ws.Cells(row, Get_lDocumentColumn(Title) + 1).Value = oSection.FullTitle
    row = row + 1

    
nextsection:
Next

'insert at top
oSections.GetRangeFromName("memsummary").EntireRow.Cut
oSections.Item(oSections.SectionsCollection.Count).Range.Insert
Call FormatReport


'ending
Application.ScreenUpdating = True

End Sub

Sub FormatReport()

Dim ws As Worksheet
Dim Sections As clsSections
Dim Section As clsSection
Dim Section2 As clsSection

Dim Page As Long
Dim header As Long
Dim a As Long
Dim b As Long
Dim index As Long
Dim Pagebott As Long

Set ws = Sheet3
Page = 36
header = 7

'begin
Application.ScreenUpdating = False

If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)

'check if there are any members
If Sections.Count = 0 Then
    ws.Range(Get_sDocumentColumn(left_margin) & header + 2, Get_sDocumentColumn(right_margin) & (header + Page)).Borders(xlInsideHorizontal).LineStyle = xlNone
    ws.Range(Get_sDocumentColumn(left_margin) & header + 2, Get_sDocumentColumn(right_margin) & (header + Page)).Borders(xlEdgeTop).LineStyle = xlNone
    ws.Range(Get_sDocumentColumn(left_margin) & (header + Page), Get_sDocumentColumn(right_margin) & (header + Page)).Borders(xlEdgeBottom).Weight = xlThin
    ws.PageSetup.PrintArea = ws.Range(Get_sDocumentColumn(left_margin) & 1, Get_sDocumentColumn(right_margin) & (header + Page)).address
    GoTo EndEnd
End If

'loop through sections from the bottom up and remove any areas where sections are not imediatly to each other.
For index = 1 To Sections.SectionsCollection.Count - 1 Step 1
    Set Section = Sections.Item(index)
    Set Section2 = Sections.Item(index + 1)
    Sect2Bott = Section2.Range.Rows.Count + Section2.Range.row
    If Section.Range.row <> Sect2Bott Then
        ws.Range(Get_sDocumentColumn(left_margin) & Sect2Bott).EntireRow.Resize(Section.Range.row - Sect2Bott).Delete
    End If
Next

'loop through collection in reverse from count to 1 in -1 step
For index = Sections.SectionsCollection.Count To 1 Step -1
    Set Section = Sections.Item(index)
    a = Round((Section.Range.row - (Page + header)) / Page + 1.474) ' page number of top row, accounting for 7rows of header and + 0.5 for rounding to work
    b = Round(((Section.Range.Rows.Count + Section.Range.row - 1) - (Page + header)) / Page + 1.5)
    If a <> b Then 'deal with moving the page if top row has different page number to bottom
        Pagebott = a * Page + header
        Section.Range.EntireRow.Resize(Pagebott - Section.Range.row + 2).Insert
        ws.Range(Get_sDocumentColumn(left_margin) & Pagebott, Get_sDocumentColumn(right_margin) & Pagebott).Borders(xlEdgeBottom).Weight = xlThin
    End If
Next

'finish off the bottom of document
Pagebott = Round((Sections.Item(1).Range.row - (Page + header)) / Page + 1.5) * Page + header
ws.PageSetup.PrintArea = ws.Range(Get_sDocumentColumn(left_margin) & 1, Get_sDocumentColumn(right_margin) & Pagebott).address
ws.Range(Get_sDocumentColumn(left_margin) & Sections.Item(1).Range.Rows.Count + Sections.Item(1).Range.row, Get_sDocumentColumn(right_margin) & Pagebott).Borders(xlInsideHorizontal).LineStyle = xlNone
ws.Range(Get_sDocumentColumn(left_margin) & Sections.Item(1).Range.Rows.Count + Sections.Item(1).Range.row, Get_sDocumentColumn(right_margin) & Pagebott).Borders(xlEdgeTop).LineStyle = xlNone
ws.Range(Get_sDocumentColumn(left_margin) & Pagebott, Get_sDocumentColumn(right_margin) & Pagebott).Borders(xlEdgeBottom).Weight = xlThin

EndEnd:
'ending
Application.ScreenUpdating = True
End Sub

