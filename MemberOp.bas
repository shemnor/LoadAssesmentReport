Attribute VB_Name = "MemberOp"
Option Explicit

Sub DeleteMember(ws As Worksheet)

Dim oSelectedMemberRange As Range
Dim Sections As clsSections

Application.ScreenUpdating = False

If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)

'Get selected section object. If nothing then no section selected
    Set oSelectedMemberRange = Sections.GetSelectedMemberRange(Worksheet:=ws)

'Delete Rows in member Range
    If oSelectedMemberRange Is Nothing Then
        MsgBox ("No member currently selected")
        Exit Sub
    Else
        oSelectedMemberRange.Delete Shift:=xlUp
    End If
    
Application.ScreenUpdating = True

End Sub

Sub Operator_member(name As String)

Dim DestinationWS As Worksheet
Dim SearchWS As Worksheet
Dim Sections As clsSections
Dim Seciton As clsSection

Dim FreeRow As Long
Dim CopyRNG As Range

Set DestinationWS = Sheet3
Set SearchWS = Sheet2

'begin
Application.ScreenUpdating = False
Application.EnableEvents = False
On Error GoTo ErrorEnd

If Sections Is Nothing Then Set Sections = factory.clsSections(SearchWS)

'set range to be copied
Set CopyRNG = Sections.GetRangeFromName(name)

'See if a whole row is selected meaning to make space above the row. Check if insertion point is inside a multimember
If Selection.Rows.Count = 1 And Selection.Columns.Count = DestinationWS.Columns.Count Then
    If DestinationWS.Range(Get_sDocumentColumn(member_marker) & ActiveCell.row).Value <> "member" Then
        MsgBox ("Cannot insert into combined members!")
        GoTo ErrorEnd
    End If
    ActiveCell.EntireRow.Resize(CopyRNG.Rows.Count, DestinationWS.Columns.Count).Insert
    FreeRow = ActiveCell.row
Else
'find first free row in destination ws
    FreeRow = DestinationWS.Range(Get_sDocumentColumn(section_marker) & Rows.Count).End(xlUp).row + 1
End If

'cpy and paste
CopyRNG.Copy
Application.ScreenUpdating = False
DestinationWS.Range(Get_sDocumentColumn(left_margin) & FreeRow).PasteSpecial

ActiveCell.Offset(1, 0).Select

'ending
ErrorEnd:
Application.CutCopyMode = False
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
Sub operator_loading(row As Integer)

Dim DestinationWS As Worksheet
Dim SearchWS As Worksheet

Dim PasteRow As Long
Dim CopyRNG As Range
Dim loadname As String

Set DestinationWS = Sheet3
Set SearchWS = Sheet2

'begin
Application.ScreenUpdating = False
Application.EnableEvents = False
On Error GoTo ErrorEnd

Set CopyRNG = SearchWS.Rows(row)
PasteRow = ActiveCell.row


'insert new above mode
If Selection.Rows.Count = 1 And Selection.Columns.Count = DestinationWS.Columns.Count Then
    ActiveCell.EntireRow.Insert
    CopyRNG.Copy
    DestinationWS.Range(Get_sDocumentColumn(left_margin) & PasteRow).PasteSpecial
    
'replace existing mode
ElseIf Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
'paste new load above
    ActiveCell.EntireRow.Insert
    CopyRNG.Copy
    DestinationWS.Range(Get_sDocumentColumn(left_margin) & PasteRow).PasteSpecial
'copy over old name and delete original
    loadname = DestinationWS.Range(Get_sVariableColumn("load_description") & PasteRow + 1).Value
    DestinationWS.Range(Get_sVariableColumn("load_description") & PasteRow).Value = loadname
    DestinationWS.Range("A" & PasteRow + 1).EntireRow.Delete
End If

'Select pasted row for next insert
DestinationWS.Range("A" & PasteRow + 1).EntireRow.Select

'ending
ErrorEnd:
Application.CutCopyMode = False
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
Sub AddLoadBoundary(mode As String)

'control As IRibbonControl, mode As String
Dim ws As Worksheet
Dim row As Integer
Dim oCopyRange As Range
Dim oPasteRange As Range

Set ws = Sheet3
row = Application.Selection.row
Set oCopyRange = ws.Range(Get_sDocumentColumn(boundaryA_left_margin) & row, Get_sDocumentColumn(boundaryA_right_margin) & row)
Set oPasteRange = ws.Range(Get_sDocumentColumn(boundaryB_left_margin) & row)

'begin
Application.ScreenUpdating = False
Application.EnableEvents = False
On Error GoTo ErrorEnd

'check if selection is valid
If ws.Range(Get_sVariableColumn("type_") & row).Value <> "Full UDL" Then
    If ws.Range(Get_sVariableColumn("type_") & row).Value <> "Partial UDL" Then
        If ws.Range(Get_sVariableColumn("type_") & row).Value <> "Partial VDL" Then
            MsgBox ("Load Not Selected")
            GoTo ErrorEnd
        End If
    End If
End If

'add locations
ws.Range(Get_sVariableColumn("load_PosA") & row).Value = 0
ws.Range(Get_sVariableColumn("load_PosA") & row).Offset(0, -1).Value = "@"

'paste B boundary
oCopyRange.Copy
oPasteRange.PasteSpecial

'Set B boundary load formula
With ws
    .Range(Get_sVariableColumn("load_valueB") & row).Formula = "=" & .Range(Get_sDocumentColumn(boundaryB_effarea) & row).address(False, False) & "*" & .Range(Get_sDocumentColumn(load_intensity) & row).address(False, False)
End With

'apply settings
Select Case mode
    Case "VUDL"
        ws.Range(Get_sVariableColumn("type_") & row).Value = "Partial VDL"
        ws.Range(Get_sDocumentColumn(load_type) & row).Value = "VDL (kN/m)"
        
    Case "PUDL"
        ws.Range(Get_sVariableColumn("type_") & row).Value = "Partial UDL"
        ws.Range(Get_sDocumentColumn(load_type) & row).Value = "PUDL (kN/m)"
End Select


'Select pasted row for next insert
ws.Range("A" & row + 1).EntireRow.Select


'ending
ErrorEnd:
Application.CutCopyMode = False
Application.ScreenUpdating = True
Application.EnableEvents = True


End Sub
Sub CopyLoadsFromOtherMember(control As IRibbonControl)

Dim ws As Worksheet
Dim oSections As clsSections
Dim oOriginLoads As Range
Dim oDestinationLoads As Range
Dim oRng As Range

Set ws = Sheet3

'begin
Application.ScreenUpdating = False
Application.EnableEvents = False
On Error GoTo ErrorEnd

If oSections Is Nothing Then Set oSections = factory.clsSections(ws)

'Check if multiselection is active, Get destination(where to) and origin(from where) Load Ranges.
If Application.Selection.Areas.Count > 1 Then
    For Each oRng In Application.Selection.Areas
        If oRng.Rows.Count > 1 Then
            Set oOriginLoads = oSections.GetSectionFromRange(Worksheet:=ws, address:=oRng).LoadRange
        Else
            Set oDestinationLoads = oSections.GetSectionFromRange(Worksheet:=ws, address:=oRng).LoadRange
        End If
    Next oRng
Else
    MsgBox ("Multiselection not detected. Please select origin loads and destination member")
    End
End If

'check if origin and destination were established
If oDestinationLoads Is Nothing Or oOriginLoads Is Nothing Then
    MsgBox ("Multiselection not detected. Please select origin loads and destination member")
    End
End If

'Delete Destination Ranges
oDestinationLoads.Resize(oDestinationLoads.Rows.Count - 1, 1).Offset(1, 0).EntireRow.Delete

'insert Origin range
oOriginLoads.Resize(oOriginLoads.Rows.Count - 1, 1).Offset(1, 0).EntireRow.Copy
oDestinationLoads.Offset(1, 0).EntireRow.Insert xlShiftUp

oDestinationLoads.Resize(1, 1).Select

'ending
ErrorEnd:
Application.CutCopyMode = False
Application.ScreenUpdating = True
Application.EnableEvents = True


End Sub

'###### LOADINGS ######

Sub UDLSYSTEM(control As IRibbonControl)

Call operator_loading(5)

End Sub
Sub UDLAREA(control As IRibbonControl)

Call operator_loading(6)

End Sub
Sub PUDL(control As IRibbonControl)

Call AddLoadBoundary("PUDL")

End Sub
Sub VUDL(control As IRibbonControl)

Call AddLoadBoundary("VUDL")

End Sub
Sub UDLWALL(control As IRibbonControl)

Call operator_loading(7)

End Sub
Sub CUDL(control As IRibbonControl)

Call operator_loading(8)

End Sub
Sub PLD(control As IRibbonControl)

Call operator_loading(9)

End Sub
Sub PLL(control As IRibbonControl)

Call operator_loading(10)

End Sub
Sub UDLC(control As IRibbonControl)

Call operator_loading(11)

End Sub

'###### LOADINGS ######

'###### MEMBERS ######

Sub Insert_rafter(control As IRibbonControl)

Call Operator_member("rafter")

End Sub
Sub Insert_joist(control As IRibbonControl)

Call Operator_member("joist")

End Sub
Sub Insert_steel_beam(control As IRibbonControl)

Call Operator_member("steel beam")

End Sub
Sub Insert_timber_beam(control As IRibbonControl)

Call Operator_member("timber beam")

End Sub
Sub Insert_concrete_lintel(control As IRibbonControl)

Call Operator_member("concrete lintel")

End Sub
Sub Insert_rooflight(control As IRibbonControl)

Call Operator_member("rooflight")
End Sub
Sub Insert_strip(control As IRibbonControl)

Call Operator_member("strip")

End Sub
Sub Insert_timber_truss(control As IRibbonControl)

Call Operator_member("truss")

End Sub
Sub Insert_dormer(control As IRibbonControl)

Call Operator_member("dormer")

End Sub
Sub Insert_racking_panel(control As IRibbonControl)

Call Operator_member("racking panel")

End Sub
Sub Insert_wall_opening(control As IRibbonControl)

Call Operator_member("TWall opening")

End Sub

'###### MEMBERS ######

'###### SUMMARIES ######

Sub WindSummary(control As IRibbonControl)

Call Operator_member("windsummary")

End Sub

'###### SUMMARIES ######
