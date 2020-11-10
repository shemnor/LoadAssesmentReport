Attribute VB_Name = "Tedds_and_TFW"
Option Explicit

Sub ExportToTFW(control As IRibbonControl)

Dim ws As Worksheet
Dim Sections As clsSections
Dim Section As clsSection
Dim TFW As clsTedds

Dim TeddsSection As TeddsCalcSection

Dim CustomVariablesXML As String
Dim index As Long

'control
If CheckToExport = False Then Exit Sub

Set ws = Sheet3

If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)

'Prepp a TFW document
If TFW Is Nothing Then Set TFW = New clsTedds

'format tedds doc with Jobno and project name
Call TFW.WordDocEdit(TFW.TeddsDoc)


For index = Sections.SectionsCollection.Count To 1 Step -1
    
    Set Section = Sections.Item(index)
    
    'check if its supported for export
    If Section.CalcItemName <> "Timber beam analysis & design" And _
    Section.CalcItemName <> "Steel beam analysis & design" And _
    Section.CalcItemName <> "Robeslee lintel check" _
    Then
        GoTo nextsection
    End If
    
    'Get XML variables string
    CustomVariablesXML = ReturnTeddsSectionXML(Section, Section.CalcItemName)

    'Create Section to run calculation
    Set TeddsSection = TFW.CreateCalcSection(Title:=Section.Title, CalcFileName:=Section.CalcFileName, CalcItemName:=Section.CalcItemName)
    'save section ID
    ActiveSheet.Range(Get_sVariableColumn("CalcSectionId") & Section.PropertyRow).Value = TeddsSection.ID

    'paste variables into section
    ClipBoardSetText CustomVariablesXML, RegisterClipboardFormat("TEDDS::VariableSection")
    TFW.TeddsDoc.PasteCalcSection 1, -1
    
nextsection:

Next

End Sub

Sub ExportToTedds_All(control As IRibbonControl)

Dim ws As Worksheet
Dim Sections As clsSections
Dim Section As clsSection
Dim clsTedds As clsTedds

Dim CustomVariablesXML As String
Dim index As Long
Dim TEP As String
Dim TED As String

'control
If CheckToExport = False Then Exit Sub
Application.StatusBar = "Working...."
Application.ScreenUpdating = False

Set ws = Sheet3


'Construct all sections
If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)

'check if project file already in folder before making a new one
If FileExists(ThisWorkbook.Path & "*.tep") = False Then

    'read Project file
    TEP = ReadTEP_and_TED(Get_sDirectory(PresetCalcs) & "Project.tep")
    
    'Format project file
    TEP = FormatTED_TEP("TEP", TEP)
    
    'save tep
    PrintTEP (TEP)

End If

'export each section
For index = Sections.SectionsCollection.Count To 1 Step -1

    Set Section = Sections.Item(index)

    If Section.MemName = "memsummary" Then GoTo nextsection

    'check if its supported for export
    If Section.CalcItemName <> "Timber beam analysis & design" And _
    Section.CalcItemName <> "Steel beam analysis & design" And _
    Section.CalcItemName <> "Robeslee lintel check" _
    Then
        GoTo nextsection
    End If
    
    If CheckTEDOverwrite(Section.FullTitle) = False Then GoTo nextsection
    
        'Get XML variables string
        CustomVariablesXML = Section.ConstructNewVariablesXML
    
        'read Project file
        TED = ReadTEP_and_TED(Get_sDirectory(PresetCalcs) & Section.CalcItemName & ".ted")
        
        'Format project file
        TED = FormatTED_TEP("TED", TED, Section.CalcFileName, Section.CalcItemName, CustomVariablesXML)
        
        'save tep
        PrintTED TED, Section.FullTitle

nextsection:

Next

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub

Sub ExportToTedds_SingleSection(control As IRibbonControl)

Dim ws As Worksheet
Dim Sections As clsSections
Dim Section As clsSection
'Dim clsTedds As clsTedds

Dim CustomVariablesXML As String
Dim TED As String

'control
If CheckToExport = False Then Exit Sub
Application.StatusBar = "Working...."
Application.ScreenUpdating = False

Set ws = Sheet3


'Construct all sections
If Sections Is Nothing Then Set Sections = factory.clsSections(Worksheet:=ws)

'Get selected section
Set Section = Sections.GetSelectedSection(ws)

'Validate section

If Section.MemName = "memsummary" Then GoTo EndExit

If Section.CalcItemName <> "Timber beam analysis & design" And _
Section.CalcItemName <> "Steel beam analysis & design" And _
Section.CalcItemName <> "Robeslee lintel check" _
Then
    GoTo EndExit
End If

'check if okay to overwrite
If CheckTEDOverwrite(Section.FullTitle) = False Then GoTo EndExit

'Export Section
'Get XML variables string
CustomVariablesXML = Section.ConstructNewVariablesXML

'read Project file
TED = ReadTEP_and_TED(Get_sDirectory(PresetCalcs) & Section.CalcItemName & ".ted")

'Format project file
TED = FormatTED_TEP("TED", TED, Section.CalcFileName, Section.CalcItemName, CustomVariablesXML)

'save tep
PrintTED TED, Section.FullTitle


EndExit:

Application.ScreenUpdating = True
Application.StatusBar = False

End Sub

Function CheckToExport() As Boolean

Dim MsgBoxAnswer As Long

MsgBoxAnswer = MsgBox("Would you like to export members to Tedds?", vbYesNo)

If MsgBoxAnswer = 6 Then
    CheckToExport = True
End If

End Function

Function CheckTEDOverwrite(TEDname) As Boolean

Dim MsgBoxAnswer As Long

If FileExists(ThisWorkbook.Path & "\" & TEDname & ".ted") = True Then
    
    MsgBoxAnswer = MsgBox("A file called " & TEDname & " already exists; Would you like to overwrite?", vbYesNo)
    If MsgBoxAnswer = 6 Then
        CheckTEDOverwrite = True
    End If

Else
    CheckTEDOverwrite = True
End If

End Function
Private Function FileExists(fname) As Boolean
' Returns TRUE if the file exists,
' fname path should evaluate directly to the folder being checked for_
' ex Desktop\test to check for if desktop contains test. If name of folder in fname path is not nothing_
' then folder exists

    Dim x As String
    x = Dir(fname, vbDirectory)
    If x <> "" Then FileExists = True Else FileExists = False 'if name of folder in fname path is no nothing then folder exists

End Function

