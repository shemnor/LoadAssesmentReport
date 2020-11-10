VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Worksheet As Worksheet
Private Sections As New Collection

Sub Initialize(Worksheet As Worksheet)
    
    Set Sections = New Collection
    
'scraping data
    ConstructSectionsRanges Worksheet:=Worksheet
    ConstructLoadsRanges Worksheet:=Worksheet
    ConstructMemberRanges Worksheet:=Worksheet

'manipulating data
    AssignSectionsProperties Worksheet:=Worksheet
    
'check if constructed
'    If Sections.Count = 0 Then
'
'    End If
    
End Sub

Sub Add(Section As clsSection)
'create a new person and add to collection
    Sections.Add Section
End Sub

Sub Remove(index As Variant)
'remove this person from collection
    Sections.Remove index
End Sub

Property Get Count() As Long
'return the number of people
    Count = Sections.Count
End Property

Property Get Item(index As Variant) As clsSection
'return this particular person
    Set Item = Sections(index)
End Property

Property Get SectionsCollection() As Collection
    Set SectionsCollection = Sections
End Property

Property Get GetSelectedSection(Worksheet As Worksheet) As clsSection
'returns section object which range property inersects with selected cell range

    Dim Section As clsSection

'do until selection intersects or Collection runs out of entries meaning no intersection found
    For Each Section In Sections
        If Not Intersect(Selection, Section.Range) Is Nothing Then
            Set GetSelectedSection = Section
            Exit For
        End If
    Next Section

    If Sections.Count = 0 Then Debug.Print ("Collection Sections is empty. Cannot evaluate selected Range")
    'If GetSelectedSectionRange Is Nothing Then Debug.Print ("Selected Range not in any section")

End Property

Property Get GetSectionFromRange(Worksheet As Worksheet, address As Range) As clsSection
'returns section object which range property inersects with provided cell range

    Dim Section As clsSection

'do until selection intersects or Collection runs out of entries meaning no intersection found
    For Each Section In Sections
        If Not Intersect(address, Section.Range) Is Nothing Then
            Set GetSectionFromRange = Section
            Exit For
        End If
    Next Section

    If Sections.Count = 0 Then Debug.Print ("Collection Sections is empty. Cannot evaluate selected Range")
    'If GetSelectedSectionRange Is Nothing Then Debug.Print ("Selected Range not in any section")

End Property

Property Get GetSelectedMemberRange(Worksheet As Worksheet) As Range
'returns member range which inersects with selected cell range

    Dim Section As clsSection

'do until selection intersects or Collection runs out of entries meaning no intersection found
    For Each Section In Sections
        If Not Intersect(Selection, Section.MemberRange) Is Nothing Then
            Set GetSelectedMemberRange = Section.MemberRange
            Exit For
        End If
    Next Section

    If Sections.Count = 0 Then Debug.Print ("Collection Sections is empty. Cannot evaluate selected member Range")
    'If GetSelectedSectionRange Is Nothing Then Debug.Print ("Selected Range not in any section")

End Property

Property Get GetRangeFromName(name As String) As Range
'returns range of member assigned with name

    Dim Section As clsSection
    
'do until secion with membername same as request
    For Each Section In Sections
        If Section.MemName = name Then
            Set GetRangeFromName = Section.MemberRange
            Exit For
        End If
    Next Section
    
    
End Property

Private Sub ConstructSectionsRanges(Worksheet As Worksheet)
'creates a collection of section objects and assigns .endrow and .startrow properties depending on positions of section and \section markers
    
    Dim Section As New clsSection 'initial section object
    Dim oRng As Range
    Dim column As String
    Dim StartRow As Long
    Dim EndRow As Long
    
    Dim leftmargin As Long
    Dim rightmargin As Long
    
    leftmargin = Get_lDocumentColumn(left_margin)
    rightmargin = Get_lDocumentColumn(right_margin)
    
'initialise
    column = Get_sDocumentColumn(section_marker)

'save current cell selection for eval
    Set oRng = Worksheet.Range(column & Rows.Count).End(xlUp)

'do until top margin or row 7 is reached
    Do Until oRng.Value = "calcstart" Or oRng.row < 7
        If oRng.Value = "\section" Then
            EndRow = oRng.row
            Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
        ElseIf oRng.Value = "section" Then
            StartRow = oRng.row
            Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
        ElseIf oRng.Value <> "\section" Or oRng.Value <> "section" Then
            Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
        End If
'if loop has added a startrow to object(last info req) then create new object
        If StartRow <> 0 Then
            Set Section.Range = Worksheet.Range(Worksheet.Cells(StartRow, leftmargin), Worksheet.Cells(EndRow, rightmargin))
            Sections.Add Section
            Set Section = New clsSection
            StartRow = 0
            EndRow = 0
        End If
    Loop

End Sub

Private Sub ConstructLoadsRanges(Worksheet As Worksheet)
'LOAD START AND END ROWS - loops through collection objects and assigns extra information to them - LOAD START AND END ROWS

    Dim oRng As Range
    Dim column As String
    Dim Section As clsSection 'initial section object
    Dim LoadStartRow As Long
    Dim LoadEndRow As Long
    
    Dim leftmargin As Long
    Dim rightmargin As Long
    
    leftmargin = Get_lDocumentColumn(left_margin)
    rightmargin = Get_lDocumentColumn(right_margin)

'initialise
    column = Get_sDocumentColumn(load_marker)

'save current cell selection for eval
    Set oRng = Worksheet.Range(column & Rows.Count).End(xlUp)

'For each object in collection
    For Each Section In Sections
        Do Until LoadEndRow <> 0 And LoadStartRow <> 0
            If oRng.Value = "\load" Then
                LoadEndRow = oRng.row
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            ElseIf oRng.Value = "load" Then
                LoadStartRow = oRng.row
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            Else
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            End If
        Loop
        
        Set Section.LoadRange = Worksheet.Range(Worksheet.Cells(LoadStartRow, leftmargin + 4), Worksheet.Cells(LoadEndRow - 1, rightmargin))
        LoadEndRow = 0
        LoadStartRow = 0
    Next Section

End Sub

Private Sub ConstructMemberRanges(Worksheet As Worksheet)
'LOAD START AND END ROWS - loops through collection objects and assigns extra information to them - member range

    Dim oRng As Range
    Dim omemRng As Range
    Dim column As String
    Dim Section As clsSection 'initial section object
    Dim Section2 As clsSection 'initial section object
    Dim StartRow As Long
    Dim EndRow As Long
    
    Dim leftmargin As Long
    Dim rightmargin As Long
    
    leftmargin = Get_lDocumentColumn(left_margin)
    rightmargin = Get_lDocumentColumn(right_hiddenmargin)

'initialise
    column = Get_sDocumentColumn(member_marker)

'save current cell selection for eval
    Set oRng = Worksheet.Range(column & Rows.Count).End(xlUp)

'For each object in collection
    Do Until oRng.Value = "calcstart" Or oRng.row < 7
        Do Until EndRow <> 0 And StartRow <> 0
            If oRng.Value = "\member" Then
                EndRow = oRng.row
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            ElseIf oRng.Value = "member" Then
                StartRow = oRng.row
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            Else
                Set oRng = Worksheet.Range(column & oRng.row).End(xlUp)
            End If
        Loop
        
        'If oRng.Row < 8 Then MsgBox ("could not compile member ranges. GetMemberRanges")
        
        Set omemRng = Worksheet.Range(Worksheet.Cells(StartRow, leftmargin), Worksheet.Cells(EndRow, rightmargin))
        EndRow = 0
        StartRow = 0
        
        For Each Section2 In Sections
            If Not Intersect(Section2.Range, omemRng) Is Nothing Then
                Set Section2.MemberRange = omemRng
                Section2.CombinedMem = True
            End If
        Next Section2
    Loop

End Sub

Private Sub AssignSectionsProperties(Worksheet As Worksheet)
'evaluate collection of section objects to establish property rows from range properties

    Dim Section As clsSection
    
    For Each Section In Sections
        
        'assign property row
        If Not Section.Range Is Nothing Then
            Section.PropertyRow = Section.Range.Offset(2, 0).row
        End If
        
        'assign Title
        Section.Title = Worksheet.Range(Get_sVariableColumn("section_title") & Section.PropertyRow).Value
        
        'assign FullTitle
        Section.FullTitle = Worksheet.Range(Get_sVariableColumn("section_FullTitle") & Section.PropertyRow).Value
                
        'assign CalcFileName
        Section.CalcFileName = Worksheet.Range(Get_sVariableColumn("CalcFileName") & Section.PropertyRow).Value
        
        'assign CalcItemName
        Section.CalcItemName = Worksheet.Range(Get_sVariableColumn("CalcItemName") & Section.PropertyRow).Value

        'assign CalcSectionId
        Section.CalcItemName = Worksheet.Range(Get_sVariableColumn("CalcItemName") & Section.PropertyRow).Value
        
        'assign MemName
        Section.MemName = Worksheet.Range(Get_sVariableColumn("mem_name") & Section.PropertyRow).Value
    
    Next Section

End Sub
