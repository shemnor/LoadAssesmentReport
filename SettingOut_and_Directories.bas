Attribute VB_Name = "SettingOut_and_Directories"
Option Explicit
Public Enum DocumentColumns
    left_margin
    right_margin
    right_hiddenmargin
    section_marker
    load_marker
    member_marker
    section_id
    load_type
    Title
    load_intensity
    boundaryA_effarea
    boundaryB_effarea
    boundaryA_left_margin
    boundaryA_right_margin
    boundaryB_left_margin
    boundaryB_right_margin
End Enum

Public Enum FilePaths
    PresetCalcs
    SDCTeddsLibraries
    SDCTeddsCalcSets
End Enum

'Public Enum PropertyColumns
'    section_title
'    centers
'End Enum
Function Get_sVariableColumn(Variable As String) As String

'case strings need to match Userform textbox names!!

Select Case Variable

'MISC
    Case "section_title"
        Get_sVariableColumn = "BC"
    Case "section_FullTitle"
        Get_sVariableColumn = "BD"
    Case "centers"
        Get_sVariableColumn = "AX"
    Case "length"
        Get_sVariableColumn = "AW"
    Case "type_"
        Get_sVariableColumn = "AV"
    Case "mem_name"
        Get_sVariableColumn = "BB"

'LOADS
    Case "load_valueA"
        Get_sVariableColumn = "AB"
    Case "load_valueB"
        Get_sVariableColumn = "AL"
    Case "load_PosA"
        Get_sVariableColumn = "AD"
    Case "load_PosB"
        Get_sVariableColumn = "AN"
    Case "load_description"
        Get_sVariableColumn = "I"

'TEDDS
    Case "CalcFileName"
        Get_sVariableColumn = "AY"
    Case "CalcItemName"
        Get_sVariableColumn = "AZ"
    Case "CalcSectionId"
        Get_sVariableColumn = "BA"
        
End Select

If Get_sVariableColumn = "" Then MsgBox ("Problem with Setting out")

End Function

Function Get_lVariableColumn(Variable As String) As Long

Get_lVariableColumn = Range(Get_sVariableColumn(Variable) & 1).column

End Function

Function Get_sDocumentColumn(Variable As DocumentColumns) As String

'case strings need to match documentcolumn ENUMS @ top!!

Select Case Variable
    Case DocumentColumns.left_margin 'left_margin
        Get_sDocumentColumn = "A"
    Case DocumentColumns.right_margin 'right_margin
        Get_sDocumentColumn = "AQ"
    Case DocumentColumns.right_hiddenmargin 'right_margin
        Get_sDocumentColumn = "BF"
    Case DocumentColumns.member_marker 'member_marker
        Get_sDocumentColumn = "AS"
    Case DocumentColumns.section_marker 'section_marker
        Get_sDocumentColumn = "AT"
    Case DocumentColumns.load_marker 'load_marker
        Get_sDocumentColumn = "AU"
    Case DocumentColumns.load_type 'load_type
        Get_sDocumentColumn = "E"
    Case DocumentColumns.Title 'Title
        Get_sDocumentColumn = "E"
    Case DocumentColumns.load_intensity 'load_intensity
        Get_sDocumentColumn = "T"
    Case DocumentColumns.boundaryA_effarea 'boundaryA_effarea
        Get_sDocumentColumn = "Z"
    Case DocumentColumns.boundaryB_effarea 'boundaryB_effarea
        Get_sDocumentColumn = "AJ"
    Case DocumentColumns.boundaryA_left_margin 'boundaryA_left_margin
        Get_sDocumentColumn = "X"
    Case DocumentColumns.boundaryA_right_margin 'boundaryA_right_margin
        Get_sDocumentColumn = "AE"
    Case DocumentColumns.boundaryB_left_margin 'boundaryB_left_margin
        Get_sDocumentColumn = "AH"
    Case DocumentColumns.boundaryB_right_margin 'boundaryB_right_margin
        Get_sDocumentColumn = "AO"

End Select

If Get_sDocumentColumn = "" Then
    MsgBox ("Get Document Column function is ded. Terminating")
    End
End If

End Function

Function Get_lDocumentColumn(Variable As DocumentColumns) As Long

Get_lDocumentColumn = Range(Get_sDocumentColumn(Variable) & 1).column

End Function

Function Get_sDirectory(name As FilePaths)

Select Case name
    Case FilePaths.PresetCalcs
        Get_sDirectory = "P:\Design&Reference\Calcs\SDC Tedds Calcs\Preset Calcs\"
    Case FilePaths.PresetCalcs
        Get_sDirectory = " "
    Case FilePaths.PresetCalcs
        Get_sDirectory = " "
        
End Select

End Function
