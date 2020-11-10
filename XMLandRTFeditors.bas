Attribute VB_Name = "XMLandRTFeditors"
Option Explicit

Function FormatTED_TEP(mode As String, ByRef document As String, Optional CalcFileName As String, Optional CalcItemName As String, Optional NewVariablesXML As String) As String

Dim iFile As Integer

Dim wb As Workbook
Dim ws As Worksheet
Dim address As String
Dim jobno As String
Dim File As String

Set wb = ThisWorkbook
Set ws = Sheet3
address = ws.Range("E5").Value
jobno = ws.Range("Z2").Value

'Edit RTF depending on type of file
Select Case mode
    Case "TEP"
        document = Replace(document, "[[JOBNO]]", jobno)
        document = Replace(document, "[[ADDRESS]]", address)
        document = Replace(document, "[[DATE]]", Date)
        'SaveStringToTXT (document)
    Case "TED"
        document = Replace(document, "[[JOBNO]]", jobno)
        document = Replace(document, "[[ADDRESS]]", address)
        document = Replace(document, "[[DATE]]", Date)
        document = Replace(document, "[[CALCFILENAME]]", CalcFileName)
        document = Replace(document, "[[CALCITEMNAME]]", CalcItemName)
        document = Replace(document, "[[VARIABLES]]", NewVariablesXML)
End Select

FormatTED_TEP = document

End Function
Function ReadTEP_and_TED(FilePath As String) As String

Dim iFile As Integer

iFile = FreeFile
Open FilePath For Input As #iFile
ReadTEP_and_TED = Input(LOF(iFile), iFile)
Close #iFile

End Function

Function PrintTEP(document As String)

Dim iFile As Integer
Dim wb As Workbook
Dim ws As Worksheet
Dim address As String
Dim jobno As String

Set wb = ThisWorkbook
Set ws = Sheet3
address = ws.Range("E5").Value
jobno = ws.Range("Z2").Value

iFile = FreeFile
Open wb.Path & "\" & jobno & "-" & address & ".tep" For Output As #iFile
Print #iFile, document
Close #iFile

End Function

Function PrintTED(document As String, name As String) As String

Dim wb As Workbook
Dim iFile As Integer

Set wb = ThisWorkbook

iFile = FreeFile
Open wb.Path & "\" & name & ".ted" For Output As #iFile
Print #iFile, document
Close #iFile

End Function
Function ExtractVariables(variablesXML As String) As String
'called by section class
'uses tedds.calculator.getvariables function to extract only variables xml from full section variables xml
'to insert into existing calulations.

Dim Newstart As Integer
Dim Newend As Integer

Newstart = InStr(1, variablesXML, "<Variable>")
Newend = InStr(1, variablesXML, "</Variables>")
ExtractVariables = Mid(variablesXML, Newstart, Newend - Newstart)

End Function

'*************////////////     T F W    \\\\\\\\\\\***************


Function ReturnTeddsSectionXML(Section As Object, CalcItemName As String) As String
'Called by TFWExport; creates long string of variables in XML format for inserting into the word section. Works by combining preset varables in .txt with constructed variables in excel.
'reffer to clsTEDDS for more info

Dim NewVariablesXML As String
Dim PresetVariablesXML As String

'get excel section (NEW) variables
NewVariablesXML = Section.ConstructNewVariablesXML

'get preset variables
PresetVariablesXML = GetPresetVariablesXML(CalcItemName)

'combine excel section variable XML with Preset.txt
ReturnTeddsSectionXML = ConstructFullVariablesXML(NewVariablesXML, PresetVariablesXML)

'**FOR VIEWING OF XML VARIABLEs RESULTS**
SaveStringToTXT ReturnTeddsSectionXML

End Function

Function GetPresetVariablesXML(CalcItemName As String) As String

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim FileName As String
Dim content As String
Dim iFile As Integer

FileName = Get_sDirectory(PresetCalcs) & CalcItemName & ".txt"

If objFSO.FileExists(FileName) Then
    iFile = FreeFile
    Open FileName For Input As #iFile
    GetPresetVariablesXML = Input(LOF(iFile), iFile)
    Close #iFile
End If

Set objFSO = Nothing

End Function

Function ConstructFullVariablesXML(extraXML As String, OriginalXML As String) As String

    Dim InsertionPoint As Integer
    Dim OriginalVar As String
    Dim OriginalEnd As String
    Dim Newstart As Integer
    Dim Newend As Integer
    Dim NewVar As String
    
    Debug.Print extraXML
    'Debug.Print OriginalXML
    
    InsertionPoint = InStr(1, OriginalXML, "</Variables>")
    
    OriginalVar = Mid(OriginalXML, 1, InsertionPoint - 1)
    OriginalEnd = Mid(OriginalXML, InsertionPoint, Len(OriginalXML))
    
    Newstart = InStr(1, extraXML, "<Variable>")
    Newend = InStr(1, extraXML, "</Variables>")
    NewVar = Mid(extraXML, Newstart, Newend - Newstart)
    ConstructFullVariablesXML = OriginalVar & NewVar & OriginalEnd

End Function

Sub SaveStringToTXT(content As String)

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim FileName As String
Dim objTextfile As Object

FileName = Get_sDirectory(PresetCalcs) & "test.txt"

Set objTextfile = objFSO.CreateTextFile(FileName, True, True)
objTextfile.Write content
objTextfile.Close


End Sub
