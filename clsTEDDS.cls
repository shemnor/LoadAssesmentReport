VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTEDDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private wordApp As Object
Private oTedds As teddsForWordAddIn
Private oTeddsDoc As teddsDocument
Private TFWAddIn As COMAddIn

Property Get TeddsDoc() As Object
    Set TeddsDoc = oTeddsDoc
End Property

Property Get Tedds() As Object
    Set Tedds = oTedds
End Property
'Property Get Tedds() As Object
'    Set Tedds = TFWAddIn.Object
'End Property

Sub Class_Initialize()
'Start Word and return Tedds for Word Add-In
    
On Error Resume Next

'Get Word
    Set wordApp = GetObject(, "Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = True
    End If
    If wordApp Is Nothing Then
        MsgBox ("Failed to start Microsoft Word, Unable to continue")
        End
    End If
    
'Get TFW add in
    Set TFWAddIn = wordApp.COMAddIns("Tedds for Word")
    If TFWAddIn Is Nothing Then
        MsgBox ("Failed to access Tedds for Word Add-In, Unable to continue")
        End
    End If

    If TFWAddIn.Connect = False Then
        TFWAddIn.Connect = True
    End If
    
    Set oTedds = TFWAddIn.Object
    
'Do some part of initialisig, not sure what is it needed for but batch design has it
    If Not (TFWAddIn Is Nothing) Then
        'if TFW loaded properly, open new tedds calc sheet
        oTedds.NewDefaultCalcSheet
    End If

'initialise document in active word
    Set oTeddsDoc = Tedds.InitDocument(wordApp.ActiveDocument)
    
'format document for the job


End Sub

Function CreateCalcSection(Title As String, CalcFileName As String, CalcItemName As String) As TeddsCalcSection

Dim newField As Object
Dim fieldCode As String

    'insert title
    Set CreateCalcSection = TeddsDoc.NewCalcSection(Title, -1)
    
    'insert calcfield
    wordApp.Selection.Style = wordApp.ActiveDocument.Styles("Calc 2 Column")
    fieldCode = "=CSC|CALL EvalCalcItem(""" & CalcFileName & """, """ & CalcItemName & """)"
    Set newField = wordApp.ActiveDocument.Fields.Add(wordApp.Selection.Range, -1, fieldCode, False)
    newField.result.text = CalcItemName
    newField.Select
    wordApp.Selection.Collapse 0
    wordApp.Selection.InsertParagraphBefore
    wordApp.Selection.MoveDown
    
    
End Function

Sub WordDocEdit(ByRef WordDoc As Object)

'formats information into word doc. Sub in word.normal.


Dim address As String
Dim jobno As Integer
Dim ws As Worksheet

Set ws = Sheet3
address = ws.Range("E5").Value
jobno = ws.Range("Z2").Value

'run doc formatting macro from word.normal
    WordDoc.WordDocument.Application.Run "StableMacros.FormatDocInfoCC", address, jobno, " "


End Sub
