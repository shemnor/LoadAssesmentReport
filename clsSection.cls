VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' CLASS MODULE CODE -- clsSection

Private lIndex As Long
Private lPropertyRow As Long

Private sMemName As String
Private sTitle As String
Private sFullTitle As String
Private sCalcFileName As String
Private sCalcItemName As String
Private sCalcSectionId As Long

Private oRange As Range
Private oLoadRange As Range
Private oMemberRange As Range

Private bCombinedMem As Boolean

Private Variables As Collection

Property Get index() As Long
    index = lIndex
End Property

Property Let index(lValue As Long)
    lIndex = lValue
End Property

Property Get PropertyRow() As Long
    PropertyRow = lPropertyRow
End Property

Property Let PropertyRow(lValue As Long)
    lPropertyRow = lValue
End Property

Property Get Title() As String
    Title = sTitle
End Property

Property Let Title(sValue As String)
    sTitle = sValue
End Property

Property Get FullTitle() As String
    FullTitle = sFullTitle
End Property

Property Let FullTitle(sValue As String)
    sFullTitle = sValue
End Property

Property Get CalcFileName() As String
    CalcFileName = sCalcFileName
End Property

Property Let CalcFileName(sValue As String)
    sCalcFileName = sValue
End Property

Property Get CalcItemName() As String
    CalcItemName = sCalcItemName
End Property

Property Let CalcItemName(sValue As String)
    sCalcItemName = sValue
End Property

Property Get MemName() As String
    MemName = sMemName
End Property

Property Let MemName(sValue As String)
    sMemName = sValue
End Property

Property Get Range() As Range
   Set Range = oRange
End Property

Property Set Range(oValue As Range)
    Set oRange = oValue
End Property
Property Get LoadRange() As Range
   Set LoadRange = oLoadRange
End Property

Property Set LoadRange(oValue As Range)
    Set oLoadRange = oValue
End Property
Property Get MemberRange() As Range
   Set MemberRange = oMemberRange
End Property

Property Set MemberRange(oValue As Range)
    Set oMemberRange = oValue
End Property

Property Get CombinedMem() As Boolean
    CombinedMem = bCombinedMem
End Property

Property Let CombinedMem(bValue As Boolean)
    bCombinedMem = bValue
End Property

Property Get VariableCollection() As Collection
'this returns Collection of variables stored for section object

    If Not Variables Is Nothing Then
        Set VariableCollection = Variables
    Else
        ConstructAllVariables Sheet3
        Set VariableCollection = Variables
    End If
End Property

Private Sub ConstructAllVariables(Worksheet As Worksheet)
'this runs subs which save variables directly to the collection property

'clear variables collection (thinking this might be retained from before?)
Set Variables = New Collection
'Section Variables
ConstructSectionVariables Worksheet:=Worksheet

'Load Variables
ConstructLoadVariables Worksheet:=Worksheet

End Sub

Private Sub ConstructSectionVariables(Worksheet As Worksheet)
'construct section variables specific to a type of tedds calculation

Dim Variable As clsVariable

Select Case CalcItemName

    Case "Timber beam analysis & design", "Steel beam analysis & design"
        Set Variable = New clsVariable
        Variable.Add _
            name:="L_{s1}", _
            Value:=Worksheet.Range(Get_sVariableColumn("length") & Me.PropertyRow).Value, _
            Unit:="mm"
        Variables.Add Variable
        
    Case "Robeslee lintel check"
        Set Variable = New clsVariable
        Variable.Add _
            name:="L", _
            Value:=Worksheet.Range(Get_sVariableColumn("length") & Me.PropertyRow).Value / 1000, _
            Unit:=""
        Variables.Add Variable
        
End Select

End Sub

Private Sub ConstructLoadVariables(Worksheet As Worksheet)
'construct section variables specific to a type of tedds calculation

Dim Variable As clsVariable

Select Case CalcItemName

    Case "Timber beam analysis & design", "Steel beam analysis & design"
        Call ConstructLoadVariablesForBeamDesign(Worksheet)
    
    Case "Robeslee lintel check"

        
End Select

End Sub

Private Sub ConstructLoadVariablesForBeamDesign(Worksheet As Worksheet)
'This is required for beam type design calcs where a number of specific variables for each load are required.
'Sub constructs a collection of variables for each load. Loops through each row and creates load object with assigned properties.


Dim Load As clsLoad
Dim LoadRow As Range
Dim Nature As String
Dim LoadCounter As Integer

'preindex loads to account for selfweight applied in TEDDS calcs
LoadCounter = 1

If LoadRange.Rows.Count < 3 Then Exit Sub

For Each LoadRow In LoadRange.Rows

    'check if not a load and use to determine nature
    If Worksheet.Range(Get_sVariableColumn("type_") & LoadRow.row).Value = "Permanent" Then
        Nature = "Permanent"
        GoTo NextLoad
    ElseIf Worksheet.Range(Get_sVariableColumn("type_") & LoadRow.row).Value = "Variable" Then
        Nature = "Variable"
        GoTo NextLoad
    End If
    
    'incremetnt load index and
    LoadCounter = LoadCounter + 1
    'Worksheet.Range(Get_sVariableColumn("load_index") & LoadRow.Row).Value = LoadCounter
    
    'create load variable
    If Load Is Nothing Then Set Load = New clsLoad
    
    'assign Properties
    Load.Nature = Nature
    Load.index = LoadCounter
    
    'Use clsLoad to return collection of clsVariable objects
    Load.GetLoadVariables_ForBeamDesign Variables, Worksheet, LoadRow.row
    
NextLoad:
Next LoadRow
End Sub

Function ConstructNewVariablesXML() As String

Dim calc As Object
Dim Variable As clsVariable

'create a calculator object
    Set calc = CreateObject("Tedds.Calculator")
    calc.Initialize
    calc.Functions.eval "DeleteVarsInAllSections()"

'loop through variable objects in section and add variables
    For Each Variable In VariableCollection
        If IsEmpty(Variable.Value) = False Then
            If Variable.name Like "Load_Type_*" Then
                calc.Functions.SetVar Variable.name, Variable.Value
            ElseIf IsNumeric(Variable.Value) = True Then
                calc.Functions.SetVar Variable.name, Variable.Value, Variable.Unit
            Else
                calc.Functions.SetVar Variable.name, Variable.Value
            End If
        End If
    Next Variable

'get section variables and extract only variables from xml
ConstructNewVariablesXML = ExtractVariables(calc.GetVariables)

End Function
