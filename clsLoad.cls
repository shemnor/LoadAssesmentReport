VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'properties
Private sNature As String
Private sScope As String
Private sType_ As String
Private sPrefixGen As String
Private sPrefixGenA As String
Private sPrefixGenB As String
Private sPrefixVDLA As String
Private sPrefixVDLB As String

Private lIndex As Long

Property Get Nature() As String
    Nature = sNature
End Property

Property Let Nature(sValue As String)
    sNature = sValue
End Property

Property Get Scope() As String
    Scope = sScope
End Property

Property Get index() As Long
    index = lIndex
End Property

Property Let index(lValue As Long)
    lIndex = lValue
End Property

Sub Class_Initialize()

sScope = "Beam"

End Sub
Sub GetLoadVariables_ForBeamDesign(VariablesColl As Collection, Worksheet As Worksheet, LoadRow As Long)
'scrapes information from worksheet to create obiects of clsVariable

'Prep
CreatePrefixes

'Applies to all Type_
CreateDefaultVariables VariablesColl, Worksheet, LoadRow

'Depending on Type_
CreateLoadVariables VariablesColl, Worksheet, LoadRow

End Sub

Private Sub CreatePrefixes()
'Creates Prefixes for the load object and stores them for later

    sPrefixGen = "{" & lIndex & "_" & sScope & "}"
    sPrefixGenA = "{a" & lIndex & "_" & sScope & "}"
    sPrefixGenB = "{b" & lIndex & "_" & sScope & "}"
    sPrefixVDLA = "{" & lIndex & "}v_{a_" & sScope & "}"
    sPrefixVDLB = "{" & lIndex & "}v_{b_" & sScope & "}"
    
End Sub

Private Sub CreateDefaultVariables(ByRef Variables As Collection, Worksheet As Worksheet, LoadRow As Long)
'Creates Load_Type_ (nature of load) and Type_ (type of load) variables for the load object, stores in collection object passed by ref

Dim Variable As clsVariable

'set the type variable
    sType_ = Worksheet.Range(Get_sVariableColumn("type_") & LoadRow).Value
    Set Variable = New clsVariable
    Variable.Add "Type_" & sPrefixGen, sType_
    Variables.Add Variable

'set the description variable
    Set Variable = New clsVariable
    Variable.Add "Load_Desc_" & sPrefixGen, Worksheet.Range(Get_sVariableColumn("load_description") & LoadRow).Value
    Variables.Add Variable

'set the nature variable
    If Nature = "Permanent" Then
        Set Variable = New clsVariable
        Variable.Add "Load_Type_" & sPrefixGen, 0
        Variables.Add Variable
    Else
        Set Variable = New clsVariable
        Variable.Add "Load_Type_" & sPrefixGen, 1
        Variables.Add Variable
    End If
    
End Sub

Private Sub CreateLoadVariables(ByRef Variables As Collection, Worksheet As Worksheet, LoadRow As Long)

Dim Variable As clsVariable

    Select Case sType_
    
        Case "Full UDL"
            Set Variable = New clsVariable
            Variable.Add "w_" & sPrefixGen, Round(Worksheet.Range(Get_sVariableColumn("load_valueA") & LoadRow).Value, 2), "kN/m"
            Variables.Add Variable
            
        Case "Point Load"
            Set Variable = New clsVariable
            Variable.Add "P_" & sPrefixGen, Round(Worksheet.Range(Get_sVariableColumn("load_valueA") & LoadRow).Value, 2), "kN"
            Variables.Add Variable

            Set Variable = New clsVariable
            Variable.Add "L_" & sPrefixGenA, Worksheet.Range(Get_sVariableColumn("load_PosA") & LoadRow).Value, "m"
            Variables.Add Variable
            
        Case "Partial UDL"
            Set Variable = New clsVariable
            Variable.Add "w_" & sPrefixGen, Round(Worksheet.Range(Get_sVariableColumn("load_valueA") & LoadRow).Value, 2), "kN/m"
            Variables.Add Variable

            Set Variable = New clsVariable
            Variable.Add "L_" & sPrefixGenA, Worksheet.Range(Get_sVariableColumn("load_PosA") & LoadRow).Value, "m"
            Variables.Add Variable

            Set Variable = New clsVariable
            Variable.Add "L_" & sPrefixGenB, Worksheet.Range(Get_sVariableColumn("load_PosB") & LoadRow).Value, "m"
            Variables.Add Variable
            
        Case "Partial VDL"
            Set Variable = New clsVariable
            Variable.Add "w_" & sPrefixVDLA, Round(Worksheet.Range(Get_sVariableColumn("load_valueA") & LoadRow).Value, 2), "kN/m"
            Variables.Add Variable
            
            Set Variable = New clsVariable
            Variable.Add "w_" & sPrefixVDLB, Round(Worksheet.Range(Get_sVariableColumn("load_valueB") & LoadRow).Value, 2), "kN/m"
            Variables.Add Variable
            
            Set Variable = New clsVariable
            Variable.Add "L_" & sPrefixGenA, Worksheet.Range(Get_sVariableColumn("load_PosA") & LoadRow).Value, "m"
            Variables.Add Variable

            Set Variable = New clsVariable
            Variable.Add "L_" & sPrefixGenB, Worksheet.Range(Get_sVariableColumn("load_PosB") & LoadRow).Value, "m"
            Variables.Add Variable
    
    End Select
End Sub
