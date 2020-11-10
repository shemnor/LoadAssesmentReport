VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sName As String
Private sValue As Variant
Private sUnit As String

Property Get name() As String
    name = sName
End Property

Property Let name(sValue As String)
    sName = sValue
End Property

Property Get Value() As Variant
    Value = sValue
End Property

Property Let Value(sVal As Variant)
    sValue = sVal
End Property

Property Get Unit() As String
    Unit = sUnit
End Property

Property Let Unit(sValue As String)
    sUnit = sValue
End Property

Sub Add(name As String, Value As Variant, Optional Unit As String)

Me.name = name
Me.Value = Value
Me.Unit = Unit

End Sub
