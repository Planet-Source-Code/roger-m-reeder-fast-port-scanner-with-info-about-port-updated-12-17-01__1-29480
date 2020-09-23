VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarPortNumber As Long 'local copy
Private mvarPossibleTrojans As String 'local copy
Private mvarPossibleRegister As String 'local copy
Public Property Let PossibleRegister(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PossibleRegister = 5
    mvarPossibleRegister = vData
End Property


Public Property Get PossibleRegister() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PossibleRegister
    PossibleRegister = mvarPossibleRegister
End Property



Public Property Let PossibleTrojans(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PossibleTrojans = 5
    mvarPossibleTrojans = vData
End Property


Public Property Get PossibleTrojans() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PossibleTrojans
    PossibleTrojans = mvarPossibleTrojans
End Property



Public Property Let PortNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PortNumber = 5
    mvarPortNumber = vData
End Property


Public Property Get PortNumber() As Long
Attribute PortNumber.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PortNumber
    PortNumber = mvarPortNumber
End Property


