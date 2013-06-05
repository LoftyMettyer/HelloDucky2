VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Key As String
Attribute Key.VB_VarDescription = "Property Key (property name)"

'Local variables to hold property values
Private strName As String
Private varValue As Variant

Public Property Get Name() As String
Attribute Name.VB_Description = "Property Name"
  Name = strName
End Property

Public Property Let Name(NewName As String)
  strName = NewName
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Property Value"
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0
  Value = varValue
End Property

Public Property Let Value(NewValue As Variant)
  varValue = NewValue
End Property
