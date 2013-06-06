VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLookupValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private msLookupColName As String
Private mlParentHwnd As Long
Private mlColumnID As Long
Private mlChildHwnd As Long

Public Property Get LookupColName() As String

    LookupColName = msLookupColName

End Property

Public Property Let LookupColName(ByVal sLookupColName As String)

    msLookupColName = sLookupColName

End Property


Public Property Get ParentHwnd() As Long

    ParentHwnd = mlParentHwnd

End Property

Public Property Get ChildHwnd() As Long

    ChildHwnd = mlChildHwnd

End Property


Public Property Get ColumnID() As Long

    ColumnID = mlColumnID

End Property


Public Property Let ParentHwnd(ByVal lHwnd As Long)

    mlParentHwnd = lHwnd

End Property
Public Property Let ChildHwnd(ByVal lHwnd As Long)

    mlChildHwnd = lHwnd

End Property

Public Property Let ColumnID(ByVal plNewValue As Long)

    mlColumnID = plNewValue

End Property

