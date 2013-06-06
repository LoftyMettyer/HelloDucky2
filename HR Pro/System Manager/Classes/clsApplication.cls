VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Local variables to hold property values
Private blnChanged As Boolean
Private blnChangedDiaryLink As Boolean
Private blnChangedEmailLink As Boolean
Private blnChangedOutlookLink As Boolean
Private blnChangedWorkflowLink As Boolean
Private blnChangedTableName As Boolean
Private blnChangedViewName As Boolean
Private blnChangedColumnName As Boolean
Private mblnChangedOvernightJobSchedule As Boolean
Private mblnChangedDataStructure As Boolean
Private mintAccessMode As AccessModes

Private blnLoggedIn As Boolean

Public Property Get Changed() As Boolean
  Changed = blnChanged And (mintAccessMode <> accSystemReadOnly)
End Property

Public Property Let Changed(ByVal IsChanged As Boolean)
  blnChanged = IsChanged
End Property

Public Property Get LoggedIn() As Boolean
  LoggedIn = blnLoggedIn
End Property

Public Property Let LoggedIn(ByVal IsLoggedIn As Boolean)
  blnLoggedIn = IsLoggedIn
End Property

Public Property Get Name() As String
  Name = App.ProductName
End Property

Private Sub Class_Initialize()
  mintAccessMode = accFull
End Sub

Public Property Get SelfServiceIntranetModule() As Boolean
  SelfServiceIntranetModule = gfSSIntranetModule
End Property

Public Property Get MobileModule() As Boolean
  MobileModule = gfMobileModule
End Property

Public Property Get WorkflowModule() As Boolean
  WorkflowModule = gfWorkflowModule
End Property

Public Property Get Version1Module() As Boolean
  Version1Module = gfVersion1Module
End Property

Public Property Get TrainingBookingModule() As Boolean
  TrainingBookingModule = gfTrainingBookingModule
End Property

Public Property Get PersonnelModule() As Boolean
  PersonnelModule = gfPersonnelModule
End Property

Public Property Get AbsenceModule() As Boolean
  AbsenceModule = gfAbsenceModule
End Property

Public Property Get ChangedDiaryLink() As Boolean
  ChangedDiaryLink = blnChangedDiaryLink
End Property

Public Property Let ChangedDiaryLink(ByVal blnNewValue As Boolean)
  blnChangedDiaryLink = blnNewValue
End Property

Public Property Get ChangedEmailLink() As Boolean
  ChangedEmailLink = blnChangedEmailLink
End Property

Public Property Get ChangedWorkflowLink() As Boolean
  ChangedWorkflowLink = blnChangedWorkflowLink
End Property

Public Property Let ChangedEmailLink(ByVal blnNewValue As Boolean)
  blnChangedEmailLink = blnNewValue
End Property

Public Property Let ChangedWorkflowLink(ByVal blnNewValue As Boolean)
  blnChangedWorkflowLink = blnNewValue
End Property

Public Property Get ChangedOutlookLink() As Boolean
  ChangedOutlookLink = blnChangedOutlookLink
End Property

Public Property Let ChangedOutlookLink(ByVal blnNewValue As Boolean)
  blnChangedOutlookLink = blnNewValue
End Property

Public Property Get ChangedTableName() As Boolean
  ChangedTableName = blnChangedTableName
End Property

Public Property Let ChangedTableName(ByVal blnNewValue As Boolean)
  blnChangedTableName = blnNewValue
End Property

Public Property Get ChangedViewName() As Boolean
  ChangedViewName = blnChangedViewName
End Property

Public Property Let ChangedViewName(ByVal blnNewValue As Boolean)
  blnChangedViewName = blnNewValue
End Property

Public Property Get ChangedColumnName() As Boolean
  ChangedColumnName = blnChangedColumnName
End Property

Public Property Let ChangedColumnName(ByVal blnNewValue As Boolean)
  blnChangedColumnName = blnNewValue
End Property

Public Property Get ChangedOvernightJobSchedule() As Boolean
  ChangedOvernightJobSchedule = mblnChangedOvernightJobSchedule
End Property

Public Property Let ChangedOvernightJobSchedule(ByVal blnNewValue As Boolean)
  mblnChangedOvernightJobSchedule = blnNewValue
End Property

Public Property Get ChangedDataStructure() As Boolean
  ChangedDataStructure = mblnChangedDataStructure
End Property

Public Property Let ChangedDataStructure(ByVal blnNewValue As Boolean)
  mblnChangedDataStructure = blnNewValue
End Property

Public Property Get AccessMode() As AccessModes
  AccessMode = mintAccessMode
End Property

Public Property Let AccessMode(ByVal intNewValue As AccessModes)
  mintAccessMode = intNewValue
End Property



