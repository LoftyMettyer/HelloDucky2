VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public MobileID As Long

' Edit a mobile designer page
Public Sub Edit()

  Dim frmSelection As frmSelection
  Dim lngID As Long
  Dim fExit As Boolean

  fExit = False
  Set frmSelection = New frmSelection
  
  Do While Not fExit
    Set frmSelection.MobileDesigner = Me
    frmSelection.Show vbModal

    lngID = frmSelection.SelectedID

    Select Case frmSelection.Action
  
      Case edtEdit
        EditMobileDesign (lngID)
      
      Case edtPrint
        'PrintFolderDefinition(lngID)
  
      Case edtSelect, edtCancel
        fExit = True
  
    End Select
  Loop

  Set frmSelection = Nothing

End Sub

Private Sub EditMobileDesign(ByRef lngID As Long)

'  dim objDesigner as SystemFramework


End Sub
