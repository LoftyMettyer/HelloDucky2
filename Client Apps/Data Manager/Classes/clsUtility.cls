VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private fOK As Boolean
Private frmDef As Form
Private mlngDefID As Long
Private mblnChanged As Boolean
Private mlngTimeStamp As Long
Private mblnLoading As Boolean
Private mblnFromCopy As Boolean
Private mblnDefCreator As Boolean
Private mblnForceHidden As Boolean


Public Property Get OK() As Boolean
  OK = fOK
End Property
Public Property Let OK(ByVal vNewValue As Boolean)
  fOK = vNewValue
End Property


Public Property Get ID() As Long
  DefID = mlngDefID
End Property
Public Property Let ID(ByVal vNewValue As Long)
  mlngDefID = vNewValue
End Property


Public Property Get Changed() As Boolean
  Changed = mblnChanged
End Property
Public Property Let Changed(ByVal vNewValue As Boolean)
  mblnChanged = vNewValue
End Property


Public Property Get TimeStamp() As Long
  TimeStamp = mlngTimeStamp
End Property
Public Property Let TimeStamp(ByVal vNewValue As Long)
  mlngTimeStamp = vNewValue
End Property


Public Property Get Loading() As Boolean
  Loading = mblnLoading
End Property
Public Property Let Loading(ByVal vNewValue As Boolean)
  mblnLoading = vNewValue
End Property


Public Property Get FromCopy() As Boolean
  FromCopy = mblnFromCopy
End Property
Public Property Let FromCopy(ByVal vNewValue As Boolean)
  mblnFromCopy = vNewValue
End Property


Public Property Get Creator() As Boolean
  DefCreator = mblnDefCreator
End Property
Public Property Let Creator(ByVal vNewValue As Boolean)
  mblnDefCreator = vNewValue
End Property


Public Property Get ForceHidden() As Boolean
  ForceHidden = mblnForceHidden
End Property
Public Property Let ForceHidden(ByVal vNewValue As Boolean)
  mblnForceHidden = vNewValue
End Property


Public Property Get DefForm() As Form
  DefForm = frmDef
End Property

Public Property Let DefForm(ByVal vNewValue As Form)
  Set frmDef = vNewValue
End Property


Private Sub ControlsChangedClear()

  'Set all controls datachanged flags to false

  Dim ctl As Control
  For Each ctl In frmDef
    If TypeOf ctl Is TextBox Or _
       TypeOf ctl Is ASRSpinner Or _
       TypeOf ctl Is CheckBox Or _
       TypeOf ctl Is ComboBox Then
          If ctl.DataChanged Then
            ctl.DataChanged = False
          End If
    End If
  Next ctl

  mblnChanged = False

End Sub

Private Function ControlsChangedCheck() As Boolean

  'Have any of the controls datachanged flag been set ?

  Dim ctl As Control

  ControlsChangedCheck = mblnChanged

  If ControlsChangedCheck = True Then
    Exit Function
  End If

  For Each ctl In frmDef
    If TypeOf ctl Is TextBox Or _
       TypeOf ctl Is ASRSpinner Or _
       TypeOf ctl Is CheckBox Or _
       TypeOf ctl Is ComboBox Then
          If ctl.DataChanged Then
            ControlsChangedCheck = True
            Exit Function
          End If
    End If
  Next ctl

End Function


Public Sub RecordSelectionClick(blnPicklist As Boolean, blnFilter As Boolean)

  If mblnLoading Then
    Exit Sub
  End If

  With frmDef
    .cmdPicklist.Enabled = blnPicklist
    If blnPicklist = False Then
      .txtPicklist.Text = vbNullString
      .txtPicklist.Tag = vbNullString
    ElseIf txtPicklist.Text = vbNullString Then
      .txtPicklist.Text = "<None>"
    End If

    .cmdFilter.Enabled = blnFilter
    If blnFilter = False Then
      .txtFilter.Text = vbNullString
      .txtFilter.Tag = vbNullString
    ElseIf txtFilter.Text = vbNullString Then
      .txtFilter.Text = "<None>"
    End If

    .optReadWrite.Enabled = mblnDefinitionCreator
    .optReadOnly.Enabled = mblnDefinitionCreator
    .optHidden.Enabled = mblnDefinitionCreator

  End With

  mblnChanged = True

End Sub


Public Sub GetPicklist(frmDef As Form)

  Dim sSQL As String
  Dim lParent As Long
  Dim fExit As Boolean
  Dim frmPick As frmPicklists
  Dim blnEnabled As Boolean

  On Error GoTo LocalErr
  Screen.MousePointer = vbHourglass

  fExit = False

  With frmDefSel

    .TableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(frmDef.txtPicklist.Tag) > 0 Then
      .SelectedID = Val(frmDef.txtPicklist.Tag)
    End If

    'loop until a picklist has been selected or cancelled
    Do While Not fExit

      If .ShowList("PICKLISTS") Then
        .Show vbModal

        Select Case frmDefSel.Action
        Case edtAdd
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(True, False, cboBaseTable.ItemData(cboBaseTable.ListIndex)) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          End With

        Case edtEdit
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(False, frmDefSel.FromCopy, cboBaseTable.ItemData(cboBaseTable.ListIndex), frmDefSel.SelectedID) Then
              .Show vbModal
            End If
            If frmDefSel.FromCopy Then
              frmDefSel.SelectedID = .SelectedID
            End If
            Unload frmPick
            Set frmPick = Nothing
          End With

        Case edtSelect

          If frmDefSel.HiddenDef Then
            If Not mblnDefinitionCreator Then
              MsgBox "Unable to select this picklist as it is a hidden picklist and you are not the owner of this definition", vbExclamation
              Set frmDefSel = Nothing
              Exit Sub
            Else
              frmDef.optReadWrite.Enabled = False
              frmDef.optReadOnly.Enabled = False
              frmDef.optHidden.Enabled = False
              If frmDef.optHidden.Value <> True Then
                frmDef.optHidden.Value = True
                MsgBox "This definition will now be hidden as a hidden picklist has been selected", vbInformation
              End If
            End If
          Else
            frmDef.optReadWrite.Enabled = mblnDefinitionCreator
            frmDef.optReadOnly.Enabled = mblnDefinitionCreator
            frmDef.optHidden.Enabled = mblnDefinitionCreator
          End If


          frmDef.txtPicklist = frmDefSel.SelectedText
          frmDef.txtPicklist.Tag = frmDefSel.SelectedID
          frmDef.txtFilter = ""
          frmDef.txtFilter.Tag = ""
          frmDef.fExit = True

        Case 0
          fExit = True

        End Select
      End If

    Loop

  End With

  Set frmDefSel = Nothing

Exit Sub

LocalErr:
  MsgBox "Error selecting picklist", vbCritical

End Sub





Private Function RetreiveDefinition(strSQL As String) As Recordset

  Dim rsTemp As Recordset
  Dim blnReadOnly As Boolean
  Dim blnHiddenPicklistOrFilter As Boolean

  On Error GoTo LocalErr
  
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  If rsTemp.BOF And rsTemp.EOF Then
    MsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Cross Tab"
    fOK = False
    Exit Function
  End If

  
  blnHiddenPicklistOrFilter = False
  mblnRecordSelectionInvalid = False
  
  If rsTemp!PickListID > 0 Then
    optPicklist = True
    If IsNull(rsTemp!PicklistName) Then
      MsgBox "The picklist used in this definition has been deleted", vbExclamation + vbOKOnly, "Custom Reports"
      mblnRecordSelectionInvalid = True
      txtPicklist.Tag = 0
      txtPicklist.Text = "<None>"
    Else
      txtPicklist.Tag = rsTemp!PickListID
      txtPicklist.Text = rsTemp!PicklistName
      blnHiddenPicklistOrFilter = (rsTemp!PicklistAccess = "HD")
    End If
  ElseIf rsTemp!FilterID > 0 Then
    optFilter = True
    If IsNull(rsTemp!FilterName) Then
      mblnRecordSelectionInvalid = True
      MsgBox "The filter used in this definition has been deleted", vbExclamation + vbOKOnly, "Custom Reports"
      txtFilter.Tag = 0
      txtFilter.Text = "<None>"
    Else
      txtFilter.Tag = rsTemp!FilterID
      txtFilter.Text = rsTemp!FilterName
      blnHiddenPicklistOrFilter = (rsTemp!FilterAccess = "HD")
    End If
  Else
    optAllRecords = True
  End If

  
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = datGeneral.UserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!UserName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!UserName) = LCase$(datGeneral.UserName))
  End If

  blnReadOnly = Not datGeneral.SystemPermission("CROSSTABS", "EDIT")

  'Only the creator of a definition can change the
  'access regardless of definition access (but can't
  'change access if selected hidden picklist or filter)
  If Not mblnDefinitionCreator Or blnHiddenPicklistOrFilter Then
    optReadWrite.Enabled = False
    optReadOnly.Enabled = False
    optHidden.Enabled = False
  End If
  
  Select Case rsTemp!Access
  Case "RW"
    optReadWrite = True
  Case "RO"
    optReadOnly = True
    blnReadOnly = (blnReadOnly Or Not mblnDefinitionCreator)
  Case "HD"
    optHidden = True
  End Select

  If blnReadOnly Then
    ControlsDisableAll Me
  End If
  
  mlngTimeStamp = rsTemp!intTimeStamp

  rsTemp.Close
  Set rsTemp = Nothing

Exit Function

LocalErr:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASR DEVELOPMENT"
    Stop
  End If
  ErrorMsgbox "Error retrieving Cross Tab definition"

End Function


Private Function ValidateDefinition() As Boolean

  'Check that all required information has been completed before attempting to save
  
  Dim blnContinueSave As Boolean
  Dim blnSaveAsNew As Boolean
  Dim strName As String
  
  On Error GoTo LocalErr
  
  ValidateDefinition = False
  strName = Trim(txtName.Text)

  If Len(strName) = 0 Then
    SSTab1.Tab = 0
    MsgBox "No name entered.", vbExclamation
    txtName.SetFocus
    Exit Function
  End If
  
  If optFilter Then
    
    If Val(txtFilter.Tag) = 0 Then
      SSTab1.Tab = 0
      MsgBox "No Filter entered for the base table.", vbExclamation
      cmdFilter.SetFocus
      Exit Function
    
    ElseIf IsRecordSelectionValid = False Then
      SSTab1.Tab = 0
      MsgBox "The filter used in this definition has been deleted by another user", vbExclamation
      txtFilter = "<None>"
      txtFilter.Tag = 0
      cmdFilter.SetFocus
      Exit Function
    End If
  
  End If
    
  If optPicklist Then
    
    If Val(txtPicklist.Tag) = 0 Then
      SSTab1.Tab = 0
      MsgBox "No Picklist entered for the base table.", vbExclamation
      cmdPicklist.SetFocus
      Exit Function
    
    ElseIf IsRecordSelectionValid = False Then
      SSTab1.Tab = 0
      MsgBox "The picklist used in this definition has been deleted by another user", vbExclamation
      txtPicklist = "<None>"
      txtPicklist.Tag = 0
      cmdPicklist.SetFocus
      Exit Function
    End If
  
  End If
    

  'If Val(mskHorizontalRange(0)) > 0 Then
  If Val(mskHorizontalRange(0)) <> 0 Or Val(mskHorizontalRange(1)) <> 0 Then
    If Val(mskHorizontalRange(1)) <= Val(mskHorizontalRange(0)) Then
      SSTab1.Tab = 1
      MsgBox "Horizontal stop value must be greater than Horizontal start value", vbExclamation
      mskHorizontalRange(1).SetFocus
      Exit Function
    End If
    If Val(mskHorizontalRange(2)) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Horizontal increment must be greater than zero", vbExclamation
      mskHorizontalRange(2).SetFocus
      Exit Function
    End If
  End If
    
  'If Val(mskVerticalRange(0)) > 0 Then
  If Val(mskVerticalRange(0)) <> 0 Or Val(mskVerticalRange(1)) <> 0 Then
    If Val(mskVerticalRange(1)) <= Val(mskVerticalRange(0)) Then
      SSTab1.Tab = 1
      MsgBox "Vertical stop value must be greater than Vertical start value", vbExclamation
      mskVerticalRange(1).SetFocus
      Exit Function
    End If
    If Val(mskVerticalRange(2)) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Vertical increment must be greater than zero", vbExclamation
      mskVerticalRange(2).SetFocus
      Exit Function
    End If
  End If
    
  If Val(mskPageBreakRange(0)) > 0 Then
    If Val(mskPageBreakRange(1)) <= Val(mskPageBreakRange(0)) Then
      SSTab1.Tab = 1
      MsgBox "Page Break stop value must be greater than Page Break start value", vbExclamation
      mskPageBreakRange(1).SetFocus
      Exit Function
    End If
    If Val(mskPageBreakRange(2)) <= 0 Then
      SSTab1.Tab = 1
      MsgBox "Page Break increment must be greater than zero", vbExclamation
      mskPageBreakRange(2).SetFocus
      Exit Function
    End If
  End If
    
    
  If optOutput(1).Value Then
    If chkSave.Value And txtFileName = "" Then
      SSTab1.Tab = 2
      MsgBox "You must select a filename if you opt to save the document !", vbExclamation
      cmdFileName.SetFocus
      Exit Function
    End If
  End If
  
  
  'Check if this definition has been changed by another user
  Call UtilityDefAmended("ASRSysCrossTab", "CrossTabID", mlngCrossTabID, mlngTimeStamp, blnContinueSave, blnSaveAsNew)
  If blnContinueSave = False Then
    Exit Function
  ElseIf blnSaveAsNew Then
    txtUserName = datGeneral.UserName
    optReadWrite.Enabled = True
    optReadOnly.Enabled = True
    optHidden.Enabled = True
    mlngCrossTabID = 0
  End If


  If ValidateDefinitionUniqueName(strName) = False Then
    SSTab1.Tab = 0
    MsgBox "Duplicate Cross Tabs Name.", vbExclamation
    txtName.SetFocus
    Exit Function
  End If
  
  
  ValidateDefinition = True

Exit Function

LocalErr:
  ErrorMsgbox "Error validating Cross Tab definition"

End Function

