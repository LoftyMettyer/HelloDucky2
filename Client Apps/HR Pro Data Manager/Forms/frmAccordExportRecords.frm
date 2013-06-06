VERSION 5.00
Begin VB.Form frmAccordExportRecords 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administer Transfers"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1149
   Icon            =   "frmAccordExportRecords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOptions 
      Caption         =   "Maintenance : "
      Height          =   915
      Left            =   120
      TabIndex        =   13
      Top             =   4485
      Width           =   4905
      Begin VB.CommandButton cmdPurge 
         Caption         =   "P&urge"
         Height          =   400
         Left            =   3435
         TabIndex        =   7
         Top             =   285
         Width           =   1183
      End
      Begin VB.Label lblZap 
         Caption         =   "Purge Transfer Table "
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Run"
      Height          =   400
      Left            =   3555
      TabIndex        =   6
      Top             =   3780
      Width           =   1183
   End
   Begin VB.Frame fraTransfer 
      Caption         =   "Transfer Type : "
      Height          =   4290
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   4905
      Begin VB.CommandButton cmdFilter 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         TabIndex        =   21
         Top             =   3195
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.CommandButton cmdPicklist 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4320
         TabIndex        =   20
         Top             =   2820
         UseMaskColor    =   -1  'True
         Width           =   300
      End
      Begin VB.OptionButton optAllRecords 
         Caption         =   "&All"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1245
         TabIndex        =   19
         Top             =   2475
         Value           =   -1  'True
         Width           =   720
      End
      Begin VB.OptionButton optPicklist 
         Caption         =   "&Picklist"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1245
         TabIndex        =   18
         Top             =   2880
         Width           =   930
      End
      Begin VB.OptionButton optFilter 
         Caption         =   "&Filter"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1245
         TabIndex        =   17
         Top             =   3270
         Width           =   930
      End
      Begin VB.TextBox txtPicklist 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   16
         Tag             =   "0"
         Top             =   2820
         Width           =   2085
      End
      Begin VB.TextBox txtFilter 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2235
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "0"
         Top             =   3210
         Width           =   2085
      End
      Begin VB.CheckBox chkBypassFilter 
         Caption         =   "&Bypass system filter"
         Height          =   330
         Left            =   225
         TabIndex        =   3
         Top             =   2025
         Width           =   3390
      End
      Begin VB.Frame fraDefaults 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   360
         TabIndex        =   9
         Top             =   1035
         Width           =   4380
         Begin VB.ComboBox cboType 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmAccordExportRecords.frx":000C
            Left            =   1125
            List            =   "frmAccordExportRecords.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   45
            Width           =   3135
         End
         Begin VB.ComboBox cboStatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmAccordExportRecords.frx":0010
            Left            =   1125
            List            =   "frmAccordExportRecords.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   450
            Width           =   3135
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status : "
            Height          =   285
            Left            =   135
            TabIndex        =   11
            Top             =   495
            Width           =   825
         End
         Begin VB.Label lblType 
            Caption         =   "Type : "
            Height          =   285
            Left            =   135
            TabIndex        =   10
            Top             =   90
            Width           =   645
         End
      End
      Begin VB.CheckBox chkManualType 
         Caption         =   "Override &Defaults"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   765
         Width           =   2085
      End
      Begin VB.ComboBox cboTransfer 
         Height          =   315
         ItemData        =   "frmAccordExportRecords.frx":0014
         Left            =   1485
         List            =   "frmAccordExportRecords.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   305
         Width           =   3135
      End
      Begin VB.Label lblBaseRecords 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records :"
         Height          =   195
         Left            =   495
         TabIndex        =   22
         Top             =   2475
         Width           =   690
      End
      Begin VB.Label lblTransfer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transfer :"
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   405
      Left            =   3690
      TabIndex        =   0
      Top             =   5520
      Width           =   1320
   End
End
Attribute VB_Name = "frmAccordExportRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarUDFsRequired() As String
Private mlngTransferTypeID As Long
Private mbCancelled As Boolean
Private mlngIndividualRecordID As Long

Public Property Let Cancelled(bCancelled As Boolean)
  mbCancelled = bCancelled
End Property
Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let RecordID(lngRecordID As Long)
  mlngIndividualRecordID = lngRecordID
End Property

Private Sub cboTransfer_Change()
  optAllRecords.Value = True
End Sub

Private Sub cboTransfer_Click()

  Dim objTableView As CTablePrivilege
  Dim iTableID As Integer
  Dim bFound As Boolean
  
  iTableID = GetBaseTableID(GetComboItem(cboTransfer))
  bFound = False
  For Each objTableView In gcoTablePrivileges.Collection
    If (objTableView.TableID = iTableID) And _
      (objTableView.AllowUpdate) Then
      bFound = True
      Exit For
    End If
  Next objTableView
  Set objTableView = Nothing

  If Not bFound Then
    COAMsgBox "You need full table access to run the " & Trim(cboTransfer.Text) & " export.", vbExclamation, Me.Caption
  End If

  cmdOK.Enabled = bFound
  
  'optAllRecords.Value = True
  'optAllRecords.Enabled = bFound
  'lblBaseRecords.Enabled = bFound
  'optFilter.Enabled = bFound
  'optPicklist.Enabled = bFound
  
End Sub

Private Sub chkBypassFilter_Click()
  
  Dim bEnable As Boolean
  bEnable = (chkBypassFilter.Value = vbChecked)
  
  EnableControl lblBaseRecords, bEnable
  
  optAllRecords.Value = True
  optAllRecords.Enabled = bEnable
  optPicklist.Enabled = bEnable
  optFilter.Enabled = bEnable
    
  cmdFilter.Enabled = optFilter.Enabled And optFilter.Value = True
  cmdPicklist.Enabled = optPicklist.Enabled And optPicklist.Value = True
    
End Sub

Private Sub chkManualType_Click()
  ControlsDisableAll fraDefaults, (chkManualType.Value = vbChecked)
End Sub

Private Sub cmdFilter_Click()
  ' Allow the user to select/create/modify a filter for the Data Transfer.
  Dim objExpression As clsExprExpression

  On Error GoTo LocalErr
  
  ' Instantiate a new expression object.
  Set objExpression = New clsExprExpression

  With objExpression
    ' Initialise the expression object.
    If .Initialise(GetBaseTableID(GetComboItem(cboTransfer)), Val(txtFilter.Tag), giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
  
      ' Instruct the expression object to display the expression selection/creation/modification form.
      If .SelectExpression(True) Then

        ' Read the selected expression info.
        txtFilter.Text = .Name
        txtFilter.Tag = .ExpressionID
        txtPicklist.Text = ""
        txtPicklist.Tag = 0
      End If

    End If
  End With

  Set objExpression = Nothing
  
Exit Sub

LocalErr:
  COAMsgBox "Error selecting filter"

End Sub

Private Sub cmdPicklist_Click()
  
  Dim fExit As Boolean
  Dim frmPick As frmPicklists

  On Error GoTo LocalErr
  
  Screen.MousePointer = vbHourglass

  fExit = False
  
  With frmDefSel
    
    .TableID = GetBaseTableID(GetComboItem(cboTransfer))
    .TableComboVisible = True
    .TableComboEnabled = False
    If Val(txtPicklist.Tag) > 0 Then
      .SelectedID = Val(txtPicklist.Tag)
    End If

    'loop until a picklist has been selected or cancelled
    Do While Not fExit

      If .ShowList(utlPicklist) Then
        .Show vbModal

        Select Case frmDefSel.Action
        Case edtAdd
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(True, False, GetBaseTableID(GetComboItem(cboTransfer))) Then
              .Show vbModal
            End If
            frmDefSel.SelectedID = .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          End With

        Case edtEdit
          Set frmPick = New frmPicklists
          With frmPick
            If .InitialisePickList(False, frmDefSel.FromCopy, GetBaseTableID(GetComboItem(cboTransfer)), frmDefSel.SelectedID) Then
              .Show vbModal
            End If
            If frmDefSel.FromCopy And .SelectedID > 0 Then
              frmDefSel.SelectedID = .SelectedID
            End If
            Unload frmPick
            Set frmPick = Nothing
          End With

        'MH20050728 Fault 10232
        Case edtPrint
          Set frmPick = New frmPicklists
          frmPick.PrintDef .TableID, .SelectedID
          Unload frmPick
          Set frmPick = Nothing
        
        Case edtSelect

          txtPicklist = frmDefSel.SelectedText
          txtPicklist.Tag = frmDefSel.SelectedID
          txtFilter.Text = ""
          txtFilter.Tag = 0
          fExit = True

        Case 0
          fExit = True
        End Select
      End If

    Loop

  End With
  
  Set frmDefSel = Nothing
   
Exit Sub

LocalErr:
  COAMsgBox "Error selecting picklist"
End Sub

Private Sub cmdCancel_Click()
  mbCancelled = True
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  Dim iTransactionType As Integer
  
  ' Validate options
  If optPicklist.Value Then
    If txtPicklist.Text = "" Or txtPicklist.Tag = "0" Or txtPicklist.Tag = "" Then
      COAMsgBox "You must select a picklist, or change the record selection.", vbExclamation + vbOKOnly
      cmdPicklist.SetFocus
      Exit Sub
    End If
  End If
  
  If optFilter.Value Then
    If txtFilter.Text = "" Or txtFilter.Tag = "0" Or txtFilter.Tag = "" Then
      COAMsgBox "You must select a filter, or change the record selection.", vbExclamation + vbOKOnly
      cmdFilter.SetFocus
      Exit Sub
    End If
  End If
  
  iTransactionType = GetComboItem(cboType)
  SendAccordTransactions GetComboItem(cboTransfer), iTransactionType

End Sub

Public Property Let TransferTypeID(plngID As Long)
  mlngTransferTypeID = plngID
End Property

Public Function Initialise() As Boolean
  
  Initialise = True
  PopulateControls

  If cboTransfer.ListCount = 0 Then
    COAMsgBox "No transfer definitions defined. Please see your system administrator.", vbExclamation, Me.Caption
    Initialise = False
  End If

  ControlsDisableAll fraDefaults, False
  
  ' Ensure bypass filter enabling/disabling takes place
  chkBypassFilter.Value = vbChecked
  chkBypassFilter.Value = vbUnchecked

End Function

Private Sub PopulateControls()

  ' Transfer Types
  PopulateAccordTransferTypes cboTransfer, False

  ' Transaction Types
  With cboType
    
    .AddItem "Calculated"
    .ItemData(.NewIndex) = -1
    
    .AddItem "New"
    .ItemData(.NewIndex) = 0
    
    .AddItem "Update"
    .ItemData(.NewIndex) = 1
    
    .AddItem "Delete"
    .ItemData(.NewIndex) = 2
      
  End With
  
  ' Status
  With cboStatus
    .AddItem "Pending"
    .ItemData(.NewIndex) = 1
  
    .AddItem "Blocked"
    .ItemData(.NewIndex) = 30
  
    .AddItem "Success"
    .ItemData(.NewIndex) = 10
    
  End With

  ' Set the default values
  'NHRD13092006 Fault 11402
  SetComboItem cboStatus, Val(GetModuleParameter(gsMODULEKEY_ACCORD, gsPARAMETERKEY_DEFAULTSTATUS))
  SetComboItem cboType, -1


End Sub

Private Sub cmdPurge_Click()

On Error GoTo ErrorTrap

  Dim strSQL As String

  If COAMsgBox("Purging the transfer table will reset the statuses of all records." & vbNewLine & "THIS PROCESS CANNOT BE UNDONE." & vbNewLine & "Are you sure you want to proceed?", vbYesNo + vbQuestion) = vbYes Then

    strSQL = "DELETE FROM ASRSysAccordTransactions"
    gADOCon.Execute strSQL, , adExecuteNoRecords

  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  COAMsgBox Err.Description, vbCritical
  Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_Load()
  
  ReDim mvarUDFsRequired(0)

End Sub

Private Sub optAllRecords_Click()

  cmdPicklist.Enabled = False
  With txtPicklist
    .Text = ""
    .Tag = 0
  End With
  
  cmdFilter.Enabled = False
  With txtFilter
    .Text = ""
    .Tag = 0
  End With

End Sub

Private Sub optFilter_Click()

  cmdPicklist.Enabled = False
  With txtPicklist
    .Text = ""
    .Tag = 0
  End With

  cmdFilter.Enabled = True
  txtFilter.Text = "<None>"

End Sub

Private Sub optPicklist_Click()

  cmdFilter.Enabled = False
  With txtFilter
    .Text = ""
    .Tag = 0
  End With

  cmdPicklist.Enabled = True
  txtPicklist.Text = "<None>"

End Sub

Private Function GetComboItem(cboTemp As ComboBox) As Long
  GetComboItem = 0
  If cboTemp.ListIndex <> -1 Then
    GetComboItem = cboTemp.ItemData(cboTemp.ListIndex)
  End If
End Function

Private Function GetBaseTableID(ByVal iTransferTypeID As Integer) As Long
  
  Dim strSQL As String
  Dim rstTemp As New ADODB.Recordset
  Dim datData As New clsDataAccess
  
  strSQL = "SELECT ASRBaseTableID FROM ASRSysAccordTransferTypes" _
    & " WHERE TransferTypeID = " & iTransferTypeID
  Set rstTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  If rstTemp.BOF And rstTemp.EOF Then
    GetBaseTableID = 0
  Else
    GetBaseTableID = rstTemp.Fields("ASRBaseTableID").Value
  End If

End Function

Public Sub SendAccordTransactions(ByVal piTransferType As Integer, ByVal iTransactionType As Integer, Optional plngRecordID As Long)

  On Error GoTo ErrorTrap
  
  Dim strSQL As String
  Dim strSQLWhere As String
  Dim strSQLAccordWhere As String
  Dim clsGeneral As clsGeneral
  Dim datData As clsDataAccess
  Dim objFilter As clsExprExpression
  Dim rstTemp As ADODB.Recordset
  Dim strFilter As String
  Dim strPickListIDs As String
  Dim bOK As Boolean
  Dim strErrorString As String
  Dim lngBaseTableID As Long
  Dim strBaseTableName As String
  Dim strColumnName As String
  Dim lngAffectedRecords As Long
  Dim objDeadlock As clsDeadlock
  Dim blnUpdated As Boolean
  Dim lngStartTransactionID As Long
  Dim bBypassFilter As Boolean

  On Local Error GoTo ErrorTrap

  bOK = True
  strSQLWhere = "''"
  
  Set clsGeneral = New clsGeneral
  Set datData = New clsDataAccess
  Set rstTemp = New ADODB.Recordset
  
  ' Start export
  With gobjProgress
    '.AviFile = App.Path & "\videos\export.avi"
    .AVI = dbAccord
    .MainCaption = "Administer Transfers"
    .NumberOfBars = 1
    .Caption = "Payroll Transfer records"
    .Bar1Value = 5
    .Time = False
    .Cancel = True
    .Bar1MaxValue = 100
    .OpenProgress
  End With
  
  gobjEventLog.AddHeader eltAccordExport, "Payroll Transfer"
   
  ' Start Transaction
  gADOCon.BeginTrans
   
  bBypassFilter = IIf(plngRecordID > 0, False, (chkBypassFilter.Value = vbChecked))
  lngBaseTableID = GetBaseTableID(piTransferType)
    
  ' Get the transfer types
  strSQL = "SELECT c.TableID, MIN(c.ColumnID) AS ColumnID, t.TransferTypeID FROM ASRSysColumns c" _
    & " INNER JOIN ASRSysAccordTransferTypes t ON t.ASRBaseTableID = c.TableID" _
    & " WHERE ColumnName NOT IN ('Timestamp', 'ID')" _
    & " AND t.TransferTypeID = " & piTransferType _
    & " GROUP BY c.TableID, t.TransferTypeID"
  Set rstTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
 
  If Not (rstTemp.BOF And rstTemp.EOF) Then
    bOK = True
    strBaseTableName = clsGeneral.GetTableName(rstTemp.Fields("TableID").Value)
    strColumnName = clsGeneral.GetColumnName(rstTemp.Fields("ColumnID").Value)
  Else
    strErrorString = "Cannot find transfer type."
    bOK = False
    strBaseTableName = ""
    strColumnName = ""
  End If
  
  If bOK Then
  
    If plngRecordID > 0 Then
      strSQLWhere = " WHERE ID = " & plngRecordID
      strSQLAccordWhere = " WHERE HRProRecordID = " & plngRecordID
  
    ElseIf optAllRecords.Value = True Then
      strSQLWhere = ""
      bOK = True
    
    ElseIf optPicklist.Value = True Then
    
      Set rstTemp = datData.OpenRecordset("EXEC sp_ASRGetPickListRecords " & txtPicklist.Tag, adOpenForwardOnly, adLockReadOnly)
      bOK = True
      
      If rstTemp.BOF And rstTemp.EOF Then
        strErrorString = "You must select a picklist, or change the record selection for your transfer table."
        bOK = False
      End If
        
      Do While Not rstTemp.EOF
        strPickListIDs = strPickListIDs & IIf(Len(strPickListIDs) > 0, ", ", "") & rstTemp.Fields(0)
        rstTemp.MoveNext
      Loop
      
      rstTemp.Close
  
      strSQLWhere = " WHERE ID IN (" & strPickListIDs & " )"
      strSQLAccordWhere = " WHERE HRProRecordID IN (" & strPickListIDs & " )"
       
    ElseIf optFilter.Value = True Then
        
      ReDim alngSourceTables(2, 0)
      Set objFilter = New clsExprExpression
      bOK = objFilter.Initialise(lngBaseTableID, txtFilter.Tag, giEXPR_RUNTIMECALCULATION, giEXPRVALUE_LOGIC)
      If bOK Then
        bOK = objFilter.RuntimeFilterCode(strFilter, False)
      
        ' Load any required UDFs
        If bOK And gbEnableUDFFunctions Then
          bOK = objFilter.UDFFilterCode(mvarUDFsRequired(), False)
        End If
      
      End If
      
      If Not bOK Then
        strErrorString = "You must select a filter, or change the record selection for your transfer table."
      End If
      
      strSQLWhere = " WHERE id IN (" & strFilter & " )"
      strSQLAccordWhere = " WHERE HRProRecordID IN (" & strFilter & " )"
       
    End If
                  
    
    ' Run the selected generation of transfer records
    If bOK Then
      
      UDFFunctions mvarUDFsRequired, True
      
      gobjProgress.Bar1Caption = "Running Payroll Transfer..."
      gobjProgress.Bar1Value = 15
      gobjProgress.UpdateProgress
      
      ' Payroll - The transfer type selected
      strSQL = "IF NOT EXISTS (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordTransferType' AND [SettingKey] = @@SPID) " _
              & " INSERT ASRSysSystemSettings ([Section],[SettingKey],[SettingValue]) " _
              & " VALUES ('TMP_AccordTransferType',@@SPID," & piTransferType & ")"
      gADOCon.Execute strSQL, , adExecuteNoRecords
      
      ' Payroll - Do we bypass the system filter?
      strSQL = "IF NOT EXISTS (SELECT SettingValue FROM ASRSysSystemSettings WHERE [Section] = 'TMP_AccordBypassFilter' AND [SettingKey] = @@SPID) " _
              & " INSERT ASRSysSystemSettings ([Section],[SettingKey],[SettingValue]) " _
              & " VALUES ('TMP_AccordBypassFilter',@@SPID," & IIf(bBypassFilter, "1", "0") & ")"
      gADOCon.Execute strSQL, , adExecuteNoRecords
      
      Set objDeadlock = New clsDeadlock

      strSQL = "SELECT MAX(TransactionID) FROM asrSysAccordTransactions"
      Set rstTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
      If Not (rstTemp.EOF And rstTemp.BOF) Then
        lngStartTransactionID = IIf(IsNull(rstTemp.Fields(0).Value), 0, rstTemp.Fields(0).Value)
      Else
        lngStartTransactionID = 0
      End If
      
      ' Mark existing transactions for these records as blocked-automatic (type 31)
      strSQL = "UPDATE asrSysAccordTransactions SET Status = " & ACCORD_STATUS_VOID & strSQLAccordWhere & IIf(Len(strSQLAccordWhere) > 0, " AND ", " WHERE ") & "TransferType = " & piTransferType
      blnUpdated = objDeadlock.UpdateTableRecordJustDoIt(strSQL)
      
      strSQL = "UPDATE " & strBaseTableName & " SET " & strColumnName & " = " & strColumnName & strSQLWhere
      blnUpdated = objDeadlock.UpdateTableRecordJustDoIt(strSQL)

      If blnUpdated = False Then
        strErrorString = objDeadlock.ErrorString
        bOK = False
      Else
      
        ' Update the newly created records with the selected transfer type details.
        If chkManualType.Value = vbChecked Then
          strSQL = "UPDATE asrSysAccordTransactions SET TransactionType = " & GetComboItem(cboType) _
            & ", Status = " & GetComboItem(cboStatus) _
            & " WHERE TransactionID > " & lngStartTransactionID
          blnUpdated = objDeadlock.UpdateTableRecordJustDoIt(strSQL)
        ElseIf plngRecordID > 0 Then
          strSQL = "UPDATE asrSysAccordTransactions SET TransactionType = " & iTransactionType _
            & " WHERE TransactionID > " & lngStartTransactionID
          blnUpdated = objDeadlock.UpdateTableRecordJustDoIt(strSQL)
        End If
      
      End If
      
      UDFFunctions mvarUDFsRequired, False
                          
    End If

EndExport:

    ' Clear the Payroll trigger settings
    strSQL = "DELETE FROM asrSysSystemSettings WHERE [Section] = 'TMP_AccordTransferType' AND [SettingKey] = @@SPID"
    gADOCon.Execute strSQL, , adExecuteNoRecords

    strSQL = "DELETE FROM asrSysSystemSettings WHERE [Section] = 'TMP_AccordBypassFilter' AND [SettingKey] = @@SPID"
    gADOCon.Execute strSQL, , adExecuteNoRecords


    If gobjProgress.Cancelled Then
      strErrorString = "Cancelled by user."
      gobjEventLog.ChangeHeaderStatus elsCancelled, lngAffectedRecords, 0
    ElseIf bOK Then
      strErrorString = "Transfer completed successfully."
      gobjEventLog.ChangeHeaderStatus elsSuccessful, lngAffectedRecords, 0
    Else
      gobjEventLog.ChangeHeaderStatus elsFailed, lngAffectedRecords, 0
      gobjEventLog.AddDetailEntry strErrorString
      strErrorString = "Failed." & vbNewLine & vbNewLine & strErrorString
    End If
  
    gobjProgress.CloseProgress
   
  End If

TidyUpAndExit:

  ' Commit transaction
  If bOK Then
    gADOCon.CommitTrans
    
    If Not gblnBatchMode And Not plngRecordID > 0 Then
      COAMsgBox strErrorString, IIf(bOK, vbInformation, vbExclamation) + vbOKOnly, "Administer Transfers"
    End If
    
  Else
    If InStr(1, strErrorString, "SELECT Permission", vbTextCompare) And Not gblnBatchMode Then
      COAMsgBox "You need access to the " & strBaseTableName & " base table to use this functionality.", vbInformation, Me.Caption
    End If
  
    gADOCon.RollbackTrans
  End If

  Unload Me
  Exit Sub

ErrorTrap:
  bOK = False
  strErrorString = Err.Description
  GoTo TidyUpAndExit
  
End Sub


