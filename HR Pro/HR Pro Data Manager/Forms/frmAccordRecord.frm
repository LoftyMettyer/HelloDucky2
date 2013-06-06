VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAccordRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payroll Transfer Details"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1148
   Icon            =   "frmAccordRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDates 
      Caption         =   "Date : "
      Height          =   885
      Left            =   135
      TabIndex        =   15
      Top             =   1815
      Width           =   7665
      Begin VB.TextBox txtCreatedDate 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1740
         TabIndex        =   17
         Top             =   300
         Width           =   1950
      End
      Begin VB.TextBox txtTransferredDate 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5520
         TabIndex        =   16
         Top             =   300
         Width           =   1950
      End
      Begin VB.Label lblCreatedDate 
         Caption         =   "Created Date :"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   345
         Width           =   1635
      End
      Begin VB.Label lblTransferredDate 
         Caption         =   "Transfer Date :"
         Height          =   195
         Left            =   3880
         TabIndex        =   18
         Top             =   345
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   6600
      TabIndex        =   13
      Top             =   7860
      Width           =   1200
   End
   Begin VB.Frame frmStatus 
      Caption         =   "Warnings : "
      Height          =   1500
      Left            =   135
      TabIndex        =   11
      Top             =   6285
      Width           =   7665
      Begin VB.TextBox txtReason 
         BackColor       =   &H8000000F&
         Height          =   900
         Left            =   225
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   360
         Width           =   7260
      End
   End
   Begin VB.Frame frmDetails 
      Caption         =   "Transfer : "
      Height          =   1635
      Left            =   135
      TabIndex        =   3
      Top             =   135
      Width           =   7665
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   1755
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1125
         Width           =   1965
      End
      Begin VB.ComboBox cboTransactionType 
         Height          =   315
         ItemData        =   "frmAccordRecord.frx":000C
         Left            =   5535
         List            =   "frmAccordRecord.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   720
         Width           =   1950
      End
      Begin VB.TextBox txtTransferType 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5535
         TabIndex        =   6
         Top             =   315
         Width           =   1950
      End
      Begin VB.TextBox txtCompanyCode 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1755
         TabIndex        =   5
         Top             =   315
         Width           =   1950
      End
      Begin VB.TextBox txtEmployeeCode 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1755
         TabIndex        =   4
         Top             =   720
         Width           =   1950
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status :"
         Height          =   255
         Left            =   225
         TabIndex        =   20
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label lblTransferType 
         Caption         =   "Transfer Type :"
         Height          =   195
         Left            =   3880
         TabIndex        =   10
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label lblCompanyCode 
         Caption         =   "Company Code :"
         Height          =   195
         Left            =   225
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblTransactionType 
         Caption         =   "Transaction Type :"
         Height          =   195
         Left            =   3880
         TabIndex        =   8
         Top             =   765
         Width           =   1635
      End
      Begin VB.Label lblEmployeeCode 
         Caption         =   "Employee Code :"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   765
         Width           =   1500
      End
   End
   Begin VB.Frame frmData 
      Caption         =   "Data Fields : "
      Height          =   3435
      Left            =   135
      TabIndex        =   1
      Top             =   2775
      Width           =   7665
      Begin SSDataWidgets_B.SSDBGrid grdFieldData 
         Height          =   2940
         Left            =   225
         TabIndex        =   2
         Top             =   315
         Width           =   7245
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   3
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   0
         ForeColorEven   =   -2147483640
         ForeColorOdd    =   -2147483640
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   4736
         Columns(0).Caption=   "Field"
         Columns(0).Name =   "Field"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3836
         Columns(1).Caption=   "Old Data"
         Columns(1).Name =   "Old Data"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3545
         Columns(2).Caption=   "New Data"
         Columns(2).Name =   "New Data"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   12779
         _ExtentY        =   5186
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   400
      Left            =   5310
      TabIndex        =   0
      Top             =   7860
      Width           =   1200
   End
End
Attribute VB_Name = "frmAccordRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miConnectionType As DataMgr.AccordConnection
Private datData As clsDataAccess
Private mbCancelled As Boolean
Private mlngTransactionID As Long
Private mstrStatus As String
Private miWarningErrorCount As Integer
Private mbChanged As Boolean
Private mbEnableChangeStatus As Boolean

Public Property Let ConnectionType(ByVal piNewValue As DataMgr.AccordConnection)
  miConnectionType = piNewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Get Changed() As Boolean
  Changed = mbChanged
End Property

Private Sub PopulateControls()

  Dim sSQL As String
  Dim rsDetails As ADODB.Recordset

  sSQL = "SELECT tt.TransferType AS TransferName, t.* FROM ASRSysAccordTransactions t" _
        & " INNER JOIN ASRSysAccordTransferTypes tt ON tt.TransferTypeID = t.TransferType" _
        & " WHERE t.TransactionID = " & mlngTransactionID
  Set rsDetails = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rsDetails
    If Not (.EOF And .BOF) Then
      txtTransferType.Text = .Fields("TransferName").Value
      txtCompanyCode.Text = IIf(IsNull(.Fields("CompanyCode").Value), "", .Fields("CompanyCode").Value)
      txtEmployeeCode.Text = IIf(IsNull(.Fields("EmployeeCode").Value), "", .Fields("EmployeeCode").Value)
      txtCreatedDate.Text = IIf(IsNull(.Fields("CreatedDateTime").Value), "", .Fields("CreatedDateTime").Value)
      cboTransactionType.ListIndex = .Fields("TransactionType").Value
      SetComboItem cboStatus, IIf(IsNull(.Fields("Status").Value), 1, .Fields("Status").Value)
      txtTransferredDate.Text = IIf(IsNull(.Fields("TransferedDateTime").Value), "", .Fields("TransferedDateTime").Value)
            
      Select Case IIf(IsNull(.Fields("Status")), 0, .Fields("Status"))
        Case ACCORD_STATUS_PENDING
          frmDetails.Caption = "Pending Transfer : "
    
        Case ACCORD_STATUS_SUCCESS
          frmDetails.Caption = "Successful Transfer : "
        
        Case ACCORD_STATUS_SUCCESS_WARNINGS
          frmDetails.Caption = "Successful with Warnings : "
        
        Case ACCORD_STATUS_FAILURE_UNKNOWN
          frmDetails.Caption = "Unknown Failure : "
          
        Case ACCORD_STATUS_IGNORED
          frmDetails.Caption = "Ignored : "
        
        Case ACCORD_STATUS_ALREADY_EXISTS
          frmDetails.Caption = "Record Already Exists : "
  
        Case ACCORD_STATUS_MOREINFO_REQUIRED
          frmDetails.Caption = "More Information Required : "
        
        Case ACCORD_STATUS_BLOCKED
          frmDetails.Caption = "Blocked Transfer : "
          
        Case ACCORD_STATUS_VOID
          frmDetails.Caption = "Void Transfer : "
          
      End Select
          
    End If
  End With

  Set rsDetails = Nothing


End Sub

Private Sub PopulateFieldDataGrid()

  Dim sSQL As String
  Dim rsDetails As ADODB.Recordset
  Dim strOldData As String
  Dim strNewData As String

  sSQL = "SELECT td.*, tf.[Description] FROM ASRSysAccordTransactionData td" _
      & " INNER JOIN ASRSysAccordTransactions t ON td.TransactionID = t.TransactionID" _
      & " INNER JOIN ASRSysAccordTransferFieldDefinitions tf ON td.FieldID = tf.TransferFieldID AND t.TransferType = tf.TransferTypeID" _
      & " WHERE td.TransactionID = " & mlngTransactionID _
      & " ORDER BY td.FieldID"
  
  Set rsDetails = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)

  With rsDetails
    While Not .EOF
    
      strOldData = IIf(IsNull(.Fields("OldData")), "<Empty>", .Fields("OldData"))
      strNewData = IIf(IsNull(.Fields("NewData")), "<Empty>", .Fields("NewData"))
      grdFieldData.AddItem .Fields("Description") & vbTab & strOldData & vbTab & strNewData & vbTab
      .MoveNext
    Wend
  End With
  
  Set rsDetails = Nothing

End Sub

Public Property Let TransactionID(plngID As Long)
  mlngTransactionID = plngID
End Property

Private Sub cboStatus_Click()
  
  Dim iNewStatus As AccordTransactionStatus
  
  iNewStatus = cboStatus.ItemData(cboStatus.ListIndex)
  
  ' Change icon at top
  Select Case iNewStatus
  
    Case ACCORD_STATUS_UNKNOWN
      frmDetails.Caption = "Unknown Status : "
  
    Case ACCORD_STATUS_PENDING
      frmDetails.Caption = "Pending Transfer : "
  
    Case ACCORD_STATUS_SUCCESS
      frmDetails.Caption = "Successful Transfer : "
    
    Case ACCORD_STATUS_SUCCESS_WARNINGS
      frmDetails.Caption = "Successful with Warnings : "
    
    Case ACCORD_STATUS_FAILURE_UNKNOWN
      frmDetails.Caption = "Unknown Failure : "
      
    Case ACCORD_STATUS_IGNORED
      frmDetails.Caption = "Ignored : "
    
    Case ACCORD_STATUS_ALREADY_EXISTS
      frmDetails.Caption = "Record Already Exists : "
  
    Case ACCORD_STATUS_MOREINFO_REQUIRED
      frmDetails.Caption = "More Information Required : "
    
    Case ACCORD_STATUS_DOESNOT_EXIST
      frmDetails.Caption = "Record Does Not Exist : "
    
    Case ACCORD_STATUS_BLOCKED
      frmDetails.Caption = "Blocked Transfer : "
      
    Case ACCORD_STATUS_VOID
      frmDetails.Caption = "Void Transfer : "
      
  End Select
  
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cboTransactionType_Change()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cboTransactionType_Click()
  mbChanged = True
  RefreshButtons
End Sub

Private Sub cmdCancel_Click()
  mbCancelled = True
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If SaveChanges Then
    mbCancelled = False
    mbChanged = True
    Unload Me
  End If
End Sub

Public Sub Initialise()
  
  mbEnableChangeStatus = GetModuleParameter(gsMODULEKEY_ACCORD, gsPARAMETERKEY_ALLOWSTATUSCHANGE)
  
  PopulateCombos
  PopulateControls
  PopulateFieldDataGrid
  PopulateWarningsAndErrors

  EnableControl cboStatus, mbEnableChangeStatus
  EnableControl cboTransactionType, mbEnableChangeStatus

End Sub

' Display all the warnings associated with this transaction
Private Sub PopulateWarningsAndErrors()

  Dim sSQL As String
  Dim rsDetails As ADODB.Recordset
  Dim strWarnings As String

  sSQL = "SELECT ErrorText FROM ASRSysAccordTransactions WHERE TransactionID = " & mlngTransactionID
  Set rsDetails = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsDetails
  
    If Not (.EOF And .BOF) Then
      strWarnings = IIf(IsNull(.Fields("ErrorText").Value), "", .Fields("ErrorText").Value)
    End If
    .Close
  End With

  sSQL = "SELECT [WarningMessage] FROM ASRSysAccordTransactionWarnings w" _
      & " WHERE w.TransactionID = " & mlngTransactionID & " ORDER BY w.FieldID"

  Set rsDetails = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  With rsDetails
    While Not .EOF
      strWarnings = strWarnings & IIf(Len(strWarnings) > 0, vbNewLine, "") _
          & IIf(IsNull(.Fields("WarningMessage").Value), "", Trim(.Fields("WarningMessage").Value))
      miWarningErrorCount = miWarningErrorCount + 1
      .MoveNext
    Wend
    .Close
  End With
  
  txtReason.Text = strWarnings
  Set rsDetails = Nothing
  mbChanged = False
  
  RefreshButtons

End Sub

Private Sub Form_Initialize()
  Set datData = New clsDataAccess
  miWarningErrorCount = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case vbKeyEscape
    Unload Me
End Select

End Sub

Private Sub RefreshButtons()
  cmdOk.Enabled = mbChanged
End Sub

Private Function SaveChanges() As Boolean
  
  On Error GoTo ErrorTrap
  
  Dim bOK As Boolean
  Dim sSQL As String
  
  bOK = True
  sSQL = "UPDATE [ASRSysAccordTransactions] SET [TransactionType] = " & cboTransactionType.ListIndex & _
    ", [Status] = " & GetComboItem(cboStatus) & _
    " WHERE TransactionID = " & mlngTransactionID
  
  gADOCon.Execute sSQL, , adExecuteNoRecords

TidyUpAndExit:
  SaveChanges = bOK
  Exit Function

ErrorTrap:
  bOK = False
  GoTo TidyUpAndExit

End Function

Private Sub PopulateCombos()

  With cboTransactionType
    .AddItem "New", 0
    .AddItem "Update", 1
    .AddItem "Delete", 2
  End With

  With cboStatus
     
    .AddItem "Unknown", 0
    .ItemData(0) = ACCORD_STATUS_UNKNOWN
    
    .AddItem "Pending", 1
    .ItemData(1) = ACCORD_STATUS_PENDING
    
    .AddItem "Success", 2
    .ItemData(2) = ACCORD_STATUS_SUCCESS
    
    .AddItem "Success With Warnings", 3
    .ItemData(3) = ACCORD_STATUS_SUCCESS_WARNINGS
    
    .AddItem "Unknown Failure", 4
    .ItemData(4) = ACCORD_STATUS_FAILURE_UNKNOWN
    
    .AddItem "Ignored", 5
    .ItemData(5) = ACCORD_STATUS_IGNORED
    
    .AddItem "Record Already Exists", 6
    .ItemData(6) = ACCORD_STATUS_ALREADY_EXISTS
    
    .AddItem "More Information Required", 7
    .ItemData(7) = ACCORD_STATUS_MOREINFO_REQUIRED
        
    .AddItem "Record Does Not Exist", 8
    .ItemData(8) = ACCORD_STATUS_DOESNOT_EXIST
        
    .AddItem "Blocked", 9
    .ItemData(9) = ACCORD_STATUS_BLOCKED
    
    .AddItem "Void", 10
    .ItemData(10) = ACCORD_STATUS_VOID
    
  End With

End Sub

