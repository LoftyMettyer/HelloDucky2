VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmEmailQueue 
   Caption         =   "Email Queue"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1032
   Icon            =   "frmEmailQueue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   13530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdView 
      Caption         =   "&View..."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12260
      TabIndex        =   11
      Top             =   735
      Width           =   1200
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   14
      Top             =   4845
      Width           =   13530
      _ExtentX        =   23865
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   23336
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdEmailQueue 
      Height          =   3525
      Left            =   90
      TabIndex        =   0
      Top             =   1185
      Width           =   12000
      _Version        =   196617
      DataMode        =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RecordSelectors =   0   'False
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
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ExtraHeight     =   26
      CaptionAlignment=   0
      Columns.Count   =   18
      Columns(0).Width=   3889
      Columns(0).Caption=   "Link Title"
      Columns(0).Name =   "Email Title"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3889
      Columns(1).Caption=   "Record Description"
      Columns(1).Name =   "RecDesc"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Table Name"
      Columns(2).Name =   "Table Name"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3889
      Columns(3).Caption=   "Column Name"
      Columns(3).Name =   "Column Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3201
      Columns(4).Caption=   "Column Value"
      Columns(4).Name =   "Column Value"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1773
      Columns(5).Caption=   "Email Due"
      Columns(5).Name =   "Email Due"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1852
      Columns(6).Caption=   "Email Sent"
      Columns(6).Name =   "Email Sent"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "QueueID"
      Columns(7).Name =   "QueueID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "RepTo"
      Columns(8).Name =   "RepTo"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "RepCC"
      Columns(9).Name =   "RepCC"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "RepBCC"
      Columns(10).Name=   "RepBCC"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   3200
      Columns(11).Visible=   0   'False
      Columns(11).Caption=   "Subject"
      Columns(11).Name=   "Subject"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Visible=   0   'False
      Columns(12).Caption=   "Attachment"
      Columns(12).Name=   "Attachment"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Visible=   0   'False
      Columns(13).Caption=   "MsgText"
      Columns(13).Name=   "MsgText"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   3200
      Columns(14).Visible=   0   'False
      Columns(14).Caption=   "LinkID"
      Columns(14).Name=   "LinkID"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      Columns(15).Width=   3200
      Columns(15).Visible=   0   'False
      Columns(15).Caption=   "RecordID"
      Columns(15).Name=   "RecordID"
      Columns(15).DataField=   "Column 15"
      Columns(15).DataType=   8
      Columns(15).FieldLen=   256
      Columns(16).Width=   3200
      Columns(16).Visible=   0   'False
      Columns(16).Caption=   "RecalculateRecordDesc"
      Columns(16).Name=   "RecalculateRecordDesc"
      Columns(16).DataField=   "Column 16"
      Columns(16).DataType=   8
      Columns(16).FieldLen=   256
      Columns(17).Width=   3200
      Columns(17).Visible=   0   'False
      Columns(17).Caption=   "WorkflowInstanceID"
      Columns(17).Name=   "WorkflowInstanceID"
      Columns(17).DataField=   "Column 17"
      Columns(17).DataType=   8
      Columns(17).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   21167
      _ExtentY        =   6218
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.CommandButton cmdPurge 
      Caption         =   "&Purge..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12260
      TabIndex        =   13
      Top             =   1845
      Width           =   1200
   End
   Begin VB.CommandButton cmdRebuild 
      Caption         =   "Re&build..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12260
      TabIndex        =   12
      Top             =   1290
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   12260
      TabIndex        =   10
      Top             =   180
      Width           =   1200
   End
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   12020
      Begin VB.ComboBox cboRecDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   500
         Width           =   2820
      End
      Begin VB.ComboBox cboStatus 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEmailQueue.frx":000C
         Left            =   10700
         List            =   "frmEmailQueue.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   500
         Width           =   1170
      End
      Begin VB.ComboBox cboTitle 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   500
         Width           =   2820
      End
      Begin VB.ComboBox cboTable 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   500
         Width           =   2505
      End
      Begin VB.Label lblRecDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Description :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3000
         TabIndex        =   4
         Top             =   250
         Width           =   1755
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10700
         TabIndex        =   8
         Top             =   250
         Width           =   705
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Link Title :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   250
         Width           =   945
      End
      Begin VB.Label lblTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table Name :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5760
         TabIndex        =   6
         Top             =   250
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmEmailQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declare sizing constants
Const BUTTON_GAP = 240
Const BUTTON_WIDTH = 1200
Const GAP_AFTER_LISTVIEW = 1850

' Flag to prevent grid refreshing when combos are being populated and set initially
Private mblnLoading As Boolean

' Flag showing if user is viewing all or just his own entries
Private mblnViewAllEntries As Boolean

' Must be public so the details form can change the bookmark of the recordset
Public mrstHeaders As Recordset

' Data access class
Private mclsData As New clsDataAccess

' Variables to hold the column clicked on, its field and the order to sort the grid
Private mintOrderCol As Integer
Private mblnOrderDesc As Boolean

Private mblnDeleteEnabled As Boolean
Private mblnPurgeEnabled As Boolean

Private fEmailQueueDetails As frmEmailQueueDetails

Private mlngMaxWidth(6) As Long

Private Sub cboRecDesc_Click()
  RefreshGrid
End Sub

Private Sub cboTable_Click()
  RefreshGrid
End Sub

Private Sub cboStatus_Click()
  RefreshGrid
End Sub

Private Sub cboTitle_Click()
  RefreshGrid
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdPurge_Click()

  frmEmailQueuePurge.Show vbModal
  RefreshGrid

End Sub

Private Sub cmdRebuild_Click()

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBCaption As String
  Dim intMBResponse As Integer
  
  On Error GoTo LocalErr
  
  strMBText = "Are you sure that you would like to rebuild the email queue?"
  intMBButtons = vbQuestion + vbYesNo
  strMBCaption = Me.Caption
  intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBCaption)
  
  If intMBResponse <> vbYes Then
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  With gobjProgress
    .AVI = dbEMail
    .MainCaption = "Email Queue"
    .NumberOfBars = 0
    .Caption = "Email Queue Rebuild"
    .Time = False
    .Cancel = False
    .OpenProgress
    .Bar1Caption = "Email Queue Rebuild"
  End With

  ' Event Log Header
  gobjEventLog.AddHeader eltEmailRebuild, "Email Rebuild"

  datGeneral.ExecuteSql "EXEC spASREmailRebuild", ""
  Call RefreshGrid
  
  gobjProgress.CloseProgress
  Screen.MousePointer = vbDefault

  ' Event Log Header
  gobjEventLog.ChangeHeaderStatus elsSuccessful

Exit Sub

LocalErr:
  COAMsgBox "Error rebuilding email queue", vbCritical, Me.Caption
  ' Event Log Header
  gobjEventLog.ChangeHeaderStatus elsFailed

End Sub

Private Sub Form_Activate()

  DoColumnSizes

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
    Case KeyCode = vbKeyEscape
      Unload Me
    Case KeyCode = vbKeyF5
      RefreshGrid
  End Select
End Sub

Private Sub Form_Load()
  
  Hook Me.hWnd, 11700, 5550
  
  Dim rstUsers As Recordset
  Set mclsData = New clsDataAccess
  
  mblnLoading = True

  grdEmailQueue.RowHeight = 239


  'MH20070711 Fault 12350 Discussed with QA and Lofty and agreed that we don't need to update records descriptions here
  ''''This will update all of the record descriptions in the email queue
  '''mclsData.ExecuteSql "spASREmailQueue"

  If datGeneral.SystemPermission("EMAIL", "REBUILDPURGE") = False Then
    cmdPurge.Enabled = False
    cmdRebuild.Enabled = False
  End If


  ' Set height and width to last saved. Form is centred on screen
  Me.Height = GetPCSetting(gsDatabaseName & "\EmailQueue", "Height", Me.Height)
  Me.Width = GetPCSetting(gsDatabaseName & "\EmailQueue", "Width", Me.Width)

  ' Set default sort order to be date desc
  mintOrderCol = 5
  mblnOrderDesc = True

  PopulateCombos
  mblnLoading = False
  
  ' Populate the grid
  RefreshGrid

  Set fEmailQueueDetails = New frmEmailQueueDetails
  
  ' Get rid of the icon off the form
  RemoveIcon Me

  Screen.MousePointer = vbDefault

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  ' Save the window size ready to recall next time user views the event log
  SavePCSetting gsDatabaseName & "\EmailQueue", "Height", Me.Height
  SavePCSetting gsDatabaseName & "\EmailQueue", "Width", Me.Width

End Sub

Private Sub Form_Resize()
  
  Const COMBO_GAP As Integer = 170
  
  Dim lngComboWidth As Long
  
  'JPD 20030908 Fault 5756
  DisplayApplication

  ' Ensure form does not get too small/big. Also reposition controls as necessary
'  If Me.Width < 11700 Or Me.Width > Screen.Width Then
'    Me.Width = 11700
'  End If
'  If Me.Height < 5550 Or Me.Height > Screen.Height Then
'    Me.Height = 5550
'  End If

  'AE20071005 Fault #12196
  'cmdOK.Left = Me.Width - (BUTTON_GAP + BUTTON_WIDTH)
  cmdOK.Left = Me.ScaleWidth - (BUTTON_WIDTH + (BUTTON_GAP / 2))
  cmdView.Left = cmdOK.Left
  cmdRebuild.Left = cmdOK.Left
  cmdPurge.Left = cmdOK.Left

  fraFilters.Width = cmdOK.Left - BUTTON_GAP
  
  cboStatus.Left = fraFilters.Width - (cboStatus.Width + COMBO_GAP)
  lblStatus.Left = cboStatus.Left
  
  lngComboWidth = (cboStatus.Left - (COMBO_GAP * 4)) / 3
  
  cboTitle.Move COMBO_GAP, 500, lngComboWidth
  lblTitle.Left = cboTitle.Left
  
  cboRecDesc.Move cboTitle.Left + cboTitle.Width + COMBO_GAP, 500, lngComboWidth
  lblRecDesc.Left = cboRecDesc.Left

  cboTable.Move cboRecDesc.Left + cboRecDesc.Width + COMBO_GAP, 500, lngComboWidth
  lblTable.Left = cboTable.Left
  
  
  'NPG20080416 Suggestion
  'cboTitle.Width = fraFilters.Width * 0.25
  
  'lblType.Left = fraFilters.Width * 0.4
  'cboTable.Left = fraFilters.Width * 0.52
  'cboTable.Width = fraFilters.Width * 0.2
  
  'lblStatus.Left = fraFilters.Width * 0.79
  'cboStatus.Left = fraFilters.Width * 0.87
  
    
  grdEmailQueue.Width = fraFilters.Width
  'grdEmailQueue.Height = Me.Height - GAP_AFTER_LISTVIEW
  grdEmailQueue.Height = Me.Height - (Me.Height - Me.ScaleHeight) - 1500

  DoColumnSizes
  Me.Refresh
  grdEmailQueue.Refresh

  End Sub

Private Sub DoColumnSizes()

  Const lngEMAILDUEWIDTH As Integer = 1100
  Const lngEMAILSENTWIDTH As Integer = 1725

  Dim lngScrollBarWidth As Long
  Dim lngColumnsWidth As Long
  Dim lngSpareWidth As Long
  Dim lngIndex As Long
  Dim dblFactor As Double

  With grdEmailQueue

    lngScrollBarWidth = IIf(.Rows > .VisibleRows, 255, 0)
    lngSpareWidth = .Width - (lngEMAILDUEWIDTH + lngEMAILSENTWIDTH + lngScrollBarWidth + 20)

    For lngIndex = 0 To 4
      .Columns(lngIndex).Width = (lngSpareWidth / 5)
    Next

    .Columns(5).Width = lngEMAILDUEWIDTH
    .Columns(6).Width = lngEMAILSENTWIDTH

  End With

End Sub

Private Function RefreshGrid() As Boolean

  Dim pstrSQL As String
  Dim strWhere As String
  Dim strTitle As String
  Dim lngIndex As Long
    
  If mblnLoading = True Then Exit Function

  Screen.MousePointer = vbHourglass
  
  
  strWhere = ""
  
  With cboTitle
    If .ListIndex > 0 Then
      strWhere = IIf(strWhere <> "", strWhere & " AND ", "") & _
          "isnull(w.name+' (Workflow)', isnull(l.Title, CASE WHEN q.UserName = 'OpenHR Mobile' THEN q.Subject WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.Subject ELSE '' END)) = '" & Replace(.Text, "'", "''") & "'"
    End If
  End With
  
  With cboRecDesc
    If .ListIndex > 0 Then
      strWhere = IIf(strWhere <> "", strWhere & " AND ", "") & _
          "CASE WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.RepTo ELSE q.RecordDesc END = '" & Replace(.Text, "'", "''") & "'"
                    
    End If
  End With
  
  With cboTable
    If .ListIndex > 0 Then
      strWhere = IIf(strWhere <> "", strWhere & " AND ", "") & _
          "t.TableName = '" & Replace(.Text, "'", "''") & "'"
    End If
  End With
  
  Select Case cboStatus.ListIndex
    Case 1
      strWhere = IIf(strWhere <> "", strWhere & " AND ", "") & _
        "q.DateSent IS NULL"
    Case 2
      strWhere = IIf(strWhere <> "", strWhere & " AND ", "") & _
        "q.DateSent IS NOT NULL"
  End Select

' ' CASE WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.RepTo ELSE isnull(q.RecordDesc,'') END as [RecDesc]     ,

  pstrSQL = _
        "SELECT isnull(w.name+' (Workflow)',isnull(l.Title,CASE WHEN q.UserName = 'OpenHR Mobile' THEN q.Subject WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.Subject ELSE '' END)) as [QueueTitle]" & _
        "     , CASE WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.RepTo ELSE isnull(q.RecordDesc,'') END as [RecDesc]" & _
        "     , isnull(t.TableName,'') as [TableName]" & _
        "     , case when q.WorkflowInstanceID > 0 or q.columnID IS NULL then '' else isnull(c.ColumnName,'<Multiple Columns>') end as [ColumnName]" & _
        "     , case when q.WorkflowInstanceID > 0 or q.columnID IS NULL then '' else isnull(q.ColumnValue,'<Multiple Values>') end as [ColumnValue]" & _
        "     , q.DateDue as [DateDue]" & _
        "     , q.DateSent as [DateSent]" & _
        "     , isnull(q.QueueID,0) as [QueueID]" & _
        "     , isnull(q.RepTo,'') as [RepTo]" & _
        "     , isnull(q.RepCC,'') as [RepCC]" & _
        "     , isnull(q.RepBCC,'') as [RepBCC]" & _
        "     , isnull(q.Subject,'') as [Subject]" & _
        "     , isnull(q.Attachment,'') as [Attachment]" & _
        "     , isnull(q.MsgText,'') as [MsgText]" & _
        "     , isnull(q.LinkID,0) as [LinkID]" & _
        "     , isnull(q.RecordID,0) as [RecordID]" & _
        "     , isnull(q.RecalculateRecordDesc,0) as [RecalculateRecordDesc]" & _
        "     , isnull(q.WorkflowInstanceID,0) as [WorkflowInstanceID] " & _
        "FROM ASRSysEmailQueue q "
  
  pstrSQL = pstrSQL & _
        "LEFT OUTER JOIN ASRSysEmailLinks l ON q.LinkID = l.LinkID " & _
        "LEFT OUTER JOIN ASRSysTables t ON q.TableID = t.TableID " & _
        "LEFT OUTER JOIN ASRSysColumns c ON q.ColumnID = c.ColumnID " & _
        "LEFT OUTER JOIN ASRSysWorkflowInstances wi ON q.workflowInstanceID = wi.ID " & _
        "LEFT OUTER JOIN ASRSysWorkflows w ON wi.workflowID = w.ID "

  If strWhere <> "" Then
    pstrSQL = pstrSQL & " WHERE " & strWhere
  End If
  
  pstrSQL = pstrSQL & " ORDER BY " & CStr(mintOrderCol + 1) & " " & IIf(mblnOrderDesc, "DESC", "ASC")
  
  
  
  'Set mrstHeaders = mclsData.OpenPersistentRecordset(pstrSQL, adOpenKeyset, adLockReadOnly)
  Set mrstHeaders = mclsData.OpenPersistentRecordset(pstrSQL, adOpenStatic, adLockReadOnly)
  
  'Reset column widths
  For lngIndex = 0 To UBound(mlngMaxWidth)
    mlngMaxWidth(lngIndex) = Me.TextWidth(grdEmailQueue.Columns(lngIndex).Caption)
  Next
  
  'Rebind data
  With grdEmailQueue
    .Redraw = False
    .Rebind
    .Rows = mrstHeaders.RecordCount
    .Redraw = True
  End With

  cmdView.Enabled = (grdEmailQueue.Rows > 0)
  If grdEmailQueue.Rows > 0 Then
    grdEmailQueue.MoveFirst
    grdEmailQueue.SelBookmarks.Add grdEmailQueue.Bookmark
  End If

  StatusBar1.SimpleText = " " & mrstHeaders.RecordCount & " Record" & IIf(mrstHeaders.RecordCount <> 1, "s", "") & _
    IIf(mrstHeaders.RecordCount > 1, " sorted by '" & grdEmailQueue.Columns(mintOrderCol).Caption & "' in " & IIf(mblnOrderDesc, "descending", "ascending") & " order", "")

  DoColumnSizes

  Screen.MousePointer = vbDefault

End Function

Private Sub cmdView_Click()

  With grdEmailQueue
    If (.Rows > 0) And .SelBookmarks.Count = 1 Then
      fEmailQueueDetails.Initialise grdEmailQueue
    End If

  End With

End Sub

Private Sub Form_Terminate()
  Set fEmailQueueDetails = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdEmailQueue_DblClick()
  If cmdView.Enabled Then
    cmdView_Click
  End If
End Sub


Private Sub grdEmailQueue_HeadClick(ByVal ColIndex As Integer)

  If mintOrderCol = ColIndex Then
    mblnOrderDesc = Not mblnOrderDesc
  Else
    mintOrderCol = ColIndex
    mblnOrderDesc = (mintOrderCol = 5 Or mintOrderCol = 6)  'Descending for dates
  End If

  RefreshGrid

End Sub

'Private Sub grdEmailQueue_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  If KeyCode = 46 And mblnDeleteEnabled = True Then
'    cmdDelete_Click
'  ElseIf KeyCode = 35 And Shift = 2 Then
'    ' ctrl and end pressed
'    'grdEmailQueue.FirstRow = grdEmailQueue.Rows - grdEmailQueue.VisibleRows
'    'grdEmailQueue.MoveLast
'  ElseIf KeyCode = 36 And Shift = 2 Then
'    ' ctrl and home pressed
'    'grdEmailQueue.FirstRow = 0
'    'grdEmailQueue.MoveFirst
'  Else
'    If Shift > 0 Then KeyCode = 0
'  End If
'
'End Sub

Private Sub grdEmailQueue_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)

  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrstHeaders.MoveLast
    Else
      mrstHeaders.MoveFirst
    End If
  Else
    mrstHeaders.Bookmark = StartLocation
  End If

  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrstHeaders.Move NumberOfRowsToMove
  NewLocation = mrstHeaders.Bookmark

End Sub

Private Sub grdEmailQueue_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  Dim strOutput As String
  Dim lngWidth As Long

  sDateFormat = DateFormat

  ' This is required as recordset not set when this sub is first run
  If mrstHeaders Is Nothing Then Exit Sub
  If mrstHeaders.State = adStateClosed Then Exit Sub

  ' Do nothing if we are loading or if there are no records to display
  If mblnLoading = True And mrstHeaders.RecordCount = 0 Then Exit Sub

  If StartLocation < 0 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrstHeaders.EOF Then
        mrstHeaders.MoveLast
      End If
    Else
      If Not mrstHeaders.BOF Then
        mrstHeaders.MoveFirst
      End If
    End If
  Else
    mrstHeaders.Bookmark = StartLocation
    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If
  End If

  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrstHeaders.BOF Or mrstHeaders.EOF Then Exit For

    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrstHeaders.Fields.Count - 1)
          Select Case mrstHeaders.Fields(iFieldIndex).Name
          Case "DateDue"
            strOutput = Format(mrstHeaders.Fields(iFieldIndex), sDateFormat)
          Case "DateSent"
            strOutput = Format(mrstHeaders.Fields(iFieldIndex), sDateFormat & " hh:nn")
          Case Else
            strOutput = mrstHeaders(iFieldIndex)
          End Select
          RowBuf.Value(iRowIndex, iFieldIndex) = strOutput

          If iFieldIndex <= UBound(mlngMaxWidth) Then
            lngWidth = Me.TextWidth(strOutput)
            If mlngMaxWidth(iFieldIndex) < lngWidth Then
              mlngMaxWidth(iFieldIndex) = lngWidth
            End If
          End If

        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 2
        RowBuf.Value(iRowIndex, 0) = mrstHeaders(0)
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark
      Case 3
    End Select

    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If

    iRowsRead = iRowsRead + 1
  Next iRowIndex

  RowBuf.RowCount = iRowsRead

End Sub


'Private Function GetRecordDesc(lngExprID As Long, lngRecordID As Long)
'
'  ' Return TRUE if the user has been granted the given permission.
'  Dim cmADO As ADODB.Command
'  Dim pmADO As ADODB.Parameter
'
'  On Error GoTo LocalErr
'
'  If lngExprID < 1 Then
'    GetRecordDesc = "Record Description Undefined"
'    Exit Function
'  End If
'
'
'  ' Check if the user can create New instances of the given category.
'  Set cmADO = New ADODB.Command
'  With cmADO
'    .CommandText = "dbo.sp_ASRExpr_" & lngExprID
'    .CommandType = adCmdStoredProc
'    .CommandTimeout = 0
'    Set .ActiveConnection = gADOCon
'
'    Set pmADO = .CreateParameter("Result", adVarChar, adParamOutput, VARCHAR_MAX_Size)
'    .Parameters.Append pmADO
'
'    Set pmADO = .CreateParameter("RecordID", adInteger, adParamInput)
'    .Parameters.Append pmADO
'    pmADO.Value = lngRecordID
'
'    cmADO.Execute
'
'    GetRecordDesc = .Parameters(0).Value
'  End With
'  Set cmADO = Nothing
'
'Exit Function
'
'LocalErr:
'  GetRecordDesc = "Error reading record description" '& vbCr & _
'                  "(ID = " & CStr(lngRecordID) & ", Record Description = " & CStr(mlngRecordDescExprID)
'  'fOK = False
'
'End Function


Private Sub PopulateCombos()

  Dim strSQL As String
  

  strSQL = "SELECT DISTINCT isnull(w.name+' (Workflow)',isnull(l.Title,CASE WHEN q.UserName = 'OpenHR  Mobile' THEN q.Subject WHEN q.UserName = 'OpenHR Self-service Intranet' THEN q.Subject ELSE '' END)) " & _
        "FROM ASRSysEmailQueue q " & _
        "LEFT OUTER JOIN ASRSysEmailLinks l ON q.LinkID = l.LinkID " & _
        "LEFT OUTER JOIN ASRSysWorkflowInstances wi ON q.workflowInstanceID = wi.ID " & _
        "LEFT OUTER JOIN ASRSysWorkflows w ON wi.workflowID = w.ID " & _
        "ORDER BY 1"
  PopulateCombo lblTitle, cboTitle, strSQL
  
  ' SELECT DISTINCT isnull(q.RecordDesc,'')
  strSQL = "SELECT DISTINCT (CASE WHEN q.username =  'OpenHR Self-service Intranet' THEN q.RepTo ELSE isnull(q.RecordDesc,'') END) " & _
        "FROM ASRSysEmailQueue q " & _
        "ORDER BY 1"
  PopulateCombo lblRecDesc, cboRecDesc, strSQL


  strSQL = "SELECT DISTINCT isnull(t.TableName,'') " & _
        "FROM ASRSysEmailQueue q " & _
        "LEFT OUTER JOIN ASRSysTables t ON q.TableID = t.TableID " & _
        "ORDER BY 1"
  PopulateCombo lblTable, cboTable, strSQL

  
  With cboStatus
    .Clear
    .AddItem "<All>"
    .AddItem "Not Sent"
    .AddItem "Sent"
    .ListIndex = 1
  End With

  'lblRecDesc.Left = cboTitle.Left + cboTitle.Width + 60
  'cboRecDesc.Left = lblRecDesc.Left
  'lblTable.Left = cboRecDesc.Left + cboRecDesc.Width + 60
  'cboTable.Left = lblTable.Left
  'lblStatus.Left = cboTable.Left + cboTable.Width + 60
  'cboStatus.Left = lblStatus.Left

End Sub


Private Sub PopulateCombo(lbl As Label, cbo As ComboBox, strSQL As String)

  Dim rsTemp As Recordset
  Dim strTemp As String
  Dim lngMaxWidth As Long
  Dim lngWidth As Long
  
  Set rsTemp = mclsData.OpenPersistentRecordset(strSQL, adOpenKeyset, adLockReadOnly)

  With cbo
    .Clear
    .AddItem "<All>"
    lngMaxWidth = Me.TextWidth(lbl.Caption)
    Do While Not rsTemp.EOF
      strTemp = Trim(rsTemp.Fields(0).Value)
      If strTemp <> "" Then
        .AddItem strTemp
        lngWidth = Me.TextWidth(strTemp) + 400
        If lngMaxWidth < lngWidth Then
          lngMaxWidth = lngWidth
        End If
      End If
      rsTemp.MoveNext
    Loop
    .ListIndex = 0
    .Width = lngMaxWidth
  End With

End Sub

