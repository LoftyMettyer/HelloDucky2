VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "actbar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAudit 
   Caption         =   "Audit"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8004
   Icon            =   "frmAudit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2745
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Begin SSDataWidgets_B.SSDBGrid grdAccess 
      Height          =   1995
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4020
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Active Row"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmAudit.frx":000C
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   3519
      Columns(0).Caption=   "Date / Time"
      Columns(0).Name =   "Date / Time"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2646
      Columns(1).Caption=   "User Group"
      Columns(1).Name =   "Group Name"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2963
      Columns(2).Caption=   "User"
      Columns(2).Name =   "User Name"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Computer Name"
      Columns(3).Name =   "Computer Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "OpenHR Module"
      Columns(4).Name =   "Module Name"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   4233
      Columns(5).Caption=   "Action"
      Columns(5).Name =   "Action"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "ID"
      Columns(6).Name =   "ID"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   7091
      _ExtentY        =   3519
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
   Begin MSComctlLib.StatusBar stbAudit 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2430
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7408
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "FILTER"
            TextSave        =   "FILTER"
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdPermission 
      Height          =   1995
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2880
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Active Row"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmAudit.frx":0028
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   8
      Columns(0).Width=   2963
      Columns(0).Caption=   "User"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Date / Time"
      Columns(1).Name =   "Date / Time"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2646
      Columns(2).Caption=   "User Group"
      Columns(2).Name =   "Group Name"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2831
      Columns(3).Caption=   "View / Table"
      Columns(3).Name =   "Table Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3201
      Columns(4).Caption=   "Column"
      Columns(4).Name =   "Column Name"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1773
      Columns(5).Caption=   "Action"
      Columns(5).Name =   "Action"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2646
      Columns(6).Caption=   "Permission"
      Columns(6).Name =   "Permission"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "ID"
      Columns(7).Name =   "ID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   5080
      _ExtentY        =   3519
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
   Begin SSDataWidgets_B.SSDBGrid grdRecords 
      Height          =   1995
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3270
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Active Row"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmAudit.frx":0044
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   2699
      Columns(0).Caption=   "User"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Date / Time"
      Columns(1).Name =   "Date / Time"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2646
      Columns(2).Caption=   "Table"
      Columns(2).Name =   "Table Name"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2646
      Columns(3).Caption=   "Column"
      Columns(3).Name =   "Column Name"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2646
      Columns(4).Caption=   "Old Value"
      Columns(4).Name =   "Old Value"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2646
      Columns(5).Caption=   "New Value"
      Columns(5).Name =   "New Value"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3175
      Columns(6).Caption=   "Record Description"
      Columns(6).Name =   "Record Description"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "ColumnID"
      Columns(7).Name =   "ColumnID"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "ID"
      Columns(8).Name =   "ID"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   5768
      _ExtentY        =   3519
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
   Begin SSDataWidgets_B.SSDBGrid grdGroup 
      Height          =   1995
      Left            =   15
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   4005
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   1
      stylesets(0).Name=   "Active Row"
      stylesets(0).ForeColor=   16777215
      stylesets(0).BackColor=   8388608
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "frmAudit.frx":0060
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      RowNavigation   =   1
      MaxSelectedRows =   0
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   2963
      Columns(0).Caption=   "User"
      Columns(0).Name =   "User Name"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3519
      Columns(1).Caption=   "Date / Time"
      Columns(1).Name =   "Date / Time"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2646
      Columns(2).Caption=   "User Group"
      Columns(2).Name =   "Group Name"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "User Login"
      Columns(3).Name =   "User Login"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Action"
      Columns(4).Name =   "Action"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "ID"
      Columns(5).Name =   "ID"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   7056
      _ExtentY        =   3528
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
   Begin ActiveBarLibraryCtl.ActiveBar abAudit 
      Left            =   4650
      Top             =   1290
      _ExtentX        =   847
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Bands           =   "frmAudit.frx":007C
   End
End
Attribute VB_Name = "frmAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public miAuditType As audType
Private mrsRecords As Recordset
Private msFilter As String
Private msSortOrder As String
Private mgrdAudit As SSDBGrid
Private mfrmOrder As frmAuditOrder
Private mabView() As Boolean

' To hold the filter criteria
Private mavFilterCriteria() As Variant
Private mblnReadOnly As Boolean

Const MaxTop = 100000

Private Sub ClearFilterArray()
  ReDim mavFilterCriteria(6, 0)
End Sub

Private Sub Initialise()
  
  Dim sRow As String
  Dim lCount As Long
  
  Screen.MousePointer = vbHourglass
    
  ' Purge the audit log before we show it to the user
  gADOCon.Execute "sp_AsrAuditLogPurge", , adExecuteNoRecords
  
  miAuditType = audRecords
      
  ' Set column visible array
  Set mgrdAudit = grdRecords
  ReDim mabView(mgrdAudit.Columns.Count)
  For lCount = 0 To mgrdAudit.Columns.Count
    mabView(lCount) = True
  Next
  
  ' Get the recordset of audit log entries.
  SetDefaultFilter
  CreateFilterCode
  ReloadRecords

  ' RH 12/10/00 - This is already done above (GetAllRecords) !?
  'ReloadRecords
  RefreshFormCaption
  Screen.MousePointer = vbDefault
  
  'PG 20120801 removed delete functionality
  abAudit.Bands("bndAudit").Tools("ID_AuditDelete").Visible = False
  frmMain.abSecurity.Bands("bndAudit").Tools("ID_AuditDelete").Visible = False
  
End Sub

Private Sub SetDefaultFilter()
  
  ReDim mavFilterCriteria(6, 1)
  mavFilterCriteria(1, 1) = "Date / Time"
  mavFilterCriteria(2, 1) = "is equal to or after"
  mavFilterCriteria(3, 1) = CStr(DateAdd("m", -3, Date))
  mavFilterCriteria(4, 1) = "DateTimeStamp"
  mavFilterCriteria(5, 1) = 11
  mavFilterCriteria(6, 1) = 11
  
End Sub

Public Property Get DefaultFilter() As Boolean
  If UBound(mavFilterCriteria, 2) = 1 Then
    If mavFilterCriteria(1, 1) = "Date / Time" And _
        mavFilterCriteria(2, 1) = "is equal to or after" And _
        mavFilterCriteria(3, 1) = CStr(DateAdd("m", -3, Date)) And _
        mavFilterCriteria(4, 1) = "DateTimeStamp" And _
        mavFilterCriteria(5, 1) = 11 And _
        mavFilterCriteria(6, 1) = 11 Then
      DefaultFilter = True
    End If
  End If
End Property

Private Sub abAudit_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  EditMenu Tool.Name
End Sub

Private Sub abAudit_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  ' Do not let the user modify the layout.
  Cancel = True
End Sub

Private Sub Form_Activate()
   
  RefreshAuditMenu
  
  With mgrdAudit
    .Redraw = False
    .ReBind
    .Rows = IIf(mrsRecords.RecordCount = -1, 0, mrsRecords.RecordCount)
    .Redraw = True
  End With
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  '# RH 26/08/99. To pass shortcut keys thru to the activebar control
  Dim fHandled As Boolean
  
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
  
  fHandled = frmMain.abSecurity.OnKeyDown(KeyCode, Shift)

  If fHandled Then
    KeyCode = 0
    Shift = 0
  End If
  
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF5 Then
    AuditRefresh
  End If
  
End Sub

Private Sub Form_Load()
    
  Dim iWindowState As Integer

  'MH20010823 Fault 1917 Only allow "sa" full access to audit log
  'mblnReadOnly = (Application.AccessMode <> accFull)
  mblnReadOnly = (Application.AccessMode <> accFull Or Not gbUserCanManageLogins)  '   LCase(gsUserName) <> "sa")

  gADOCon.Execute "EXEC sp_AsrAuditLogPurge"

  Initialise
  
  ' Get the last form size and state from the registry.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  Me.Height = GetPCSetting(Me.Name, "Height", Me.Height)
  Me.Width = GetPCSetting(Me.Name, "Width", Me.Width)
  iWindowState = GetPCSetting(Me.Name, "WindowState", vbNormal)
  Me.WindowState = IIf(iWindowState <> vbMinimized, iWindowState, vbNormal)
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'CloseConnection
  If Not mfrmOrder Is Nothing Then
    Unload mfrmOrder
    Set mfrmOrder = Nothing
  End If

  RefreshAuditMenu
  
End Sub

Private Sub Form_Resize()

  Dim lngHeight As Long

  If Me.WindowState <> vbMinimized Then
    mgrdAudit.Width = Me.ScaleWidth
    lngHeight = Me.ScaleHeight - stbAudit.Height
    If lngHeight < 0 Then lngHeight = 0
    mgrdAudit.Height = lngHeight
  End If

  ' Clear the icon off the caption bar
  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If

  frmMain.RefreshMenu False
  
End Sub

Private Sub SetViewColumns()

    Dim lCount As Long
    
    For lCount = 0 To mgrdAudit.Columns.Count - 2
        mgrdAudit.Columns(lCount).Visible = mabView(lCount)
    Next

End Sub

'Private Sub DeleteRecords()
'  Dim lCol As Long
'  Dim lCount As Long
'  Dim sSQL As String
'  Dim pintLoop As Integer
'  Dim pvarBookmark As Variant
'  Dim sTable As String
'  Dim bSetDeleteFlag As Boolean
'
'  If mgrdAudit.Rows = 0 Then
'    Exit Sub
'  End If
'
'  For lCol = 0 To mgrdAudit.Columns.Count
'    If UCase(mgrdAudit.Columns(lCol).Caption) = "ID" Then
'      Exit For
'    End If
'  Next
'
'  Select Case miAuditType
'    Case audRecords
'      sTable = "ASRSysAuditTrail"
'      bSetDeleteFlag = False
'    Case audPermissions
'      sTable = "ASRSysAuditPermissions"
'      bSetDeleteFlag = False
'    Case audGroups
'      sTable = "ASRSysAuditGroup"
'      bSetDeleteFlag = False
'    Case audAccess
'      sTable = "ASRSysAuditAccess"
'      bSetDeleteFlag = False
'  End Select
'
'  If Not Filtered Then
'
'    If bSetDeleteFlag Then
'      sSQL = "UPDATE " & sTable & " SET Deleted = 1"
'    Else
'      sSQL = "DELETE FROM " & sTable
'    End If
'
'    gADOCon.Execute sSQL, , adCmdText
'    ReloadRecords
'  Else
'
'  ' RH 14/11/00 - The latest attempt to get this to work !
'
'    ' JDM - 07/02/02 - Fault 3456 - Whoops...
'    If Not bSetDeleteFlag Then
'      sSQL = "DELETE FROM " & sTable & " WHERE ID IN ("
'    Else
'      sSQL = "UPDATE " & sTable & " SET Deleted = 1 WHERE ID IN ("
'    End If
'
'    mrsRecords.MoveFirst
'    Do Until mrsRecords.EOF
'        sSQL = sSQL & mrsRecords.Fields("ID").Value & ","
'        mrsRecords.MoveNext
'    Loop
'
'   sSQL = Left(sSQL, Len(sSQL) - 1) & ")"
'
'    gADOCon.Execute sSQL, , adCmdText
'    ReloadRecords
'
'    mgrdAudit.Redraw = True
'  End If
'
'  ClearFilterArray
'
'End Sub

Private Sub PrintRecords(plngCopies As Long, pfPortrait As Boolean, pfGrid As Boolean)

  If CheckPrinterIsOK Then
    With mgrdAudit
      .PageHeaderFont.Name = "Verdana"
      .PageHeaderFont.Size = 10
      .PageHeaderFont.Bold = False
      .PageHeaderFont.Underline = True
          
      .ActiveRowStyleSet = ""
      
      .PrintData ssPrintAllRows, False, False
'      .ActiveRowStyleSet = "Active Row"
      
      MsgBox "Audit Log printing complete." & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")", vbInformation, "Audit Log"

    End With
  End If
  
End Sub

Private Sub mgrdAudit_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  Dim sTitle As String

  With ssPrintInfo
    .PageHeader = vbTab & Me.Caption & vbTab
    .PageFooter = "Printed on <date> at <time> by " & gsUserName & vbTab & vbTab & "Page <page number>"
    .PrintColumnHeaders = ssUseCaption
    'NPG20080205 Fault 7879
    ' .PrintGridLines = frmAuditPrint.PrintGrid
    .PrintGridLines = IIf(frmPrintOptions.PrintGrid, 3, 0)
    'NPG20080205 Fault 7879
    ' .Portrait = frmAuditPrint.PrintPortrait
    .Portrait = frmPrintOptions.PrintPortrait
    'NPG20080205 Fault 7879
    ' .Copies = frmAuditPrint.PrintCopies
    .Copies = frmPrintOptions.PrintCopies
    .PrintHeaders = ssTopOfPage
    .MarginTop = 20
    .MarginBottom = 20
    .MarginLeft = 20
    .MarginRight = 20
  End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Save the form size and state to the registry.
  SavePCSetting Me.Name, "WindowState", Me.WindowState
  
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
    SavePCSetting Me.Name, "Height", Me.Height
    SavePCSetting Me.Name, "Width", Me.Width
  End If
  
  frmMain.RefreshMenu True
  
  Set mrsRecords = Nothing
  
End Sub


Public Sub EditMenu(psMenuItem As String)
  
  Dim iLoop As Integer
  
  ' Process the menu selection depending on what is currently selected.
  Select Case psMenuItem
    Case "ID_AuditOpen"
      OpenAuditLog
     
    Case "ID_Type1", "ID_Type2", "ID_Type3", "ID_Type4"
      
      If Filtered And Not DefaultFilter Then
        If MsgBox("Changing audit logs will remove the current filter." & vbCrLf & "Do you wish to continue ?", vbYesNo + vbQuestion, App.Title) = vbNo Then
          Exit Sub
        End If
      End If
           
      ' Clear the existing field selection so that all fields are displayed.
      For iLoop = 0 To mgrdAudit.Columns.Count
        mabView(iLoop) = True
        If iLoop < mgrdAudit.Columns.Count - 1 Then
          mgrdAudit.Columns(iLoop).Visible = mabView(iLoop) And (mgrdAudit.Columns(iLoop).Caption <> "ColumnID")
        End If
      Next
            
      miAuditType = CInt(Right(psMenuItem, 1))
      msSortOrder = ""
      SetDefaultFilter
      CreateFilterCode
      ReloadRecords
      RefreshFormCaption
    
    Case "ID_AuditDelete"
      'DeleteAuditLog
    
    Case "ID_AuditPrint"
      PrintAuditLog
    
    Case "ID_AuditShowColumns"
      ShowColumns
     
    Case "ID_AuditSort"
      sortOrder
     
    Case "ID_AuditSetFilter"
      SetFilter
      
    Case "ID_AuditClearFilter"
      ClearFilter
      
    Case "ID_AuditRefresh"
      AuditRefresh
      
    Case "ID_AuditScheduleTasks"
      ScheduleTasks
  End Select
  
  RefreshAuditMenu
  RefreshFormCaption

End Sub


Private Sub sortOrder()
  ' Sort the records.
  If mfrmOrder Is Nothing Then
    Set mfrmOrder = New frmAuditOrder
  End If
  
  With mfrmOrder
    .Initialise miAuditType
    .Show vbModal
    
    If Not .Cancelled Then
      msSortOrder = .sortOrder
      ReloadRecords
    End If
  End With
    
End Sub

Private Sub PrintAuditLog()
  ' Print the audit log.
  'NPG20080205 Fault 7879
  ' With frmAuditPrint
  With frmPrintOptions
    .Show vbModal
    
    If Not .Cancelled Then
      Screen.MousePointer = vbHourglass
      PrintRecords .PrintCopies, .PrintPortrait, .PrintGrid
    End If
  End With
  
  Screen.MousePointer = vbDefault

End Sub

Private Sub OpenAuditLog()
  ' Prompt the user to choose which Audit Log to open.
  Dim iLoop As Integer
  
  ' Display the audit selection form.
  With frmAuditOpen
    
    .Initialise miAuditType
    .Show vbModal
    
    If Not .Cancelled Then
      
      miAuditType = .OpenType
      
      abAudit.Tools("ID_Type1").Checked = False
      abAudit.Tools("ID_Type2").Checked = False
      abAudit.Tools("ID_Type3").Checked = False
      abAudit.Tools("ID_Type4").Checked = False
      
      Select Case miAuditType
        Case 1: abAudit.Tools("ID_Type1").Checked = True: EditMenu "ID_Type1"
        Case 2: abAudit.Tools("ID_Type2").Checked = True: EditMenu "ID_Type2"
        Case 3: abAudit.Tools("ID_Type3").Checked = True: EditMenu "ID_Type3"
        Case 4: abAudit.Tools("ID_Type4").Checked = True: EditMenu "ID_Type4"
      End Select
      
    End If
    
  End With
  Unload frmAuditOpen
  
  RefreshAuditMenu
  RefreshFormCaption
End Sub

'Private Sub DeleteAuditLog()
'  ' Delete the current Audit log records.
'  Dim sMsg As String
'
'  If Filtered Then
'   sMsg = "Are you sure you want to delete these filtered audit log records ?"
'  Else
'   sMsg = "Are you sure you want to delete all the currently visible audit log records ?"
'  End If
'
'  If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
'    Screen.MousePointer = vbHourglass
'    DeleteRecords
'    RefreshFormCaption
'    Screen.MousePointer = vbDefault
'  End If
'
'End Sub

Private Sub ShowColumns()
  ' Display the form that lets the user select which columns to show in the audit log.
  With frmAuditView
    .Initialise miAuditType, mabView()
    .Show vbModal
    
    If Not .Cancelled Then
      .GetDetails mabView()
      SetViewColumns
    End If
  End With
  
  Unload frmAuditView

End Sub

Private Sub SetFilter()
  ' Display the form that lets the user define a filter for the audit records.
  
    frmAuditFilter2.Initialise miAuditType, mavFilterCriteria
    frmAuditFilter2.Show vbModal
    If frmAuditFilter2.Cancelled Then
      Unload frmAuditFilter2
      Exit Sub
    End If
    
    ' Set local array to be the newly defined filter
    mavFilterCriteria = frmAuditFilter2.FilterArray
    CreateFilterCode
    ReloadRecords
    
    Unload frmAuditFilter2
    
    RefreshFormCaption
    RefreshAuditMenu
  
End Sub

Private Sub ClearFilter()
  ' Clear the defined filter.
  ClearFilterArray
  CreateFilterCode
  ReloadRecords
  RefreshFormCaption
  RefreshAuditMenu
End Sub

Private Sub ScheduleTasks()
  ' Display the form that allows the user to schedule tasks.
  Screen.MousePointer = vbHourglass
  With frmAuditCleardown
    .AuditType = miAuditType
    .Initialise
    Screen.MousePointer = vbDefault
    .Show vbModal
    If .Cancelled Then
      Exit Sub
    End If
  End With
  ReloadRecords
  RefreshFormCaption
End Sub

Private Sub RefreshAuditMenu()
  
  Dim blnEnabled As Boolean
  ' Refresh the toolbar on this form, and the menu bar on the MDI form.
  abAudit.Bands("bndAudit").Tools("ID_AuditClearFilter").Enabled = Filtered
  
  If mrsRecords.State <> 0 Then
    blnEnabled = (mrsRecords.RecordCount > 0)
    abAudit.Bands("bndAudit").Tools("ID_AuditPrint").Enabled = blnEnabled
    frmMain.abSecurity.Bands("bndAudit").Tools("ID_AuditPrint").Enabled = blnEnabled
    'PG 20120801 removed delete functionality
    'abAudit.Bands("bndAudit").Tools("ID_AuditDelete").Enabled = blnEnabled And Not mblnReadOnly
    'frmMain.abSecurity.Bands("bndAudit").Tools("ID_AuditDelete").Enabled = blnEnabled And Not mblnReadOnly
  End If
  
  frmMain.RefreshMenu False
  SetFormCaption Me, Me.Caption
  
End Sub

Private Sub grdAccess_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  mgrdAudit_PrintInitialize ssPrintInfo

End Sub

Private Sub grdGroup_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  mgrdAudit_PrintInitialize ssPrintInfo

End Sub


Private Sub grdPermission_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  mgrdAudit_PrintInitialize ssPrintInfo

End Sub

Private Sub grdRecords_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
  mgrdAudit_PrintInitialize ssPrintInfo

End Sub


Private Sub RefreshFormCaption()
  ' Set the form caption.
  Dim sCaption As String
  
  Select Case miAuditType
    Case audRecords
      sCaption = "Data Records Log"
    Case audPermissions
      sCaption = "Data Permissions Log"
    Case audGroups
      sCaption = "User Maintenance Log"
    Case audAccess
      sCaption = "User Access Audit Log"
  End Select

  If Filtered Then
    sCaption = sCaption & " [Filtered]"
  End If
   
  RefreshStatusBar
  SetFormCaption Me, sCaption

  If Me.WindowState = vbMaximized Then
    SetBlankIcon Me
  Else
    RemoveIcon Me
    Me.BorderStyle = vbSizable
  End If

End Sub

Private Sub RefreshStatusBar()
  ' Refresh the status bar.
  With stbAudit
    .Panels(1).Text = Me.Caption & " - " & IIf(mrsRecords.RecordCount = -1, 0, mrsRecords.RecordCount) & " records"
    .Panels(2).Enabled = Filtered
  End With
  
End Sub

Public Property Get Filtered() As Boolean
  ' Return TRUE if the Audit records are filtered.
  Filtered = (msFilter <> "")
End Property

Private Sub grdPermission_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  
  'JPD 20040624 Fault 4788
  iRowsRead = 0
  
  sDateFormat = "dd/mm/yyyy" 'DateFormat
  
  ' This is required as recordset not set when this sub is first run
  If mrsRecords Is Nothing Then Exit Sub
  
  ' Do nothing if there are no records to display
  If mrsRecords.RecordCount < 1 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsRecords.EOF Then
        mrsRecords.MoveLast
      End If
    Else
      If Not mrsRecords.BOF Then
        mrsRecords.MoveFirst
      End If
    End If
  Else
    mrsRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsRecords.BOF Or mrsRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsRecords.Fields.Count - 1)
          Select Case mrsRecords.Fields(iFieldIndex).Name
            Case "ID"
              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrsRecords.Fields("ID"))
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub

Private Sub grdPermission_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  'JPD 20040624 Fault 4788
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsRecords.MoveLast
    Else
      mrsRecords.MoveFirst
    End If
  Else
    mrsRecords.Bookmark = StartLocation
  End If
  
  mrsRecords.Move NumberOfRowsToMove
  NewLocation = mrsRecords.Bookmark

End Sub


Private Sub grdGroup_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  
  'JPD 20040624 Fault 4788
  iRowsRead = 0
  
  sDateFormat = "dd/mm/yyyy" 'DateFormat
  
  ' This is required as recordset not set when this sub is first run
  If mrsRecords Is Nothing Then Exit Sub
  
  ' Do nothing if there are no records to display
  If mrsRecords.RecordCount < 1 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsRecords.EOF Then
        mrsRecords.MoveLast
      End If
    Else
      If Not mrsRecords.BOF Then
        mrsRecords.MoveFirst
      End If
    End If
  Else
    mrsRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsRecords.BOF Or mrsRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsRecords.Fields.Count - 1)
          Select Case mrsRecords.Fields(iFieldIndex).Name
            Case "ID"
              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrsRecords.Fields("ID"))
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub

Private Sub grdGroup_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  'JPD 20040624 Fault 4788
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsRecords.MoveLast
    Else
      mrsRecords.MoveFirst
    End If
  Else
    mrsRecords.Bookmark = StartLocation
  End If
  
  mrsRecords.Move NumberOfRowsToMove
  NewLocation = mrsRecords.Bookmark

End Sub

Private Sub grdAccess_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  'JPD 20040624 Fault 4788
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsRecords.MoveLast
    Else
      mrsRecords.MoveFirst
    End If
  Else
    mrsRecords.Bookmark = StartLocation
  End If
  
  mrsRecords.Move NumberOfRowsToMove
  NewLocation = mrsRecords.Bookmark
  
End Sub

Private Sub grdRecords_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)

  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  
  'JPD 20040624 Fault 4788
  iRowsRead = 0
  
  sDateFormat = "dd/mm/yyyy" 'DateFormat
  
  ' This is required as recordset not set when this sub is first run
  If mrsRecords Is Nothing Then Exit Sub
  
  ' Do nothing if there are no records to display
  If mrsRecords.RecordCount < 1 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsRecords.EOF Then
        mrsRecords.MoveLast
      End If
    Else
      If Not mrsRecords.BOF Then
        mrsRecords.MoveFirst
      End If
    End If
  Else
    mrsRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsRecords.BOF Or mrsRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsRecords.Fields.Count - 1)
          Select Case mrsRecords.Fields(iFieldIndex).Name
            Case "ID"
              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrsRecords.Fields("ID"))
            Case "Old Value", "New Value"
              If mrsRecords.Fields("IsNumeric") Then
                If UI.GetSystemDecimalSeparator <> "." Then
                  RowBuf.Value(iRowIndex, iFieldIndex) = UI.ConvertNumberForDisplay(mrsRecords(iFieldIndex))
                Else
                 RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
                End If
              Else
                RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
              End If
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub

Private Sub grdRecords_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  'JPD 20040624 Fault 4788
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsRecords.MoveLast
    Else
      mrsRecords.MoveFirst
    End If
  Else
    mrsRecords.Bookmark = StartLocation
  End If
  
  mrsRecords.Move NumberOfRowsToMove
  NewLocation = mrsRecords.Bookmark

End Sub

Private Sub grdAccess_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  
  'JPD 20040624 Fault 4788
  iRowsRead = 0
  
  sDateFormat = "dd/mm/yyyy" 'DateFormat
  
  ' This is required as recordset not set when this sub is first run
  If mrsRecords Is Nothing Then Exit Sub
  
  ' Do nothing if there are no records to display
  If mrsRecords.RecordCount < 1 Then Exit Sub

  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsRecords.EOF Then
        mrsRecords.MoveLast
      End If
    Else
      If Not mrsRecords.BOF Then
        mrsRecords.MoveFirst
      End If
    End If
  Else
    mrsRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsRecords.BOF Or mrsRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsRecords.Fields.Count - 1)
          Select Case mrsRecords.Fields(iFieldIndex).Name
            Case "ID"
              RowBuf.Value(iRowIndex, iFieldIndex) = CStr(mrsRecords.Fields("ID"))
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsRecords(iFieldIndex)
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsRecords.Bookmark
    End Select
    
    If ReadPriorRows Then
      mrsRecords.MovePrevious
    Else
      mrsRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub

Private Function CheckPrinterIsOK() As Boolean

  Dim pstrError As String
  
  On Error GoTo PrintErr
  
  If Printer.CurrentX = 0 Then
  End If
  
  CheckPrinterIsOK = True
  Exit Function
  
PrintErr:
  
  CheckPrinterIsOK = False
  
  Select Case Err.Number
    Case 482: pstrError = "Printer Error : Please check your printer connection."
    Case Else: pstrError = Err.Description
  End Select
  
  MsgBox pstrError, vbExclamation + vbOKOnly, Application.Name

End Function

Private Sub ReloadRecords()
  Dim lCount As Long
  Dim sRow As String

  Screen.MousePointer = vbHourglass
  
  If Not mrsRecords Is Nothing Then
    If mrsRecords.State <> 0 Then
      mrsRecords.Close
    End If
    Set mrsRecords = Nothing
  End If
  Set mrsRecords = New Recordset
  
  ' Open the recordset initially with the most appropriate sort order. User may change
  ' this once the grid has been loaded.
  Select Case miAuditType
    Case 1: Set mrsRecords = GetAllRecords(miAuditType, IIf(msSortOrder <> "", GetFixedSortOrder(), "Order By a.DateTimeStamp DESC, a.TableName, a.ColumnName"), msFilter, MaxTop)
    Case 2: Set mrsRecords = GetAllRecords(miAuditType, IIf(msSortOrder <> "", GetFixedSortOrder(), "Order By a.DateTimeStamp DESC, a.ViewTableName, a.ColumnName, a.Action, a.Permission"), msFilter, MaxTop)
    Case 3: Set mrsRecords = GetAllRecords(miAuditType, IIf(msSortOrder <> "", GetFixedSortOrder(), "Order By a.DateTimeStamp DESC, a.GroupName, a.UserLogin, a.Action"), msFilter, MaxTop)
    Case 4: Set mrsRecords = GetAllRecords(miAuditType, IIf(msSortOrder <> "", GetFixedSortOrder(), "Order By a.DateTimeStamp DESC, a.UserGroup, a.UserName, a.HRProModule, a.Action"), msFilter, MaxTop)
  End Select
    
  Select Case miAuditType
    Case audRecords
      grdRecords.Visible = True
      grdPermission.Visible = False
      grdGroup.Visible = False
      grdAccess.Visible = False
      Me.HelpContextID = 8045
      Set mgrdAudit = grdRecords
    Case audPermissions
      grdRecords.Visible = False
      grdPermission.Visible = True
      grdGroup.Visible = False
      grdAccess.Visible = False
      Me.HelpContextID = 8046
      Set mgrdAudit = grdPermission
    Case audGroups
      grdRecords.Visible = False
      grdPermission.Visible = False
      grdAccess.Visible = False
      grdGroup.Visible = True
      Set mgrdAudit = grdGroup
      Me.HelpContextID = 8047
    Case audAccess
      grdRecords.Visible = False
      grdPermission.Visible = False
      grdGroup.Visible = False
      grdAccess.Visible = True
      Set mgrdAudit = grdAccess
      Me.HelpContextID = 8048
  End Select
  
  Form_Resize

  With mgrdAudit

    .Redraw = False
    .ReBind
    .Rows = IIf(mrsRecords.RecordCount = -1, 0, mrsRecords.RecordCount)
    .Redraw = True
  End With

  RefreshStatusBar
  Screen.MousePointer = vbDefault

  If mrsRecords.RecordCount = MaxTop Then
    MsgBox "The search results have been limited to " & Format$(MaxTop, "#,###") & " records.", vbInformation, App.Title
  End If
  
End Sub

Private Function GetFixedSortOrder() As String
  GetFixedSortOrder = msSortOrder
  GetFixedSortOrder = Replace(GetFixedSortOrder, "Order By ", "Order By a.")
  GetFixedSortOrder = Replace(GetFixedSortOrder, ", ", ", a.")
End Function


Private Sub AuditRefresh()
  ReloadRecords
  RefreshAuditMenu
End Sub

Private Function CreateFilterCode() As String
  ' Return the filter code for the recordset's defined filter.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim dtDateValue As Date
  Dim sCurrChar As String
  Dim sFilter As String
  Dim sFilterColName As String
  Dim sFilterValue As String
  Dim sModifiedFilterValue As String
  Dim sSubFilter As String
  
  Dim mavFilterExtra() As Variant
  
  ' NB. the filter array has 6 columns :
  ' Column 1 - the column display name.
  ' Column 2 - the operator display.
  ' Column 3 - the value.
  ' Column 4 - the column fieldname.
  ' Column 5 - the operator ID.
  ' Column 6 - the datatype
  
  On Error GoTo ErrTrap
  
  Const DateFormat = "yyyymmdd"
  
  ' Construct the filter string.
  sFilter = ""
  
  For iLoop = 1 To UBound(mavFilterCriteria, 2)
    
        sSubFilter = ""
        sFilterColName = "a.[" & mavFilterCriteria(4, iLoop) & "]"
        sFilterValue = mavFilterCriteria(3, iLoop)

        Select Case mavFilterCriteria(6, iLoop)
          
          Case sqlDate
             
            Select Case mavFilterCriteria(5, iLoop)
           
              Case giFILTEROP_ON ' Equal
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NULL"
                Else
                sSubFilter = sSubFilter & sFilterColName & " >= " & "'" & Format$(sFilterValue, DateFormat) & "'" & _
                      " AND " & sFilterColName & " < " & "'" & Format$(CDate(sFilterValue) + 1, DateFormat) & "'"
                End If
                
              Case giFILTEROP_NOTON ' Not equal to
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sSubFilter & sFilterColName & " < '" & Format$(sFilterValue, DateFormat) & "'" & _
                        " OR " & sFilterColName & " >= '" & Format$(CDate(sFilterValue) + 1, DateFormat) & "'"
                End If
                
              Case giFILTEROP_BEFORE ' Less than
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sSubFilter & sFilterColName & " <= '" & Format$(CDate(sFilterValue), DateFormat) & "'"
                End If
               
              Case giFILTEROP_AFTER ' greater than
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sSubFilter & sFilterColName & " >= '" & Format$(CDate(sFilterValue) + 1, DateFormat) & "'"
                End If
              
              'TM20011102 Fault 2155
              Case giFILTEROP_ONORAFTER ' greater than or equal to.
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sSubFilter & sFilterColName & " >= '" & Format$(CDate(sFilterValue), DateFormat) & "'"
                End If
           
              Case giFILTEROP_ONORBEFORE  ' less than or equal to.
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " IS NOT NULL"
                Else
                  sSubFilter = sSubFilter & sFilterColName & " <= '" & Format$(CDate(sFilterValue) + 1, DateFormat) & "'"
                End If
           
            End Select
 
          Case sqlVarchar
            
            Select Case mavFilterCriteria(5, iLoop)
              Case giFILTEROP_IS
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " = ''"
                Else
                  ' Replace the standard * and ? characters with the SQL % and _ characters.
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
'                      Case "*"
'                        sModifiedFilterValue = sModifiedFilterValue & "%"
'                      Case "?"
'                        sModifiedFilterValue = sModifiedFilterValue & "_"
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " = '" & sModifiedFilterValue & "'"
                End If
            
              Case giFILTEROP_ISNOT
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " <> ''"
                Else
                  ' Replace the standard * and ? characters with the SQL % and _ characters.
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
'                      Case "*"
'                        sModifiedFilterValue = sModifiedFilterValue & "%"
'                      Case "?"
'                        sModifiedFilterValue = sModifiedFilterValue & "_"
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " <> '" & sModifiedFilterValue & "'"
                End If

              Case giFILTEROP_CONTAINS
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " = ''"
                Else
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " LIKE '%" & sModifiedFilterValue & "%'"
                End If
            
              Case giFILTEROP_DOESNOTCONTAIN
                If Len(sFilterValue) = 0 Then
                  sSubFilter = sFilterColName & " <> ''"
                Else
                  sModifiedFilterValue = ""
                  For iLoop2 = 1 To Len(sFilterValue)
                    sCurrChar = Mid(sFilterValue, iLoop2, 1)
                    Select Case sCurrChar
                      Case "'"
                        sModifiedFilterValue = sModifiedFilterValue & "''"
                      Case Else
                        sModifiedFilterValue = sModifiedFilterValue & sCurrChar
                    End Select
                  Next iLoop2
                  sSubFilter = sFilterColName & " NOT LIKE '%" & sModifiedFilterValue & "%'"
                End If
            End Select
        End Select
    
    ' Add this filter criterion definition to the full global definition string.
    sFilter = sFilter & IIf(Len(sFilter) > 0, " AND (", "(") & sSubFilter & ")"
  Next iLoop
  
  msFilter = sFilter
  
  CreateFilterCode = True
  Exit Function
  
ErrTrap:
  
  MsgBox "Error in filter definition.", vbExclamation + vbOKOnly
  CreateFilterCode = False

End Function

