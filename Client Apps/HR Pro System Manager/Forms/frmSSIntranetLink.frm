VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmSSIntranetLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Self-service Intranet Link"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1041
   Icon            =   "frmSSIntranetLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEmailLink 
      Caption         =   "Email Link :"
      Height          =   1245
      Left            =   2880
      TabIndex        =   41
      Top             =   6360
      Width           =   6300
      Begin VB.TextBox txtEmailSubject 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   45
         Top             =   700
         Width           =   4515
      End
      Begin VB.TextBox txtEmailAddress 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   43
         Top             =   300
         Width           =   4515
      End
      Begin VB.Label lblEmailSubject 
         AutoSize        =   -1  'True
         Caption         =   "Email Subject :"
         Height          =   195
         Left            =   195
         TabIndex        =   44
         Top             =   765
         Width           =   1275
      End
      Begin VB.Label lblEMailAddress 
         AutoSize        =   -1  'True
         Caption         =   "Email Address :"
         Height          =   195
         Left            =   195
         TabIndex        =   42
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame fraURLLink 
      Caption         =   "URL :"
      Height          =   1125
      Left            =   2880
      TabIndex        =   31
      Top             =   5520
      Width           =   6300
      Begin VB.TextBox txtURL 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   33
         Top             =   300
         Width           =   4515
      End
      Begin VB.CheckBox chkNewWindow 
         Caption         =   "&Display in new window"
         Height          =   330
         Left            =   1575
         TabIndex        =   34
         Top             =   690
         Width           =   2685
      End
      Begin VB.Label lblURL 
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   32
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame fraDocument 
      Caption         =   "Document :"
      Height          =   1125
      Left            =   120
      TabIndex        =   46
      Top             =   9240
      Width           =   9060
      Begin VB.TextBox txtDocumentFilePath 
         Height          =   315
         Left            =   1400
         MaxLength       =   500
         TabIndex        =   48
         Top             =   300
         Width           =   7365
      End
      Begin VB.CheckBox chkDisplayDocumentHyperlink 
         Caption         =   "Displa&y hyperlink to document"
         Height          =   330
         Left            =   1395
         TabIndex        =   49
         Top             =   690
         Width           =   3720
      End
      Begin VB.Label lblDocumentFilePath 
         AutoSize        =   -1  'True
         Caption         =   "URL :"
         Height          =   195
         Left            =   195
         TabIndex        =   47
         Top             =   360
         Width           =   390
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraApplicationLink 
      Caption         =   "Application :"
      Height          =   1245
      Left            =   2880
      TabIndex        =   35
      Top             =   4200
      Width           =   6300
      Begin VB.CommandButton cmdAppFilePathSel 
         Height          =   315
         Left            =   5760
         Picture         =   "frmSSIntranetLink.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.TextBox txtAppFilePath 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   37
         Top             =   300
         Width           =   4185
      End
      Begin VB.TextBox txtAppParameters 
         Height          =   315
         Left            =   1575
         MaxLength       =   500
         TabIndex        =   40
         Top             =   700
         Width           =   4515
      End
      Begin VB.Label lblAppFilePath 
         AutoSize        =   -1  'True
         Caption         =   "File Path :"
         Height          =   195
         Left            =   195
         TabIndex        =   36
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lblAppParameters 
         AutoSize        =   -1  'True
         Caption         =   "Parameters :"
         Height          =   195
         Left            =   195
         TabIndex        =   39
         Top             =   765
         Width           =   930
      End
   End
   Begin VB.Frame fraLinkType 
      Caption         =   "Link Type :"
      Height          =   2535
      Left            =   150
      TabIndex        =   9
      Top             =   1920
      Width           =   2500
      Begin VB.OptionButton optLink 
         Caption         =   "&On-screen Document Display"
         Height          =   450
         Index           =   5
         Left            =   200
         TabIndex        =   15
         Top             =   2010
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Application"
         Height          =   315
         Index           =   4
         Left            =   200
         TabIndex        =   14
         Top             =   1700
         Width           =   1545
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&Email Link"
         Height          =   315
         Index           =   3
         Left            =   200
         TabIndex        =   13
         Top             =   1350
         Width           =   1305
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&HR Pro Screen"
         Height          =   315
         Index           =   0
         Left            =   200
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optLink 
         Caption         =   "&URL"
         Height          =   315
         Index           =   1
         Left            =   200
         TabIndex        =   12
         Top             =   1000
         Width           =   960
      End
      Begin VB.OptionButton optLink 
         Caption         =   "HR &Pro Report / Utility"
         Height          =   315
         Index           =   2
         Left            =   200
         TabIndex        =   11
         Top             =   650
         Width           =   2265
      End
   End
   Begin VB.Frame fraHRProScreenLink 
      Caption         =   "HR Pro Screen :"
      Height          =   2220
      Left            =   2880
      TabIndex        =   16
      Top             =   1920
      Width           =   6300
      Begin VB.TextBox txtPageTitle 
         Height          =   315
         Left            =   1575
         MaxLength       =   100
         TabIndex        =   22
         Top             =   1100
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProScreen 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   700
         Width           =   4515
      End
      Begin VB.ComboBox cboHRProTable 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   300
         Width           =   4515
      End
      Begin VB.ComboBox cboStartMode 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":015A
         Left            =   1575
         List            =   "frmSSIntranetLink.frx":015C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1500
         Width           =   4515
      End
      Begin VB.Label lblPageTitle 
         AutoSize        =   -1  'True
         Caption         =   "Page Title :"
         Height          =   195
         Left            =   200
         TabIndex        =   21
         Top             =   1160
         Width           =   810
      End
      Begin VB.Label lblHRProScreen 
         Caption         =   "Screen :"
         Height          =   195
         Left            =   195
         TabIndex        =   19
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblHRProTable 
         Caption         =   "Table :"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   360
         Width           =   810
      End
      Begin VB.Label lblStartMode 
         AutoSize        =   -1  'True
         Caption         =   "Start Mode :"
         Height          =   195
         Left            =   200
         TabIndex        =   23
         Top             =   1560
         Width           =   900
      End
   End
   Begin VB.Frame fraOKCancel 
      Height          =   400
      Left            =   6600
      TabIndex        =   50
      Top             =   10440
      Width           =   2600
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   400
         Left            =   135
         TabIndex        =   51
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1400
         TabIndex        =   52
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraLink 
      Caption         =   "Link :"
      Height          =   1710
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   9000
      Begin VB.ComboBox cboTableView 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":015E
         Left            =   1485
         List            =   "frmSSIntranetLink.frx":0160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1100
         Width           =   3030
      End
      Begin VB.TextBox txtText 
         Height          =   315
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   4
         Top             =   700
         Width           =   3030
      End
      Begin VB.TextBox txtPrompt 
         Height          =   315
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   2
         Top             =   300
         Width           =   3030
      End
      Begin SSDataWidgets_B.SSDBGrid grdAccess 
         Height          =   1230
         Left            =   5595
         TabIndex        =   8
         Top             =   300
         Width           =   3180
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         stylesets.count =   2
         stylesets(0).Name=   "SysSecMgr"
         stylesets(0).ForeColor=   -2147483631
         stylesets(0).BackColor=   -2147483633
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
         stylesets(0).Picture=   "frmSSIntranetLink.frx":0162
         stylesets(1).Name=   "ReadOnly"
         stylesets(1).ForeColor=   -2147483631
         stylesets(1).BackColor=   -2147483633
         stylesets(1).HasFont=   -1  'True
         BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         stylesets(1).Picture=   "frmSSIntranetLink.frx":017E
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
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
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   79
         Columns.Count   =   2
         Columns(0).Width=   3889
         Columns(0).Caption=   "User Group"
         Columns(0).Name =   "GroupName"
         Columns(0).AllowSizing=   0   'False
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   1244
         Columns(1).Caption=   "Visible"
         Columns(1).Name =   "Access"
         Columns(1).CaptionAlignment=   2
         Columns(1).AllowSizing=   0   'False
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   11
         Columns(1).FieldLen=   256
         Columns(1).Style=   2
         TabNavigation   =   1
         _ExtentX        =   5609
         _ExtentY        =   2170
         _StockProps     =   79
         BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblAccess 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visibility :"
         Height          =   195
         Left            =   4665
         TabIndex        =   7
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblTableView 
         Caption         =   "Table (View) :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   1155
         Width           =   1245
      End
      Begin VB.Label lblText 
         Caption         =   "Text :"
         Height          =   195
         Left            =   200
         TabIndex        =   3
         Top             =   760
         Width           =   615
      End
      Begin VB.Label lblPrompt 
         Caption         =   "Prompt :"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame fraHRProUtilityLink 
      Caption         =   "HR Pro Report / Utility :"
      Height          =   1485
      Left            =   2880
      TabIndex        =   25
      Top             =   7680
      Width           =   6300
      Begin VB.ComboBox cboHRProUtility 
         Height          =   315
         Left            =   1400
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   700
         Width           =   4700
      End
      Begin VB.ComboBox cboHRProUtilityType 
         Height          =   315
         ItemData        =   "frmSSIntranetLink.frx":019A
         Left            =   1400
         List            =   "frmSSIntranetLink.frx":019C
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   300
         Width           =   4700
      End
      Begin VB.Label lblHRProUtilityMessage 
         AutoSize        =   -1  'True
         Caption         =   "<message>"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1395
         TabIndex        =   30
         Top             =   1160
         Width           =   4695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHRProUtility 
         Caption         =   "Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   765
         Width           =   780
      End
      Begin VB.Label lblHRProUtilityType 
         Caption         =   "Type :"
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   360
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmSSIntranetLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum SSINTRANETSCREENTYPES
  SSINTLINKSCREEN_HRPRO = 0
  SSINTLINKSCREEN_URL = 1
  SSINTLINKSCREEN_UTILITY = 2
  'NPG20080125 Fault 12873
  SSINTLINKSCREEN_EMAIL = 3
  SSINTLINKSCREEN_APPLICATION = 4
  SSINTLINKSCREEN_DOCUMENT = 5
End Enum

Private mblnCancelled As Boolean
Private miLinkType As SSINTRANETLINKTYPES
Private mfChanged As Boolean
'Private mlngPersonnelTableID As Long
Private mblnRefreshing As Boolean
Private mlngTableID As Long
Private mlngViewID As Long
Private msTableViewName As String

Private mfNewWindow As Boolean

Private mblnReadOnly As Boolean

Private mcolSSITableViews As clsSSITableViews

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Private Sub FormatScreen()

  Const GAPBETWEENTEXTBOXES = 85
  Const GAPABOVEBUTTONS = 150
  Const GAPUNDERBUTTONS = 600
  Const LEFTGAP = 200
  Const GAPUNDERLASTCONTROL = 200
  Const GAPUNDERRADIOBUTTON = -15
  
  Select Case miLinkType
    Case SSINTLINK_BUTTON
      fraLink.Caption = "Button Link :"
    Case SSINTLINK_DROPDOWNLIST
      fraLink.Caption = "Dropdown List Link :"
    Case SSINTLINK_HYPERTEXT
      fraLink.Caption = "Hypertext Link :"
    Case SSINTLINK_DOCUMENT
      fraLink.Caption = "On-screen Document Display :"
  End Select
  
  ' Prompt only required for Button Links.
  lblPrompt.Visible = (miLinkType = SSINTLINK_BUTTON)
  txtPrompt.Visible = lblPrompt.Visible
  
  ' Reposition the Text controls if required.
  If (miLinkType <> SSINTLINK_BUTTON) Then
    lblText.Top = lblPrompt.Top
    txtText.Top = txtPrompt.Top
  End If
  
  cboTableView.Top = txtText.Top + txtText.Height + GAPBETWEENTEXTBOXES
  lblTableView.Top = cboTableView.Top + (lblText.Top - txtText.Top)
 
  ' HR Pro screen links only required for Button or Dropdown List Links
  optLink(0).Enabled = (miLinkType <> SSINTLINK_HYPERTEXT) And (miLinkType <> SSINTLINK_DOCUMENT)
  
  If miLinkType = SSINTLINK_DOCUMENT Then
    fraLinkType.Visible = False
    fraDocument.Visible = True
    fraDocument.Top = fraLinkType.Top
    fraDocument.Left = fraLinkType.Left
    fraDocument.Width = fraLink.Width
  Else
    With fraHRProScreenLink
      fraHRProUtilityLink.Top = .Top
      fraHRProUtilityLink.Height = .Height
      fraHRProUtilityLink.Left = .Left
      
      fraURLLink.Top = .Top
      fraURLLink.Height = .Height
      fraURLLink.Left = .Left
      
      fraEmailLink.Top = .Top
      fraEmailLink.Height = .Height
      fraEmailLink.Left = .Left
      
      fraApplicationLink.Top = .Top
      fraApplicationLink.Height = .Height
      fraApplicationLink.Left = .Left
    End With
  End If

  ' Position the OK/Cancel buttons
  If (miLinkType = SSINTLINK_DOCUMENT) Then
    fraOKCancel.Top = fraDocument.Top + fraDocument.Height + GAPABOVEBUTTONS
  Else
    fraOKCancel.Top = fraLinkType.Top + fraLinkType.Height + GAPABOVEBUTTONS
  End If
  
  ' Redimension the form.
  Me.Height = fraOKCancel.Top + fraOKCancel.Height + GAPUNDERBUTTONS
  
End Sub

Private Sub GetHRProUtilityTypes()

  ' Populate the Utility Types combo.
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  With cboHRProUtilityType
    .Clear
      
    .AddItem "Calendar Report"
    .ItemData(.NewIndex) = utlCalendarReport
        
    .AddItem "Custom Report"
    .ItemData(.NewIndex) = utlCustomReport
    
    .AddItem "Mail Merge"
    .ItemData(.NewIndex) = utlMailMerge
    
    If ASRDEVELOPMENT Or Application.WorkflowModule Then
      .AddItem "Workflow"
      .ItemData(.NewIndex) = utlWorkflow
    End If
    
    .ListIndex = iDefaultItem
  End With
 
End Sub

Private Sub GetHRProTables()

  ' Populate the tables combo.
  Dim sSQL As String
  Dim rsTables As dao.Recordset
  Dim iDefaultItem As Integer
  
  iDefaultItem = 0
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And (miLinkType <> SSINTLINK_DOCUMENT) Then
    cboHRProTable.Clear
      
    If mlngTableID > 0 Then
      ' Add the table and its children (not grand children).
      sSQL = "SELECT tmpTables.tableID, tmpTables.tableName" & _
        " FROM tmpTables" & _
        " WHERE (tmpTables.deleted = FALSE)" & _
        " AND ((tmpTables.tableID = " & CStr(mlngTableID) & ")" & _
        " OR (tmpTables.tableID IN (SELECT childID FROM tmpRelations WHERE parentID =" & CStr(mlngTableID) & ")))" & _
        " ORDER BY tmpTables.tableName"
      Set rsTables = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsTables.EOF
        cboHRProTable.AddItem rsTables!TableName
        cboHRProTable.ItemData(cboHRProTable.NewIndex) = rsTables!TableID
        
        If mlngTableID = rsTables!TableID Then
          iDefaultItem = cboHRProTable.NewIndex
        End If
        
        rsTables.MoveNext
      Wend
      rsTables.Close
      Set rsTables = Nothing
    End If
    
    If cboHRProTable.ListCount = 0 Then
      optLink(SSINTLINKSCREEN_UTILITY).value = True
      optLink(SSINTLINKSCREEN_HRPRO).Enabled = False
      GetHRProScreens
    Else
      cboHRProTable.ListIndex = iDefaultItem
    End If
  
    GetStartModes
  End If
  
End Sub

Private Sub GetHRProUtilities(pUtilityType As UtilityType)

  ' Populate the utilities combo.
  Dim sSQL As String
  Dim sWhereSQL As String
  Dim rsUtilities As New ADODB.Recordset
  Dim rsLocalUtilities As dao.Recordset
  Dim sTableName As String
  Dim sIDColumnName As String
  Dim fLocalTable As Boolean
  
  fLocalTable = False
  
  cboHRProUtility.Clear

  Select Case pUtilityType
    Case utlBatchJob
      sTableName = "ASRSysBatchJobName"
      sIDColumnName = "ID"
      
    Case utlCalendarReport
      sTableName = "ASRSysCalendarReports"
      sIDColumnName = "ID"
    
    Case utlCrossTab
      sTableName = "ASRSysCrossTab"
      sIDColumnName = "CrossTabID"
    
    Case utlCustomReport
      sTableName = "ASRSysCustomReportsName"
      sIDColumnName = "ID"
    
    Case utlDataTransfer
      sTableName = "ASRSysDataTransferName"
      sIDColumnName = "DataTransferID"
      
    Case utlExport
      sTableName = "ASRSysExportName"
      sIDColumnName = "ID"
      
    Case UtlGlobalAdd, utlGlobalDelete, utlGlobalUpdate
      sTableName = "ASRSysGlobalFunctions"
      sIDColumnName = "functionID"

    Case utlImport
      sTableName = "ASRSysImportName"
      sIDColumnName = "ID"

    Case utlLabel
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 1 "
    
    Case utlMailMerge
      sTableName = "ASRSysMailMergeName"
      sIDColumnName = "mailMergeID"
      sWhereSQL = "ASRSysMailMergeName.IsLabel = 0 "

    Case utlRecordProfile
      sTableName = "ASRSysRecordProfileName"
      sIDColumnName = "recordProfileID"
  
    Case utlMatchReport, utlSuccession, utlCareer
      sTableName = "ASRSysMatchReportName"
      sIDColumnName = "matchReportID"
  
    Case utlWorkflow
      fLocalTable = True
      sTableName = "tmpWorkflows"
      sIDColumnName = "ID"
      sWhereSQL = "tmpWorkflows.initiationType = " & CStr(WORKFLOWINITIATIONTYPE_MANUAL) & _
        " OR tmpWorkflows.initiationType is null"
  End Select
  
  If Len(sTableName) > 0 Then
    ' Get the available utilities of the given type.
    If fLocalTable Then
      sSQL = "SELECT " & sTableName & "." & sIDColumnName & " AS [ID], " & sTableName & ".name" & _
        " FROM " & sTableName & _
        " WHERE (" & sTableName & ".deleted = FALSE)"
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " AND (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        " ORDER BY " & sTableName & ".name"
      Set rsLocalUtilities = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsLocalUtilities.EOF
        cboHRProUtility.AddItem rsLocalUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsLocalUtilities!id

        rsLocalUtilities.MoveNext
      Wend
      rsLocalUtilities.Close
      Set rsLocalUtilities = Nothing
    Else
      sSQL = "SELECT" & _
        "  " & sTableName & ".name," & _
        "  " & sTableName & "." & sIDColumnName & " AS [ID]" & _
        "  FROM " & sTableName
      If Len(sWhereSQL) > 0 Then
        sSQL = sSQL & _
          " WHERE (" & sWhereSQL & ")"
      End If
      sSQL = sSQL & _
        "  ORDER BY " & sTableName & ".name"
      
      rsUtilities.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
      While Not rsUtilities.EOF
        cboHRProUtility.AddItem rsUtilities!Name
        cboHRProUtility.ItemData(cboHRProUtility.NewIndex) = rsUtilities!id
  
        rsUtilities.MoveNext
      Wend
      rsUtilities.Close
      Set rsUtilities = Nothing
    End If
  End If
  
  If cboHRProUtility.ListCount > 0 Then
    cboHRProUtility.ListIndex = 0
  End If

End Sub

Private Sub GetHRProScreens()

  ' Populate the screens combo.
  Dim sSQL As String
  Dim rsScreens As dao.Recordset

  If miLinkType <> SSINTLINK_HYPERTEXT Then
    cboHRProScreen.Clear

    If cboHRProTable.ListIndex >= 0 Then
      ' Add any SS Int screens for the seledcted table.
      sSQL = "SELECT tmpScreens.screenID, tmpScreens.name" & _
        " FROM tmpScreens" & _
        " WHERE (tmpScreens.deleted = FALSE)" & _
        " AND (tmpScreens.ssIntranet = TRUE)" & _
        " AND (tmpScreens.tableID = " & CStr(cboHRProTable.ItemData(cboHRProTable.ListIndex)) & ")" & _
        " ORDER BY tmpScreens.name"
      Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

      While Not rsScreens.EOF
        cboHRProScreen.AddItem rsScreens!Name
        cboHRProScreen.ItemData(cboHRProScreen.NewIndex) = rsScreens!ScreenID

        rsScreens.MoveNext
      Wend
      rsScreens.Close
      Set rsScreens = Nothing
    End If
  
    If cboHRProScreen.ListCount > 0 Then
      cboHRProScreen.ListIndex = 0
    End If
  End If
  
End Sub

Private Sub GetTablesViews()

  ' Populate the table(views) combo with the selected tables (views)
  ' (as passed in the ssi table views collection)
  
  Dim oSSITableView As clsSSITableView
  Dim iIndex As Integer
  Dim iLoop As Integer
  
  cboTableView.Clear
  
  For Each oSSITableView In mcolSSITableViews.Collection
    cboTableView.AddItem (oSSITableView.TableViewName)
  Next oSSITableView
  
  If cboTableView.ListCount > 0 Then
    iIndex = 0
    For iLoop = 0 To cboTableView.ListCount - 1
      If cboTableView.List(iLoop) = msTableViewName Then
        iIndex = iLoop
        Exit For
      End If
    Next iLoop
    cboTableView.ListIndex = iIndex
  End If
  
End Sub

Public Sub Initialize(piType As SSINTRANETLINKTYPES, _
                      psPrompt As String, _
                      psText As String, _
                      psHRProScreenID As String, _
                      psPageTitle As String, _
                      psURL As String, _
                      plngTableID As Long, _
                      psStartMode As String, _
                      plngViewID As Long, _
                      psUtilityType As String, _
                      psUtilityID As String, _
                      pfCopy As Boolean, _
                      psHiddenGroups As String, _
                      psTableViewName As String, _
                      pfNewWindow As Boolean, _
                      psEMailAddress As String, _
                      psEMailSubject As String, _
                      psAppFilePath As String, _
                      psAppParameters As String, _
                      psDocumentFilePath As String, _
                      pfDisplayDocumentHyperlink As Boolean, _
                      pfIsSeparator As Boolean, _
                      ByRef pcolSSITableViews As clsSSITableViews)
  
  Set mcolSSITableViews = pcolSSITableViews
  
  miLinkType = piType
  'mlngPersonnelTableID = plngPersonnelTableID
  mlngTableID = plngTableID
  mlngViewID = plngViewID
  msTableViewName = psTableViewName
  
  FormatScreen
  
  GetTablesViews
    
  'NPG20080128 Fault 12873     ' If Len(psURL) > 0 Then
  If Len(psURL) > 0 And Len(Trim(psEMailAddress)) = 0 Then
    optLink(SSINTLINKSCREEN_URL).value = True
  End If
  
  'NPG20080125 Fault 12873
  If Len(psEMailAddress) > 0 Then
    optLink(SSINTLINKSCREEN_EMAIL).value = True
  End If

  If Len(psAppFilePath) > 0 Then
    optLink(SSINTLINKSCREEN_APPLICATION).value = True
  End If
  
  If (Len(psDocumentFilePath) > 0) Or (miLinkType = SSINTLINK_DOCUMENT) Then
    optLink(SSINTLINKSCREEN_DOCUMENT).value = True
  End If
  
  If Len(psUtilityID) > 0 Then
    If CLng(psUtilityID) > 0 Then
      optLink(SSINTLINKSCREEN_UTILITY).value = True
    End If
  End If
  
  If miLinkType = SSINTLINK_HYPERTEXT And _
    Len(psURL) = 0 And _
    Len(psEMailAddress) = 0 And _
    Len(psAppFilePath) = 0 Then
    
    optLink(SSINTLINKSCREEN_UTILITY).value = True
    
  ElseIf miLinkType = SSINTLINK_DOCUMENT Then
    optLink(SSINTLINKSCREEN_DOCUMENT).value = True
  End If
  
  GetHRProTables
  GetHRProUtilityTypes
  UtilityType = psUtilityType

  Prompt = psPrompt
  Text = psText
  HRProScreenID = psHRProScreenID
  PageTitle = psPageTitle
  'NPG20080125 Fault 12873
  EMailAddress = psEMailAddress
  EMailSubject = psEMailSubject
  URL = psURL

  StartMode = psStartMode
  UtilityType = psUtilityType
  UtilityID = psUtilityID
  NewWindow = pfNewWindow
  
  AppFilePath = psAppFilePath
  AppParameters = psAppParameters
  
  DocumentFilePath = psDocumentFilePath
  DisplayDocumentHyperlink = pfDisplayDocumentHyperlink

  PopulateAccessGrid psHiddenGroups

  mfChanged = False
  If pfCopy Then mfChanged = True
  RefreshControls
  
End Sub

Private Sub RefreshControls()

  Dim sUtilityMessage As String
  
  If mblnRefreshing Then Exit Sub
  
  sUtilityMessage = ""
  
  fraHRProScreenLink.Visible = optLink(SSINTLINKSCREEN_HRPRO).value
  fraHRProUtilityLink.Visible = optLink(SSINTLINKSCREEN_UTILITY).value
  fraURLLink.Visible = optLink(SSINTLINKSCREEN_URL).value
  fraEmailLink.Visible = optLink(SSINTLINKSCREEN_EMAIL).value
  fraApplicationLink.Visible = optLink(SSINTLINKSCREEN_APPLICATION).value
  fraDocument.Visible = optLink(SSINTLINKSCREEN_DOCUMENT).value
  
  ' Disable the HR Pro screen controls as required.
  cboHRProTable.Enabled = (optLink(SSINTLINKSCREEN_HRPRO).value) And (cboHRProTable.ListCount > 0)
  cboHRProTable.BackColor = IIf(cboHRProTable.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProTable.Enabled = cboHRProTable.Enabled
  cboHRProScreen.Enabled = (optLink(SSINTLINKSCREEN_HRPRO).value) And (cboHRProScreen.ListCount > 0)
  cboHRProScreen.BackColor = IIf(cboHRProScreen.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProScreen.Enabled = cboHRProTable.Enabled
  
  txtPageTitle.Enabled = cboHRProTable.Enabled
  txtPageTitle.BackColor = cboHRProTable.BackColor
  lblPageTitle.Enabled = cboHRProTable.Enabled
  If Not optLink(SSINTLINKSCREEN_HRPRO).value Then
    cboHRProTable.Clear
    cboHRProScreen.Clear
    cboStartMode.Clear
    txtPageTitle.Text = ""
  End If
  
  cboStartMode.Enabled = (cboStartMode.ListCount > 1) _
      And (optLink(SSINTLINKSCREEN_HRPRO).value)
  cboStartMode.BackColor = IIf(cboStartMode.Enabled, vbWindowBackground, vbButtonFace)
  lblStartMode.Enabled = cboHRProTable.Enabled
  
  ' Disable the UTILITY controls as required.
  If Not optLink(SSINTLINKSCREEN_UTILITY).value Then
    cboHRProUtilityType.Clear
    cboHRProUtility.Clear
  Else
    ' For Workflows only, check if the selected Workflow is enabled.
    ' Display a message as required.
    'JPD 20060714 Fault 11226
    If cboHRProUtilityType.ListIndex >= 0 Then
      If cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex) = utlWorkflow Then
        If cboHRProUtility.ListCount > 0 Then
          recWorkflowEdit.Index = "idxWorkflowID"
          recWorkflowEdit.Seek "=", cboHRProUtility.ItemData(cboHRProUtility.ListIndex)
    
          If Not recWorkflowEdit.NoMatch Then
            If (Not recWorkflowEdit.Fields("enabled").value) Then
              sUtilityMessage = "This Workflow is not currently enabled."
            End If
          End If
        End If
      End If
    End If
  End If
  
  cboHRProUtilityType.Enabled = optLink(SSINTLINKSCREEN_UTILITY).value
  cboHRProUtilityType.BackColor = IIf(cboHRProUtilityType.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProUtilityType.Enabled = cboHRProUtilityType.Enabled
  
  cboHRProUtility.Enabled = optLink(SSINTLINKSCREEN_UTILITY).value And (cboHRProUtility.ListCount > 0)
  cboHRProUtility.BackColor = IIf(cboHRProUtility.Enabled, vbWindowBackground, vbButtonFace)
  lblHRProUtility.Enabled = cboHRProUtility.Enabled
  
  ' Disable the URL controls as required.
  txtURL.Enabled = optLink(SSINTLINKSCREEN_URL).value
  txtURL.BackColor = IIf(txtURL.Enabled, vbWindowBackground, vbButtonFace)
  lblURL.Enabled = txtURL.Enabled
  chkNewWindow.Enabled = txtURL.Enabled
  ' 'NPG20080128 Fault 12873 - If Not txtURL.Enabled Then
  If Not txtURL.Enabled And Not optLink(SSINTLINKSCREEN_EMAIL).value Then
    txtURL.Text = ""
  End If
  
  'NPG20080125 Fault 12873
  ' Disable the EMail controls as required.
  txtEmailAddress.Enabled = optLink(SSINTLINKSCREEN_EMAIL).value
  txtEmailAddress.BackColor = IIf(txtEmailAddress.Enabled, vbWindowBackground, vbButtonFace)
  lblEMailAddress.Enabled = txtEmailAddress.Enabled
  txtEmailSubject.Enabled = optLink(SSINTLINKSCREEN_EMAIL).value
  txtEmailSubject.BackColor = IIf(txtEmailAddress.Enabled, vbWindowBackground, vbButtonFace)
  lblEmailSubject.Enabled = txtEmailAddress.Enabled
  If Not txtEmailAddress.Enabled Then
    txtEmailAddress.Text = ""
    txtEmailSubject.Text = ""
  End If

  ' Disable the Application link controls as required.
  txtAppFilePath.Enabled = optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppFilePath.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppFilePath.Enabled = txtAppFilePath.Enabled
  txtAppParameters.Enabled = optLink(SSINTLINKSCREEN_APPLICATION).value
  txtAppParameters.BackColor = IIf(txtAppFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblAppParameters.Enabled = txtAppFilePath.Enabled
  If Not txtAppFilePath.Enabled Then
    txtAppFilePath.Text = ""
    txtAppParameters.Text = ""
  End If

  ' Disable the Report Link controls as required.
  txtDocumentFilePath.Enabled = optLink(SSINTLINKSCREEN_DOCUMENT).value
  txtDocumentFilePath.BackColor = IIf(txtDocumentFilePath.Enabled, vbWindowBackground, vbButtonFace)
  lblDocumentFilePath.Enabled = txtDocumentFilePath.Enabled
  chkDisplayDocumentHyperlink.Enabled = txtDocumentFilePath.Enabled
  If Not txtDocumentFilePath.Enabled Then
    txtDocumentFilePath.Text = ""
    chkDisplayDocumentHyperlink.value = ssCBUnchecked
  End If
  
  mblnRefreshing = True
  GetStartModes
  mblnRefreshing = False
  
  lblHRProUtilityMessage.Caption = sUtilityMessage
  
  ' Disable the OK button as required.
  cmdOk.Enabled = mfChanged
  
End Sub

Private Sub GetStartModes()

  Dim iIndex As Integer
  Dim iOriginalStartMode As Integer
  Dim fPersonnelLink As Boolean
  
  iOriginalStartMode = 0
  iIndex = 0
  
  If (cboHRProTable.ListIndex >= 0) _
    And (cboTableView.ListIndex >= 0) Then
    fPersonnelLink = (cboHRProTable.ItemData(cboHRProTable.ListIndex) = mlngTableID)
  End If
  
  With cboStartMode
    If .ListIndex >= 0 Then
      iOriginalStartMode = .ItemData(.ListIndex)
    End If
    
    .Clear
    
    If optLink(SSINTLINKSCREEN_HRPRO).value Then
      If Not fPersonnelLink Then
        .AddItem "Find Window"
        .ItemData(.NewIndex) = 3
        If iOriginalStartMode = 3 Then
          iIndex = .NewIndex
        End If
      End If
      
      .AddItem "First Record"
      .ItemData(.NewIndex) = 2
      If iOriginalStartMode = 2 Then
        iIndex = .NewIndex
      End If
      
      If Not fPersonnelLink Then
        .AddItem "New Record"
        .ItemData(.NewIndex) = 1
        If iOriginalStartMode = 1 Then
          iIndex = .NewIndex
        End If
      End If
      
      .ListIndex = iIndex
    End If
    
    .Enabled = (.ListCount > 1) _
      And (optLink(SSINTLINKSCREEN_HRPRO).value)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With
  
End Sub

Private Function ValidateLink() As Boolean

  ' Return FALSE if the link definition is invalid.
  Dim fValid As Boolean
  
  fValid = True
  
  ' Check that a prompt has been entered (if required)
  'JPD 20070424 Fault 12168
  'If (miLinkType = SSINTLINK_BUTTON) And _
  '  (Len(txtPrompt.Text) = 0) Then
  '  fValid = False
  '  MsgBox "No prompt has been entered.", vbOKOnly + vbExclamation, Application.Name
  '  txtPrompt.SetFocus
  'End If
  
  ' Check that text has been entered
  If fValid Then
    If (Len(txtText.Text) = 0) Then
      fValid = False
      MsgBox "No text has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtText.SetFocus
    End If
  End If
  
  ' Check that the HR Pro screen has been selected (if required)
  If fValid Then
    If (miLinkType <> SSINTLINK_HYPERTEXT) And _
      (optLink(SSINTLINKSCREEN_HRPRO).value) Then
      If cboHRProScreen.ListIndex < 0 Then
        fValid = False
        MsgBox "No HR Pro screen has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProTable.SetFocus
      End If
    End If
  End If
  
  ' Check that the HR Pro page title been entered
  If fValid Then
    If (miLinkType <> SSINTLINK_HYPERTEXT) And _
      (optLink(SSINTLINKSCREEN_HRPRO).value) Then
      If (Len(txtPageTitle.Text) = 0) Then
        fValid = False
        MsgBox "No page title has been entered.", vbOKOnly + vbExclamation, Application.Name
        txtPageTitle.SetFocus
      End If
    End If
  End If
  
  ' Check that a URL has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_URL).value And _
      (Len(txtURL.Text) = 0) Then
      fValid = False
      MsgBox "No URL has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtURL.SetFocus
    End If
  End If
    
  'NPG20080125 Fault 12873
  ' Check that an Email Address has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_EMAIL).value And _
      ((Len(txtEmailAddress.Text) = 0) Or InStr(1, txtEmailAddress.Text, "@", 1) = 0) Then
      fValid = False
      MsgBox "No Email Address has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtEmailAddress.SetFocus
    End If
  End If

  ' Check that a utility has been entered (if required)
  If fValid Then
    If (optLink(SSINTLINKSCREEN_UTILITY).value) Then
      If cboHRProUtility.ListIndex < 0 Then
        fValid = False
        MsgBox "No " & cboHRProUtilityType.List(cboHRProUtilityType.ListIndex) & " has been selected.", vbOKOnly + vbExclamation, Application.Name
        cboHRProUtilityType.SetFocus
      End If
    End If
  End If

  ' Check that an Application File Path has been entered (if required)
  If fValid Then
    If optLink(SSINTLINKSCREEN_APPLICATION).value And _
      (Len(txtAppFilePath.Text) = 0) Then
      fValid = False
      MsgBox "No application file path has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtAppFilePath.SetFocus
    End If
  End If

'TM20090224 - Fault 13557, remove the restriction to allow links to shortcuts and other executable types.
'  If fValid Then
'    If optLink(SSINTLINKSCREEN_APPLICATION).Value And _
'      (LCase(Right(txtAppFilePath.Text, 4)) <> ".exe") Then
'      fValid = False
'      MsgBox "Please enter a valid executable file path.", vbOKOnly + vbExclamation, Application.Name
'      txtAppFilePath.SetFocus
'    End If
'  End If

  If fValid Then
    If optLink(SSINTLINKSCREEN_DOCUMENT).value And _
      (Len(txtDocumentFilePath.Text) = 0) Then
      fValid = False
      MsgBox "No Document URL has been entered.", vbOKOnly + vbExclamation, Application.Name
      txtDocumentFilePath.SetFocus
    End If
  End If

  If fValid Then
    If optLink(SSINTLINKSCREEN_DOCUMENT).value And _
      ((LCase(Left(txtDocumentFilePath.Text, 7)) <> "http://") And (LCase(Left(txtDocumentFilePath.Text, 8)) <> "https://")) Then
      fValid = False
      MsgBox "Please enter a valid URL path.", vbOKOnly + vbExclamation, Application.Name
      txtDocumentFilePath.SetFocus
    End If
  End If

  ValidateLink = fValid
  
End Function

Private Sub cboHRProScreen_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProTable_Click()
  GetHRProScreens
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProUtility_Click()
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboHRProUtilityType_Click()
  GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  mfChanged = True
  RefreshControls
End Sub

Private Sub cboStartMode_Click()

  'JPD 20050810 Fault 10241
  If Not mblnRefreshing Then
    mfChanged = True
    RefreshControls
  End If
  
End Sub

Private Sub cboTableView_Click()
  
  'TM20090512 - Fault 13680
  'Only refresh the table combo if the base table changes.
  
  Dim lngNewTableID As Long
  
  msTableViewName = cboTableView.List(cboTableView.ListIndex)
  lngNewTableID = GetTableIDFromCollection(mcolSSITableViews, msTableViewName)
  
  If mlngTableID <> lngNewTableID Then
    mlngTableID = lngNewTableID
    mlngViewID = GetViewIDFromCollection(mcolSSITableViews, msTableViewName)
  
    GetHRProTables
'TM20090512 - Fault 13680
'    GetHRProUtilityTypes
  Else
    mlngTableID = lngNewTableID
    mlngViewID = GetViewIDFromCollection(mcolSSITableViews, msTableViewName)
    
  End If
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub chkDisplayDocumentHyperlink_Click()

  Dim fValid As Boolean
  
  fValid = True

  If optLink(SSINTLINKSCREEN_DOCUMENT).value And (Len(txtDocumentFilePath.Text) = 0) And chkDisplayDocumentHyperlink.value Then
    fValid = False
    MsgBox "No Document URL has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtDocumentFilePath.SetFocus
  End If
  
  If fValid Then
    mfChanged = True
    RefreshControls
  End If

End Sub

Private Sub chkNewWindow_Click()

  Dim fValid As Boolean
  
  fValid = True

  If optLink(SSINTLINKSCREEN_URL).value And (Len(txtURL.Text) = 0) And chkNewWindow.value Then
    fValid = False
    MsgBox "No URL has been entered.", vbOKOnly + vbExclamation, Application.Name
    txtURL.SetFocus
  End If
  
  If fValid Then
    mfChanged = True
    RefreshControls
  End If
  
End Sub

Private Sub cmdAppFilePathSel_Click()

  On Local Error GoTo LocalErr
  
  With CommonDialog1

    .FileName = txtAppFilePath.Text
    If txtAppFilePath.Text = vbNullString Then
      .InitDir = "c:\"
    End If

    .CancelError = True
    .DialogTitle = "Application file path"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    
    .ShowOpen
    
    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        MsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtAppFilePath.Text = .FileName
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    MsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  UnLoad Me
End Sub

Private Sub cmdOK_Click()

  If ValidateLink Then
    Cancelled = False
    Me.Hide
  End If

End Sub

Private Sub cmdReportOutputFilePathSel_Click()

  On Local Error GoTo LocalErr
  
  With CommonDialog1

    .FileName = txtDocumentFilePath.Text
    If txtDocumentFilePath.Text = vbNullString Then
      .InitDir = "c:\"
    End If

    .CancelError = True
    .DialogTitle = "Report Output file path"
    .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
    
    .ShowOpen
    
    If .FileName <> vbNullString Then
      If Len(.FileName) > 255 Then
        MsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
      Else
        txtDocumentFilePath.Text = .FileName
      End If
    End If

  End With

Exit Sub

LocalErr:
  If Err.Number <> 32755 Then   '32755 = Cancel was selected.
    MsgBox Err.Description, vbCritical
  End If

End Sub

Private Sub Form_Initialize()
  mblnReadOnly = (Application.AccessMode <> accFull And _
    Application.AccessMode <> accSupportMode)
End Sub

Private Sub Form_Load()

  ' Position the form in the same place it was last time.
  Me.Top = GetPCSetting(Me.Name, "Top", (Screen.Height - Me.Height) / 2)
  Me.Left = GetPCSetting(Me.Name, "Left", (Screen.Width - Me.Width) / 2)
  
  fraOKCancel.BorderStyle = vbBSNone

  grdAccess.RowHeight = 239

End Sub

Private Sub PopulateAccessGrid(psHiddenGroups As String)

  ' Populate the access grid.
  Dim sSQL As String
  Dim rsGroups As New ADODB.Recordset
  Dim sVisibility As String
  Dim fAllVisible As Boolean
  
  fAllVisible = True

  ' Add the 'All Groups' item.
  With grdAccess
    .RemoveAll
    .AddItem "(All Groups)"
  End With

  ' Get the recordset of user groups and their access on this definition.
  sSQL = "SELECT name FROM sysusers" & _
    " WHERE gid = uid AND gid > 0" & _
    "   AND not (name like 'ASRSys%') AND not (name like 'db[_]%')" & _
    " ORDER BY name"
  rsGroups.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly

  With rsGroups
    Do While Not .EOF
      ' Add the user groups and their access on this definition to the access grid.
      If InStr(vbTab & UCase(psHiddenGroups) & vbTab, vbTab & UCase(Trim(!Name)) & vbTab) > 0 Then
        sVisibility = "False"
        fAllVisible = False
      Else
        sVisibility = "True"
      End If
            
      grdAccess.AddItem Trim(!Name) & vbTab & sVisibility

      .MoveNext
    Loop

    .Close
  End With
  Set rsGroups = Nothing

  With grdAccess
    .MoveFirst
    .Columns("Access").value = fAllVisible
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Cancelled = True
  
  If (UnloadMode <> vbFormCode) And mfChanged Then
    Select Case MsgBox("Apply changes ?", vbYesNoCancel + vbQuestion, Me.Caption)
      Case vbCancel
        Cancel = True
      Case vbYes
        cmdOK_Click
        Cancel = True   'MH20021105 Fault 4694
    End Select
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication
End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Save the form position to registry.
  If Me.WindowState = vbNormal Then
    SavePCSetting Me.Name, "Top", Me.Top
    SavePCSetting Me.Name, "Left", Me.Left
  End If

End Sub

Private Sub grdAccess_Change()

  Dim iLoop As Integer
  Dim varFirstRow As Variant
  Dim varCurrentRow As Variant
  Dim fNewValue As Boolean
  Dim fAllVisible As Boolean
  
  UI.LockWindow grdAccess.hWnd
  
  If (grdAccess.AddItemRowIndex(grdAccess.Bookmark) = 0) Then
    ' The 'All Groups' access has changed. Apply the selection to all other groups.
    With grdAccess
      .MoveFirst

      For iLoop = 0 To (.Rows - 1)
        If iLoop = 0 Then
          fNewValue = .Columns("Access").value
        Else
          .Columns("Access").value = fNewValue
        End If

        .MoveNext
      Next iLoop

      .MoveFirst
    End With
  Else
    fAllVisible = True
    
    With grdAccess
      varFirstRow = .FirstRow
      varCurrentRow = .Bookmark
      .MoveLast
      
      For iLoop = (.Rows - 1) To 0 Step -1
        If iLoop = 0 Then
          .Columns("Access").value = fAllVisible
        Else
          If Not .Columns("Access").value Then fAllVisible = False
        End If
        
        .MovePrevious
      Next iLoop

      .MoveFirst
    
      .FirstRow = varFirstRow
      .Bookmark = varCurrentRow
    End With
  End If
    
  UI.UnlockWindow

  grdAccess.col = 1

  mfChanged = True
  RefreshControls

End Sub

Private Sub optLink_Click(Index As Integer)

  GetHRProTables
  GetHRProUtilityTypes
  UtilityType = CStr(utlCalendarReport)
  
  mfChanged = True
  RefreshControls

End Sub

Private Sub txtDocumentFilePath_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtDocumentFilePath_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtAppFilePath_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtAppFilePath_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtAppParameters_Change()
  mfChanged = True
  RefreshControls
End Sub
Private Sub txtAppParameters_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtEmailAddress_Change()
  If Len(txtEmailAddress.Text) > 0 Then
    txtURL.Text = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  End If
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtEmailAddress_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtEmailSubject_Change()
  If Len(txtEmailAddress.Text) > 0 Then
    txtURL.Text = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  End If
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtEmailSubject_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtPageTitle_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtPageTitle_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtPrompt_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtPrompt_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtText_Change()
  mfChanged = True
  RefreshControls
End Sub

Private Sub txtText_GotFocus()
  UI.txtSelText
End Sub

Private Sub txtURL_Change()
  mfChanged = True
  RefreshControls
End Sub

Public Property Get Text() As String
  Text = txtText.Text
End Property

Public Property Let Text(ByVal psNewValue As String)
  txtText.Text = psNewValue
End Property

Public Property Get NewWindow() As Boolean
  NewWindow = chkNewWindow.value
End Property

Public Property Let NewWindow(ByVal pfNewValue As Boolean)
  chkNewWindow.value = IIf(pfNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get HiddenGroups() As String

  Dim iLoop As Integer
  Dim sHiddenGroups As String
  
  sHiddenGroups = ""
  
  With grdAccess
    For iLoop = 1 To (.Rows - 1)
      .Bookmark = .AddItemBookmark(iLoop)
      If Not .Columns("Access").value Then
        sHiddenGroups = sHiddenGroups & .Columns("GroupName").Text & vbTab
      End If
    Next iLoop
    
    .MoveFirst
  End With

  If Len(sHiddenGroups) > 0 Then
    sHiddenGroups = vbTab & sHiddenGroups
  End If
  
  HiddenGroups = sHiddenGroups
  
End Property

Public Property Get URL() As String
  URL = IIf(optLink(SSINTLINKSCREEN_URL).value Or optLink(SSINTLINKSCREEN_EMAIL).value, txtURL.Text, "")
End Property

Public Property Get EMailAddress() As String
  If optLink(SSINTLINKSCREEN_EMAIL).value Then
    EMailAddress = txtEmailAddress.Text
    URL = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  Else
    EMailAddress = ""
    URL = ""
  End If
End Property

Public Property Get EMailSubject() As String
  If optLink(SSINTLINKSCREEN_EMAIL).value Then
    EMailSubject = txtEmailSubject.Text
    URL = "mailto:" & Trim(LCase(txtEmailAddress.Text)) & "?Subject=" & Trim(txtEmailSubject.Text)
  Else
    EMailSubject = ""
    URL = ""
  End If
End Property

Public Property Get AppFilePath() As String
  If optLink(SSINTLINKSCREEN_APPLICATION).value Then
    AppFilePath = txtAppFilePath.Text
  Else
    AppFilePath = ""
  End If
End Property

Public Property Get AppParameters() As String
  If optLink(SSINTLINKSCREEN_APPLICATION).value Then
    AppParameters = txtAppParameters.Text
  Else
    AppParameters = ""
  End If
End Property

Public Property Get DocumentFilePath() As String
  If optLink(SSINTLINKSCREEN_DOCUMENT).value Then
    DocumentFilePath = txtDocumentFilePath.Text
  Else
    DocumentFilePath = ""
  End If
End Property

Public Property Get DisplayDocumentHyperlink() As Boolean
  If optLink(SSINTLINKSCREEN_DOCUMENT).value Then
    DisplayDocumentHyperlink = chkDisplayDocumentHyperlink.value
  Else
    DisplayDocumentHyperlink = False
  End If
End Property

Public Property Get UtilityType() As String

  If (cboHRProUtility.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_UTILITY).value) Then
    UtilityType = ""
  Else
    UtilityType = CStr(cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex))
  End If

End Property

Public Property Get UtilityID() As String

  If (cboHRProUtility.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_UTILITY).value) Then
    
    UtilityID = ""
  Else
    UtilityID = CStr(cboHRProUtility.ItemData(cboHRProUtility.ListIndex))
  End If

End Property

Public Property Let URL(ByVal psNewValue As String)
  txtURL.Text = IIf(optLink(SSINTLINKSCREEN_URL).value Or optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property


Public Property Let EMailAddress(ByVal psNewValue As String)
  txtEmailAddress.Text = IIf(optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property

Public Property Let EMailSubject(ByVal psNewValue As String)
  txtEmailSubject.Text = IIf(optLink(SSINTLINKSCREEN_EMAIL).value, psNewValue, "")
End Property

Public Property Let AppFilePath(ByVal psNewValue As String)
  txtAppFilePath.Text = IIf(optLink(SSINTLINKSCREEN_APPLICATION).value, psNewValue, "")
End Property

Public Property Let AppParameters(ByVal psNewValue As String)
  txtAppParameters.Text = IIf(optLink(SSINTLINKSCREEN_APPLICATION).value, psNewValue, "")
End Property

Public Property Let DocumentFilePath(ByVal psNewValue As String)
  txtDocumentFilePath.Text = IIf(optLink(SSINTLINKSCREEN_DOCUMENT).value, psNewValue, "")
End Property

Public Property Let DisplayDocumentHyperlink(ByVal pbNewValue As Boolean)
  chkDisplayDocumentHyperlink.value = IIf(optLink(SSINTLINKSCREEN_DOCUMENT).value, IIf(pbNewValue, vbChecked, vbUnchecked), vbUnchecked)
End Property

Private Sub txtURL_GotFocus()
  UI.txtSelText
End Sub

Public Property Get Prompt() As String
  Prompt = IIf(miLinkType = SSINTLINK_BUTTON, txtPrompt.Text, "")
End Property

Public Property Let Prompt(ByVal psNewValue As String)
  txtPrompt.Text = IIf(miLinkType = SSINTLINK_BUTTON, psNewValue, "")
End Property

Public Property Get HRProScreenID() As String

  If (cboHRProScreen.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_HRPRO).value) Then
    
    HRProScreenID = ""
  Else
    HRProScreenID = CStr(cboHRProScreen.ItemData(cboHRProScreen.ListIndex))
  End If

End Property

Public Property Get StartMode() As String

  If (cboHRProScreen.ListIndex < 0) Or _
    (Not optLink(SSINTLINKSCREEN_HRPRO).value) Then
    
    StartMode = ""
  Else
    StartMode = CStr(cboStartMode.ItemData(cboStartMode.ListIndex))
  End If

End Property

Public Property Let HRProScreenID(ByVal psNewValue As String)

  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim sSQL As String
  Dim rsScreens As dao.Recordset
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And _
    (optLink(SSINTLINKSCREEN_HRPRO).value) And _
    (Len(psNewValue) > 0) Then
    ' Get the given screen's table.
    sSQL = "SELECT tmpScreens.tableID" & _
      " FROM tmpScreens" & _
      " WHERE (tmpScreens.deleted = FALSE)" & _
      " AND (tmpScreens.screenID = " & psNewValue & ")"
    Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    If Not (rsScreens.EOF And rsScreens.BOF) Then
      For iLoop = 0 To cboHRProTable.ListCount - 1
        If cboHRProTable.ItemData(iLoop) = CLng(rsScreens!TableID) Then
          cboHRProTable.ListIndex = iLoop
          
          For iLoop2 = 0 To cboHRProScreen.ListCount - 1
            If cboHRProScreen.ItemData(iLoop2) = CLng(psNewValue) Then
              cboHRProScreen.ListIndex = iLoop2
              Exit For
            End If
          Next iLoop2
          
          Exit For
        End If
      Next iLoop
    End If
    rsScreens.Close
    Set rsScreens = Nothing
  End If
  
End Property

Public Property Let StartMode(ByVal psNewValue As String)
  Dim iLoop As Integer
  
  If (miLinkType <> SSINTLINK_HYPERTEXT) And _
    (optLink(SSINTLINKSCREEN_HRPRO).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboStartMode.ListCount - 1
      If cboStartMode.ItemData(iLoop) = CLng(psNewValue) Then
        cboStartMode.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
End Property

Public Property Let UtilityID(ByVal psNewValue As String)

  Dim iLoop As Integer
  
  If (optLink(SSINTLINKSCREEN_UTILITY).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtility.ListCount - 1
      If cboHRProUtility.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtility.ListIndex = iLoop
        Exit For
      End If
    Next iLoop
  End If
  
End Property

Public Property Let UtilityType(ByVal psNewValue As String)

  Dim iLoop As Integer

  If (optLink(SSINTLINKSCREEN_UTILITY).value) And _
    (Len(psNewValue) > 0) Then

    For iLoop = 0 To cboHRProUtilityType.ListCount - 1
      If cboHRProUtilityType.ItemData(iLoop) = CLng(psNewValue) Then
        cboHRProUtilityType.ListIndex = iLoop
        Exit For
      End If
    Next iLoop

    GetHRProUtilities cboHRProUtilityType.ItemData(cboHRProUtilityType.ListIndex)
  End If
    
End Property

Public Property Get PageTitle() As String
  PageTitle = IIf(optLink(SSINTLINKSCREEN_HRPRO).value, _
    txtPageTitle.Text, "")
End Property

Public Property Let PageTitle(ByVal psNewValue As String)
  txtPageTitle.Text = IIf(optLink(SSINTLINKSCREEN_HRPRO).value, _
    psNewValue, "")
End Property
 
Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Get ViewID() As Long
  ViewID = mlngViewID
End Property

Public Property Get TableViewName() As String
  TableViewName = msTableViewName
End Property

