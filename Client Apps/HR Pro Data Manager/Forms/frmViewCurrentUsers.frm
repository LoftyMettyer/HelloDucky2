VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmViewCurrentUsers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Current Users"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1061
   Icon            =   "frmViewCurrentUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraUsers 
      Height          =   3165
      Left            =   135
      TabIndex        =   1
      Top             =   75
      Width           =   6615
      Begin SSDataWidgets_B.SSDBGrid grdUsers 
         Height          =   2670
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   6255
         _Version        =   196617
         DataMode        =   2
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
         Columns(0).Width=   3360
         Columns(0).Caption=   "User"
         Columns(0).Name =   "User"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2752
         Columns(1).Caption=   "Machine"
         Columns(1).Name =   "Machine"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   4842
         Columns(2).Caption=   "Module"
         Columns(2).Name =   "Module"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   11033
         _ExtentY        =   4710
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
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   400
      Left            =   4200
      TabIndex        =   3
      Top             =   3405
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5550
      TabIndex        =   0
      Top             =   3405
      Width           =   1200
   End
End
Attribute VB_Name = "frmViewCurrentUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()

  GetUsers
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

'  If grdUsers.Rows > grdUsers.VisibleRows Then
'    grdUsers.Columns("Module").Width = 1830
'    grdUsers.ScrollBars = ssScrollBarsVertical
'  Else
'    grdUsers.Columns("Module").Width = 2400
'    grdUsers.ScrollBars = ssScrollBarsNone
'  End If

End Sub

Private Sub cmdOK_Click()

  Unload Me
  
End Sub

Private Sub Form_Activate()

  GetUsers
  
End Sub

'Private Sub GetUsers()
'
'  Dim rsUsers As Recordset
'  Dim sDisplay As String
'  Dim sSQL As String
'  Dim sDatabase As String
'  Dim sComputerName As String
'
'  'Dim sSystemName As String
'  'Dim sSecurityName As String
'  'Dim sUserModuleName As String
'
'  Dim sProgName As String
'  Dim sHostName As String
'  Dim sLoginName As String
'
'  On Error GoTo ErrRefresh
'
'  Screen.MousePointer = vbHourglass
'
'  grdUsers.RemoveAll
'
'  'Now we're connected, check for number of users logged on. First check if anyone is using
'  'NOT THESE at the moment = System Manager or Security Manager
'  'Data manager  (NB Sys/Sec left in bcos in future, poss read only access)
'
'
'  ''MH20010829 Only "sa" can run this SP...
'  '''MH20010823 Fault 2600
'  ''sSQL = "IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id('sp_ASRIntCheckPolls') AND sysstat & 0xf = 4) " & _
'  ''      "BEGIN EXEC sp_ASRIntCheckPolls END"
'  ''datGeneral.ExecuteSql sSQL, ""
'
'  ' Generate our view of the sysprocesses table -> ASRTempSysProcesses
'  sSQL = "IF EXISTS (SELECT Name FROM sysobjects WHERE id = object_id('spASRGenerateSysProcesses') AND sysstat & 0xf = 4) " & _
'         "EXEC spASRGenerateSysProcesses " & _
'         "ELSE BEGIN " & _
'         " IF EXISTS (SELECT Name FROM dbo.sysobjects WHERE id = object_id(N'[dbo].[ASRTempSysProcesses]') and OBJECTPROPERTY(id, N'IsTable') = 1) DROP TABLE dbo.ASRTempSysProcesses" & _
'         " SELECT * INTO ASRTempSysProcesses FROM master..sysprocesses" & _
'         " END"
'  gADOCon.Execute sSQL
'
''  sSQL = "SELECT DISTINCT hostname, loginame, program_name, hostprocess " & _
''     "FROM ASRTempSysProcesses " & _
''     "WHERE program_name like 'HR Pro%' " & _
''     "AND dbid in (" & _
''                   "SELECT dbid " & _
''                   "FROM master..sysdatabases " & _
''                   "WHERE name = '" & gsDatabaseName & "') " & _
''     "ORDER BY loginame"
' sSQL = "spASRGetCurrentUsers"
'
'  Set rsUsers = datGeneral.GetReadOnlyRecords(sSQL)
'
'  Do While Not rsUsers.EOF
'
'    sProgName = Trim(rsUsers!program_name)
'    sHostName = Trim(rsUsers!HostName)
'    sLoginName = Trim(rsUsers!Loginame)
'
'    'Ignore this app on this PC if this login..
'    If LCase(Trim(sHostName)) <> LCase(Trim(UI.GetHostName)) Or _
'       LCase(Trim(sProgName)) <> LCase(Trim(App.ProductName)) Or _
'       LCase(Trim(sLoginName)) <> LCase(Trim(gsUserName)) Then
'
'      grdUsers.AddItem Trim(sLoginName) & vbTab & Trim(sHostName) & vbTab & IIf(LCase(Trim(sProgName)) = "", "HR Pro", Trim(sProgName))
'
'    End If
'
'    rsUsers.MoveNext
'
'  Loop
'
'  rsUsers.Close
'  Set rsUsers = Nothing
'
'  Form_Resize
'  Screen.MousePointer = vbNormal
'
'  Exit Sub
'
'ErrRefresh:
'
'  Screen.MousePointer = vbNormal
'  MsgBox "Error whilst refreshing the grid." & vbCrLf & vbCrLf & Err.Number & " - " & Err.Description, vbExclamation + vbOKOnly, App.Title
'
'End Sub

Private Function GetUsers() As Boolean
  
  Dim intTempPointer As Integer
  Dim fOK As Boolean
  
  On Local Error GoTo LocalErr
  
  fOK = True
  
  intTempPointer = Screen.MousePointer
  Screen.MousePointer = vbHourglass

  fOK = CurrentUsersPopulate(grdUsers)
  Form_Resize

  Screen.MousePointer = intTempPointer

TidyAndExit:
  GetUsers = fOK

Exit Function

LocalErr:
  fOK = False
  Resume TidyAndExit

End Function


