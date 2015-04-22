VERSION 5.00
Begin VB.Form frmDefProp 
   Caption         =   "Definition Properties"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1027
   Icon            =   "frmDefProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCannotChangeType 
      Height          =   660
      Left            =   105
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   195
         Picture         =   "frmDefProp.frx":000C
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblCannotChangeType 
         Caption         =   "Cannot Change Type Text"
         Height          =   240
         Left            =   750
         TabIndex        =   15
         Top             =   240
         Width           =   4695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Usage :"
      Height          =   2100
      Left            =   120
      TabIndex        =   10
      Top             =   2440
      Width           =   6090
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   150
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   5730
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   5010
      TabIndex        =   8
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2410
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6090
      Begin VB.TextBox txtRecCount 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   12
         Top             =   1900
         Width           =   4275
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   1
         Top             =   300
         Width           =   4275
      End
      Begin VB.TextBox txtRun 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   7
         Top             =   1500
         Width           =   4275
      End
      Begin VB.TextBox txtSaved 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   5
         Top             =   1100
         Width           =   4275
      End
      Begin VB.TextBox txtCreated 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1600
         TabIndex        =   3
         Top             =   700
         Width           =   4275
      End
      Begin VB.Label lblRecCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Record Count :"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         Height          =   195
         Left            =   225
         TabIndex        =   0
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblRun 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Run :"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblSave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Save :"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   1160
         Width           =   810
      End
      Begin VB.Label lblCreated 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created :"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   760
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmDefProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private datData As clsDataAccess

Private mblnUsage As Boolean
Private miUsageCount As Integer
Private mbChangeErrorMode As Boolean

Public Property Get UtilName() As String
  UtilName = txtName.Text
End Property

Public Property Let UtilName(ByVal strNewValue As String)
  txtName.Text = strNewValue
End Property


Private Function FormatText(varDate As Variant, varUser As Variant, varHost As Variant) As String
  If IsNull(varUser) Then
    FormatText = "<None>"
  Else
    FormatText = Format(varDate, DateFormat & " hh:nn") & _
                 "  by  " & StrConv(varUser, vbProperCase)
    If ASRDEVELOPMENT Then
      FormatText = FormatText & " (" & StrConv(varHost, vbUpperCase) & ")"
    End If
  End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Load()

  Hook Me.hWnd, 5000, 3120
  
  Set datData = New clsDataAccess
  mbChangeErrorMode = False
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Set datData = Nothing
  
  If Not mbChangeErrorMode Then
    If Not Me.BorderStyle = vbFixedDialog Then
      If mblnUsage = True Then
        SavePCSetting gsDatabaseName & "\DefProp\Usage", "Top", Me.Top
        SavePCSetting gsDatabaseName & "\DefProp\Usage", "Left", Me.Left
        SavePCSetting gsDatabaseName & "\DefProp\Usage", "Width", Me.Width
        SavePCSetting gsDatabaseName & "\DefProp\Usage", "Height", Me.Height
      Else
        SavePCSetting gsDatabaseName & "\DefProp\NoUsage", "Top", Me.Top
        SavePCSetting gsDatabaseName & "\DefProp\NoUsage", "Left", Me.Left
        SavePCSetting gsDatabaseName & "\DefProp\NoUsage", "Width", Me.Width
      End If
    End If
  End If
  
End Sub

Private Sub Form_Resize()

  Dim objControl As Control
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
'  If Me.Width < 5000 Then Me.Width = 5000
'
'  If Frame2.Visible = False And mblnUsage = False Then
'    If Me.Height <> 3120 Then Me.Height = 3120
'  Else
'    If Me.Height < 4500 Then Me.Height = 4500
'  End If

  'UI.ClipForForm Me, IIf(Frame2.Visible = False And mblnUsage = False, 3120, 4500), 5000
  
  For Each objControl In Me.Controls
    If TypeOf objControl Is TextBox Then
      objControl.Width = Me.ScaleWidth - 2100
    ElseIf TypeOf objControl Is Frame Then
      objControl.Width = Me.ScaleWidth - 300
      If objControl.Name = "Frame2" And objControl.Visible Then objControl.Height = Me.ScaleHeight - 2750
      If objControl.Name = "Frame2" And mbChangeErrorMode Then objControl.Height = Me.ScaleHeight - 1850
      If objControl.Name = "Frame1" And (Frame2.Visible = False) And mblnUsage = False Then objControl.Height = Me.ScaleHeight - 750
    ElseIf TypeOf objControl Is ListBox Then
      objControl.Width = Me.ScaleWidth - 650
      If objControl.Visible Then objControl.Height = Me.ScaleHeight - 3200
      If objControl.Visible And mbChangeErrorMode Then objControl.Height = Me.ScaleHeight - 2300
   ElseIf TypeOf objControl Is CommandButton Then
      objControl.Left = Me.ScaleWidth - 200 - objControl.Width
      objControl.Top = Me.ScaleHeight - 550
    End If
  Next objControl
  
  ' Get rid of the icon off the form
  RemoveIcon Me
  
End Sub

Public Sub PopulateUtil(utlType As UtilityType, lngID As Long)
  
  Call GetData(utlType, lngID)
  Call DrawControls(utlType)

  Unhook Me.hWnd
  If Frame2.Visible = False And mblnUsage = False Then
    If Me.Height <> 3120 Then Me.Height = 3120
    Hook Me.hWnd, 5000, 3120
  Else
    Hook Me.hWnd, 5000, 4500
  End If

End Sub


Private Sub GetData(utlType As UtilityType, lngID As Long)

  Dim datData As clsDataAccess
  Dim rsTemp As Recordset
  Dim strSQL As String


  Set datData = New clsDataAccess

  strSQL = "SELECT * FROM ASRSysUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID) & _
           " AND Type = " & CStr(utlType)

  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  With rsTemp
    
    If Not .BOF And Not .EOF Then
      txtCreated = FormatText(!CreatedDate, !CreatedBy, !CreatedHost)
      txtSaved = FormatText(!SavedDate, !SavedBy, !SavedHost)
      txtRun = FormatText(!RunDate, !RunBy, !RunHost)
    Else
      txtCreated = "<None>"
      txtSaved = "<None>"
      txtRun = "<None>"
    End If

  End With
  
  rsTemp.Close
  Set rsTemp = Nothing
  Set datData = Nothing

End Sub


Private Sub DrawControls(utlType As UtilityType)

  Dim blnUsage As Boolean
  Dim blnLastRun As Boolean
  Dim blnRecCount As Boolean
  Dim lngOffset As Long
  
  blnLastRun = (utlType <> utlCalculation And _
                utlType <> utlFilter And _
                utlType <> utlPicklist And _
                utlType <> utlOrder And _
                utlType <> utlEmailAddress And _
                utlType <> utlEmailGroup And _
                utlType <> utlLabelType)

  'blnRecCount = (utlType = utlFilter Or _
                 utlType = utlPicklist)
  blnRecCount = False
  
  blnUsage = True '(utlType <> utlBatchJob)
  If utlType = utlBatchJob Or utlType = utlReportPack Then
    Me.Frame2.Caption = "Job Details :"
  End If

  mblnUsage = blnUsage

  lngOffset = 0

  If Not blnLastRun Then
    txtRun.Visible = False
    lblRun.Visible = False
    lngOffset = lngOffset + 400
  End If
  
  
  If Not blnRecCount Then
    txtRecCount.Visible = False
    lblRecCount.Visible = False
    lngOffset = lngOffset + 400
  Else
    txtRecCount.Top = txtRecCount.Top - lngOffset
    lblRecCount.Top = lblRecCount.Top - lngOffset
  End If
    
  
  Frame1.Height = Frame1.Height - lngOffset
  
  
  If blnUsage Then
    Frame2.Top = Frame2.Top - lngOffset
  Else
    Frame2.Visible = False
    lngOffset = lngOffset + Frame2.Height
  End If
  
  
  cmdOK.Top = cmdOK.Top - lngOffset
  Me.Height = Me.Height - lngOffset

  ' Leave form size at its default if we are using as a change type form
  If mbChangeErrorMode Then
    Me.BorderStyle = vbFixedDialog
    UI.frmAtCenter Me
    Exit Sub
  End If

  If mblnUsage Then
    ' If its the first time, then default to centering the form
    If (GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Width", 0) = 0) And _
        (GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Left", 0) = 0) And _
        (GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Top", 0) = 0) And _
        (GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Height", 0) = 0) Then
      UI.frmAtCenter Me
      Exit Sub
    End If
    Me.Width = GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Width", 3000)
    Me.Left = GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Left", (Screen.Height - 6400) / 2)
    Me.Height = GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Height", 4500)
    Me.Top = GetPCSetting(gsDatabaseName & "\DefProp\Usage", "Top", (Screen.Height - Me.Height) / 2)
  Else
    ' If its the first time, then default to centering the form
    If (GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Width", 0) = 0) And _
        (GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Left", 0) = 0) And _
        (GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Top", 0) = 0) Then
      UI.frmAtCenter Me
      Exit Sub
    End If
    Me.Width = GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Width", 3000)
    Me.Left = GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Left", (Screen.Height - 6400) / 2)
    Me.Top = GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Top", (Screen.Height - Me.Height) / 2)
    Me.Height = GetPCSetting(gsDatabaseName & "\DefProp\NoUsage", "Height", 3120)
  End If
  
End Sub


Private Sub cmdOK_Click()
  Unload Me
End Sub


Public Function CheckForUseage(piType As UtilityType, plngItemID As Long) As Boolean
  
  Dim strSQL As String
  Dim rsTemp As Recordset
  Dim strGlobalFunction As String
  Dim strMatchReport As String
  Dim strModuleDefinition As String
  Dim strID As String
  Dim strRootIDs As String
  Dim sBatchJobType As String
  
  strGlobalFunction = _
    "CASE WHEN ASRSysGlobalFunctions.Type = 'A' " & _
      "THEN 'Global Add' " & _
    "WHEN ASRSysGlobalFunctions.Type = 'D' " & _
      "THEN 'Global Delete' " & _
    "WHEN ASRSysGlobalFunctions.Type = 'U' " & _
      "THEN 'Global Update' " & _
    "ELSE 'Global Function' END"

  strMatchReport = _
    "CASE WHEN ASRSysMatchReportName.matchReportType = 0 " & _
      "THEN 'Match Report' " & _
    "WHEN ASRSysMatchReportName.matchReportType = 1 " & _
      "THEN 'Succession Planning' " & _
    "WHEN ASRSysMatchReportName.matchReportType = 2 " & _
      "THEN 'Career Progression' " & _
    "ELSE 'Match Report' END"

  strModuleDefinition = _
    "CASE WHEN ASRSysModuleSetup.ModuleKey = '" & gsMODULEKEY_TRAININGBOOKING & "' " & _
      "THEN 'Training Booking'" & _
    "WHEN ASRSysModuleSetup.ModuleKey = '" & gsMODULEKEY_PERSONNEL & "' " & _
      "THEN 'Personnel'" & _
    "WHEN ASRSysModuleSetup.ModuleKey = '" & gsMODULEKEY_ABSENCE & "' " & _
      "THEN 'Absence'" & _
    "ELSE '<unknown>' END"

  strID = CStr(plngItemID)
  List1.Clear

  Select Case piType
  Case utlPicklist

    CheckSystemSettings strID, "P", "Absence Breakdown"
    CheckSystemSettings strID, "P", "Bradford Factor"
    
    'Cross Tab
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.pickListID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.pickListID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    End If

    '9 Box Grid
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.pickListID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.pickListID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    End If
    

    'Data Transfer
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Data Transfer'," & _
        " ASRSysDataTransferName.Name," & _
        " ASRSysDataTransferName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysDataTransferName" & _
        " WHERE ASRSysDataTransferName.pickListID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Data Transfer'," & _
        " ASRSysDataTransferName.Name," & _
        " ASRSysDataTransferName.UserName," & _
        " ASRSysDataTransferAccess.Access" & _
        " FROM ASRSysDataTransferName" & _
        " INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.dataTransferID = ASRSysDataTransferAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysDataTransferAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysDataTransferName.pickListID = " & strID)
    End If
    
    'Export
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysExportName" & _
        " WHERE AsrSysExportName.pickList = " & strID & " OR AsrSysExportName.Parent1Picklist = " & strID & " OR AsrSysExportName.Parent2Picklist = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " ASRSysExportAccess.Access" & _
        " FROM AsrSysExportName" & _
        " INNER JOIN ASRSysExportAccess ON AsrSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysExportName.pickList = " & strID & " OR AsrSysExportName.Parent1Picklist = " & strID & " OR AsrSysExportName.Parent2Picklist = " & strID)
    End If
    
    'Globals
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysGlobalFunctions" & _
        " WHERE ASRSysGlobalFunctions.pickListID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " ASRSysGlobalAccess.Access" & _
        " FROM ASRSysGlobalFunctions" & _
        " INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysGlobalAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysGlobalFunctions.pickListID = " & strID)
    End If
    
    'Custom Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCustomReportsName" & _
        " WHERE AsrSysCustomReportsName.pickList = " & strID & _
        " OR AsrSysCustomReportsName.Parent1Picklist = " & strID & _
        " OR AsrSysCustomReportsName.Parent2Picklist = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM AsrSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportAccess ON AsrSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCustomReportsName.pickList = " & strID & _
        " OR AsrSysCustomReportsName.Parent1Picklist = " & strID & _
        " OR AsrSysCustomReportsName.Parent2Picklist = " & strID)
    End If
      
    'Mail Merge/Envelopes & Labels
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.PickListID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeName.PickListID = " & strID)
    End If

    'Match Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMatchReportName" & _
        " WHERE ASRSysMatchReportName.Table1Picklist = " & strID & _
        " OR ASRSysMatchReportName.Table2Picklist = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " ASRSysMatchReportAccess.Access" & _
        " FROM ASRSysMatchReportName" & _
        " INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMatchReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMatchReportName.Table1Picklist = " & strID & _
        " OR ASRSysMatchReportName.Table2Picklist = " & strID)
    End If
  
    'Calendar Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCalendarReports" & _
        " WHERE AsrSysCalendarReports.pickList = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " ASRSysCalendarReportAccess.Access" & _
        " FROM AsrSysCalendarReports" & _
        " INNER JOIN ASRSysCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSysCalendarReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCalendarReports.pickList = " & strID)
    End If

    'Record Profile
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysRecordProfileName" & _
        " WHERE ASRSysRecordProfileName.pickListID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " ASRSysRecordProfileAccess.Access" & _
        " FROM ASRSysRecordProfileName" & _
        " INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysRecordProfileAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysRecordProfileName.pickListID = " & strID)
    End If

    CheckSystemSettings strID, "P", "Stability"
    CheckSystemSettings strID, "P", "Turnover"

  Case utlFilter

    strRootIDs = GetExprRootIDs(strID)
    
    CheckSystemSettings strID, "F", "Absence Breakdown"
    CheckSystemSettings strID, "F", "Bradford Factor"
    
    'Calculations
    If strRootIDs <> vbNullString Then
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Calculation', Name, UserName, 'RW' AS Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
                " AND ExprID IN (" & strRootIDs & ")")
      Else
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Calculation', Name, UserName, Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
                " AND ExprID IN (" & strRootIDs & ")")
      End If
    End If
    
    'Cross Tab Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.FilterID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.FilterID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    End If

    '9 Box Grid
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.FilterID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.FilterID = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    End If
    
    'Custom Report Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " WHERE AsrSysCustomReportsName.Filter = " & strID & _
        " OR AsrSysCustomReportsName.Parent1Filter = " & strID & _
        " OR AsrSysCustomReportsName.Parent2Filter = " & strID & _
        " OR ASRSYSCustomReportsChildDetails.ChildFilter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM AsrSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " INNER JOIN ASRSysCustomReportAccess ON AsrSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCustomReportsName.Filter = " & strID & _
        " OR AsrSysCustomReportsName.Parent1Filter = " & strID & _
        " OR AsrSysCustomReportsName.Parent2Filter = " & strID & _
        " OR ASRSYSCustomReportsChildDetails.ChildFilter = " & strID)
    End If

    'Data Transfer Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Data Transfer'," & _
        " ASRSysDataTransferName.Name," & _
        " ASRSysDataTransferName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysDataTransferName" & _
        " WHERE ASRSysDataTransferName.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Data Transfer'," & _
        " ASRSysDataTransferName.Name," & _
        " ASRSysDataTransferName.UserName," & _
        " ASRSysDataTransferAccess.Access" & _
        " FROM ASRSysDataTransferName" & _
        " INNER JOIN ASRSysDataTransferAccess ON ASRSysDataTransferName.dataTransferID = ASRSysDataTransferAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysDataTransferAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysDataTransferName.FilterID = " & strID)
    End If
    
    'Export Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysExportName" & _
        " WHERE AsrSysExportName.Filter = " & strID & " OR AsrSysExportName.Parent1Filter = " & strID & " OR AsrSysExportName.Parent2Filter = " & strID & " OR AsrSysExportName.ChildFilter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " ASRSysExportAccess.Access" & _
        " FROM AsrSysExportName" & _
        " INNER JOIN ASRSysExportAccess ON AsrSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysExportName.Filter = " & strID & " OR AsrSysExportName.Parent1Filter = " & strID & " OR AsrSysExportName.Parent2Filter = " & strID & " OR AsrSysExportName.ChildFilter = " & strID)
    End If

    'Filters
    If strRootIDs <> vbNullString Then
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Filter', Name, UserName, 'RW' AS Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
                " AND ExprID IN (" & strRootIDs & ")")
      Else
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Filter', Name, UserName, Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
                " AND ExprID IN (" & strRootIDs & ")")
      End If
    End If
    
    'Globals Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysGlobalFunctions" & _
        " WHERE ASRSysGlobalFunctions.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " ASRSysGlobalAccess.Access" & _
        " FROM ASRSysGlobalFunctions" & _
        " INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysGlobalAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysGlobalFunctions.FilterID = " & strID)
    End If
    
    'Mail Merge/Envelopes & Labels
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeName.FilterID = " & strID)
    End If

    'Match Report Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMatchReportName" & _
        " WHERE ASRSysMatchReportName.Table1Filter = " & strID & _
        " OR ASRSysMatchReportName.Table2Filter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " ASRSysMatchReportAccess.Access" & _
        " FROM ASRSysMatchReportName" & _
        " INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMatchReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMatchReportName.Table1Filter = " & strID & _
        " OR ASRSysMatchReportName.Table2Filter = " & strID)
    End If

    'Calendar Report Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID" & _
        " WHERE AsrSysCalendarReports.filter = " & strID & _
        " OR ASRSysCalendarReportEvents.filterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " ASRSysCalendarReportAccess.Access" & _
        " FROM AsrSysCalendarReports" & _
        " LEFT OUTER JOIN ASRSysCalendarReportEvents ON ASRSysCalendarReports.ID = ASRSysCalendarReportEvents.CalendarReportID" & _
        " INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = AsrSysCalendarReports.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCalendarReports.filter = " & strID & _
        " OR ASRSysCalendarReportEvents.filterID = " & strID)
    End If
    
    'Record Profile Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " WHERE ASRSysRecordProfileName.FilterID = " & strID & _
        " OR ASRSYSRecordProfileTables.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " ASRSysRecordProfileAccess.Access" & _
        " FROM ASRSysRecordProfileName" & _
        " INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysRecordProfileAccess.groupname = b.name" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " WHERE ASRSysRecordProfileName.FilterID = " & strID & _
        " OR ASRSYSRecordProfileTables.FilterID = " & strID)
    End If
    
    'Report Pack - Override Filters
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Report Pack'," & _
        " ASRSysBatchJobName.Name," & _
        " ASRSysBatchJobName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysBatchJobName" & _
        " WHERE ASRSysBatchJobName.OverrideFilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Report Pack'," & _
        " ASRSysBatchJobName.Name," & _
        " ASRSysBatchJobName.UserName," & _
        " ASRSysBatchJobAccess.Access" & _
        " FROM ASRSysBatchJobName" & _
        " INNER JOIN ASRSysBatchJobAccess ON ASRSysBatchJobAccess.ID = ASRSysBatchJobName.ID" & _
        " INNER JOIN sysusers b ON ASRSysBatchJobAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysBatchJobName.OverrideFilterID = " & strID)
    End If
    
    CheckSystemSettings strID, "F", "Stability"
    CheckSystemSettings strID, "F", "Turnover"
  
  Case utlCalculation
    
    strRootIDs = GetExprRootIDs(strID)

    CheckSystemSettings strID, "X", "Absence Breakdown"
    CheckSystemSettings strID, "X", "Bradford Factor"
    
    CheckUserSettings "diaryprint", strID, "Diary Print Options"
    
    'Calculations
    If strRootIDs <> vbNullString Then
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Calculation', Name, UserName, 'RW' AS Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
                " AND ExprID IN (" & strRootIDs & ")")
      Else
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Calculation', Name, UserName, Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
                " AND ExprID IN (" & strRootIDs & ")")
      End If
    End If
    
    'Calendar Report Calculation
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCalendarReports" & _
        " WHERE AsrSysCalendarReports.DescriptionExpr = " & strID & _
        "   OR AsrSysCalendarReports.StartDateExpr = " & strID & _
        "   OR AsrSysCalendarReports.EndDateExpr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " ASRSysCalendarReportAccess.Access" & _
        " FROM AsrSysCalendarReports" & _
        " INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = AsrSysCalendarReports.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCalendarReports.DescriptionExpr = " & strID & _
        "   OR AsrSysCalendarReports.StartDateExpr = " & strID & _
        "   OR AsrSysCalendarReports.EndDateExpr = " & strID)
    End If
    
    'Custom Report Calculation
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = ASRSysCustomReportsDetails.CustomReportID" & _
        " WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' AND ASRSysCustomReportsDetails.colExprID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM AsrSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = ASRSysCustomReportsDetails.CustomReportID" & _
        " INNER JOIN ASRSysCustomReportAccess ON AsrSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' AND ASRSysCustomReportsDetails.colExprID = " & strID)
    End If
    
    'Export Calculation
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " ASRSysExportName.Name," & _
        " ASRSysExportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysExportDetails" & _
        " INNER JOIN ASRSysExportName ON ASRSysExportDetails.exportID = ASRSysExportName.ID " & _
        " WHERE UPPER(ASRSysExportDetails.type) = 'X' AND ASRSysExportDetails.colExprID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " ASRSysExportName.Name," & _
        " ASRSysExportName.UserName," & _
        " ASRSysExportAccess.Access" & _
        " FROM ASRSysExportDetails" & _
        " INNER JOIN ASRSysExportName ON ASRSysExportDetails.ID = ASRSysExportName.ID " & _
        " INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE UPPER(ASRSysExportDetails.type) = 'X' AND ASRSysExportDetails.colExprID = " & strID)
    End If
    
    'Filters
    If strRootIDs <> vbNullString Then
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Filter', Name, UserName, 'RW' AS Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
                " AND ExprID IN (" & strRootIDs & ")")
      Else
        Call GetNameWhereUsed( _
                "SELECT DISTINCT 'Filter', Name, UserName, Access " & _
                "FROM ASRSysExpressions" & _
                " WHERE Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
                " AND ExprID IN (" & strRootIDs & ")")
      End If
    End If

    'Globals Calculation
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysGlobalItems" & _
        " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID " & _
        " WHERE ASRSysGlobalItems.ValueType = 4 AND ASRSysGlobalItems.ExprID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strGlobalFunction & "," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " ASRSysGlobalAccess.Access" & _
        " FROM ASRSysGlobalItems" & _
        " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID " & _
        " INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysGlobalAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysGlobalItems.ValueType = 4 AND ASRSysGlobalItems.ExprID = " & strID)
    End If
    
    'Mail Merge/Envelopes & Labels
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " WHERE ASRSysMailMergeColumns.ColumnID = " & strID & _
        "   AND upper(ASRSysMailMergeColumns.type) = 'E'")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
        " ELSE 'Mail Merge' END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeColumns.ColumnID = " & strID & _
        "   AND upper(ASRSysMailMergeColumns.type) = 'E'")
    End If

    CheckSystemSettings strID, "X", "Stability"
    CheckSystemSettings strID, "X", "Turnover"

  Case utlCrossTab, utlCustomReport, utlDataTransfer, utlExport, UtlGlobalAdd, _
       utlGlobalDelete, utlGlobalUpdate, utlMailMerge, utlImport, _
       utlCalendarReport, utlRecordProfile, utlLabel, _
       utlMatchReport, utlSuccession, utlCareer

      sBatchJobType = GetBatchJobType(piType)

      'Check if this has been used in a batch job
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
          "SELECT DISTINCT 'Batch Job'," & _
          " ASRSysBatchJobName.Name," & _
          " ASRSysBatchJobName.UserName," & _
          " '" & ACCESS_READWRITE & "' AS access" & _
          " FROM ASRSysBatchJobDetails" & _
          " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
          " WHERE AsrSysBatchJobName.IsBatch = 1 " & _
          "   AND ASRSysBatchJobDetails.JobType = '" & sBatchJobType & "' " & _
          "   AND ASRSysBatchJobDetails.JobID = " & strID)
      Else
        Call GetNameWhereUsed( _
          "SELECT DISTINCT 'Batch Job'," & _
          " ASRSysBatchJobName.Name," & _
          " ASRSysBatchJobName.UserName," & _
          " ASRSysBatchJobAccess.Access" & _
          " FROM ASRSysBatchJobDetails" & _
          " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
          " INNER JOIN ASRSysBatchJobAccess ON AsrSysBatchJobName.ID = ASRSysBatchJobAccess.ID" & _
          " INNER JOIN sysusers b ON ASRSysBatchJobAccess.groupname = b.name" & _
          "   AND b.name = '" & gsUserGroup & "'" & _
          " WHERE AsrSysBatchJobName.IsBatch = 1 " & _
          "   AND ASRSysBatchJobDetails.JobType = '" & sBatchJobType & "' " & _
          "   AND ASRSysBatchJobDetails.JobID = " & strID)
      End If
      'Check if this has been used in a Report Pack
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
          "SELECT DISTINCT 'Report Pack'," & _
          " ASRSysBatchJobName.Name," & _
          " ASRSysBatchJobName.UserName," & _
          " '" & ACCESS_READWRITE & "' AS access" & _
          " FROM ASRSysBatchJobDetails" & _
          " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
          " WHERE AsrSysBatchJobName.IsBatch = 0 " & _
          "   AND ASRSysBatchJobDetails.JobType = '" & sBatchJobType & "' " & _
          "   AND ASRSysBatchJobDetails.JobID = " & strID)
      Else
        Call GetNameWhereUsed( _
          "SELECT DISTINCT 'Report Pack'," & _
          " ASRSysBatchJobName.Name," & _
          " ASRSysBatchJobName.UserName," & _
          " ASRSysBatchJobAccess.Access" & _
          " FROM ASRSysBatchJobDetails" & _
          " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
          " INNER JOIN ASRSysBatchJobAccess ON AsrSysBatchJobName.ID = ASRSysBatchJobAccess.ID" & _
          " INNER JOIN sysusers b ON ASRSysBatchJobAccess.groupname = b.name" & _
          "   AND b.name = '" & gsUserGroup & "'" & _
          " WHERE AsrSysBatchJobName.IsBatch = 0 " & _
          "   AND ASRSysBatchJobDetails.JobType = '" & sBatchJobType & "' " & _
          "   AND ASRSysBatchJobDetails.JobID = " & strID)
      End If
    'JPD 20040729 Fault 8978
    CheckModuleSetup piType, plngItemID
    
  Case utlLabelType
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Envelopes & Labels'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.LabelTypeID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Envelopes & Labels'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeName.LabelTypeID = " & strID)
    End If
    
  Case utlEmailAddress
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Default Email', AsrSysTables.TableName, " & _
      "'' as Username, '" & ACCESS_READONLY & "' as Access " & _
      "FROM ASRSysTables " & _
      "WHERE DefaultEmailID = " & CStr(plngItemID))

    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Email Link', AsrSysEmailLinks.Title, AsrSysEmailLinks.LinkID, " & _
      "'' as Username, '" & ACCESS_READONLY & "' as Access " & _
      "FROM ASRSysEmailLinksRecipients " & _
      "JOIN ASRSysEmailLinks on ASRSysEmailLinksRecipients.LinkID = ASRSysEmaillinks.LinkID " & _
      "WHERE RecipientID = " & CStr(plngItemID))

    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Email Group', ASRSysEmailGroupName.Name, ASRSysEmailGroupName.EmailGroupID, UserName, Access " & _
      "FROM ASRSysEmailGroupItems " & _
      "JOIN ASRSysEmailGroupName on ASRSysEmailGroupItems.EmailGroupID = ASRSysEmailGroupName.EmailGroupID " & _
      "WHERE ASRSysEmailGroupItems.EmailDefID = " & CStr(plngItemID))

      'Mail Merge/Envelopes & Labels
      If gfCurrentUserIsSysSecMgr Then
        Call GetNameWhereUsed( _
          "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
          " ELSE 'Mail Merge' END," & _
          " ASRSysMailMergeName.Name," & _
          " ASRSysMailMergeName.UserName," & _
          " '" & ACCESS_READWRITE & "' AS access" & _
          " FROM ASRSysMailMergeName" & _
          " WHERE ASRSysMailMergeName.EmailAddrID = " & strID)
      Else
        Call GetNameWhereUsed( _
          "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.isLabel = 1 THEN 'Envelopes & Labels'" & _
          " ELSE 'Mail Merge' END," & _
          " ASRSysMailMergeName.Name," & _
          " ASRSysMailMergeName.UserName," & _
          " ASRSysMailMergeAccess.Access" & _
          " FROM ASRSysMailMergeName" & _
          " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
          " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
          "   AND b.name = '" & gsUserGroup & "'" & _
          " WHERE ASRSysMailMergeName.EmailAddrID = " & strID)
      End If
  
    'JPD 20061205 Fault 11773
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Workflow', ASRSysWorkflows.Name, " & _
      "'' as Username, '" & ACCESS_READONLY & "' as Access " & _
      "FROM ASRSysWorkflowElements " & _
      "INNER JOIN ASRSysWorkflows ON ASRSysWorkflowElements.workflowID = ASRSysWorkflows.ID " & _
      "WHERE ASRSysWorkflowElements.EmailID = " & CStr(plngItemID))
    
    
  Case utlEmailGroup

    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysBatchJobName.IsBatch = 1 THEN 'Batch Job' ELSE 'Report Pack' END," & _
        " ASRSysBatchJobName.Name," & _
        " ASRSysBatchJobName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysBatchJobDetails" & _
        " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
        " WHERE (EmailSuccess = " & strID & " OR EmailFailed = " & strID & ")")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN ASRSysBatchJobName.IsBatch = 1 THEN 'Batch Job' ELSE 'Report Pack' END," & _
        " ASRSysBatchJobName.Name," & _
        " ASRSysBatchJobName.UserName," & _
        " ASRSysBatchJobAccess.Access" & _
        " FROM ASRSysBatchJobDetails" & _
        " INNER JOIN ASRSysBatchJobName ON ASRSysBatchJobDetails.BatchJobNameID = ASRSysBatchJobName.ID" & _
        " INNER JOIN ASRSysBatchJobAccess ON AsrSysBatchJobName.ID = ASRSysBatchJobAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysBatchJobAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE (EmailSuccess = " & strID & " OR EmailFailed = " & strID & ")")
    End If


    'Cross Tab
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.OutputEmailAddr = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.OutputEmailAddr = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType <> 4")
    End If

    '9 Box Grid
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCrossTab" & _
        " WHERE ASRSysCrossTab.OutputEmailAddr = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT '9-Box Grid Report'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysCrossTab.OutputEmailAddr = " & strID & _
        "   AND ASRSysCrossTab.CrossTabType = 4")
    End If

    'Export
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysExportName" & _
        " WHERE AsrSysExportName.OutputEmailAddr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " AsrSysExportName.Name," & _
        " AsrSysExportName.UserName," & _
        " ASRSysExportAccess.Access" & _
        " FROM AsrSysExportName" & _
        " INNER JOIN ASRSysExportAccess ON AsrSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysExportName.OutputEmailAddr = " & strID)
    End If

    'Custom Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCustomReportsName" & _
        " WHERE AsrSysCustomReportsName.OutputEmailAddr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM AsrSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportAccess ON AsrSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCustomReportsName.OutputEmailAddr = " & strID)
    End If
      
    'Match Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMatchReportName" & _
        " WHERE ASRSysMatchReportName.OutputEmailAddr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT " & strMatchReport & "," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " ASRSysMatchReportAccess.Access" & _
        " FROM ASRSysMatchReportName" & _
        " INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMatchReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMatchReportName.OutputEmailAddr = " & strID)
    End If
  
    'Calendar Report
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCalendarReports" & _
        " WHERE AsrSysCalendarReports.OutputEmailAddr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Calendar Report'," & _
        " AsrSysCalendarReports.Name," & _
        " AsrSysCalendarReports.UserName," & _
        " ASRSysCalendarReportAccess.Access" & _
        " FROM AsrSysCalendarReports" & _
        " INNER JOIN ASRSysCalendarReportAccess ON AsrSysCalendarReports.ID = ASRSysCalendarReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE AsrSysCalendarReports.OutputEmailAddr = " & strID)
    End If

    'Record Profile
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysRecordProfileName" & _
        " WHERE ASRSysRecordProfileName.OutputEmailAddr = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " ASRSysRecordProfileAccess.Access" & _
        " FROM ASRSysRecordProfileName" & _
        " INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysRecordProfileAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysRecordProfileName.OutputEmailAddr = " & strID)
    End If
    
    
  Case utlOrder

    strRootIDs = GetExprRootIDs(strID, True)

    'Calculations
    If strRootIDs <> vbNullString Then
      Call GetNameWhereUsed( _
              "SELECT DISTINCT 'Calculation', Name, UserName, Access " & _
              "FROM ASRSysExpressions" & _
              " WHERE Type = " & CStr(giEXPR_RUNTIMECALCULATION) & _
              " AND ExprID IN (" & strRootIDs & ")")
    End If

    'Default Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Default Order', tableName " & _
      ",'" & ACCESS_READWRITE & "' As Access, '" & UCase(LTrim(datGeneral.UserNameForSQL)) & "' As Username " & _
      "FROM ASRSysTables " & _
      " WHERE defaultOrderID = " & strID)

    'Filters
    If strRootIDs <> vbNullString Then
      Call GetNameWhereUsed( _
              "SELECT DISTINCT 'Filter', Name, UserName, Access " & _
              "FROM ASRSysExpressions" & _
              " WHERE Type = " & CStr(giEXPR_RUNTIMEFILTER) & _
              " AND ExprID IN (" & strRootIDs & ")")
    End If
    
    'Screen Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Screen Order', ASRSysScreens.name " & _
      ", '" & ACCESS_READWRITE & "' As Access, '" & UCase(LTrim(datGeneral.UserNameForSQL)) & "' As Username " & _
      "FROM ASRSysScreens " & _
      " WHERE ASRSysScreens.OrderID = " & strID)
    
    'Module Setup Order
    'JPD 20031009 Fault 7096
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Module Setup', " & strModuleDefinition & " " & _
      ",'" & ACCESS_READWRITE & "' As Access, '" & UCase(LTrim(datGeneral.UserNameForSQL)) & "' As Username " & _
      "FROM ASRSysModuleSetup " & _
      " WHERE parameterType = '" & gsPARAMETERTYPE_ORDERID & "'" & _
      " AND parameterValue = '" & strID & "'")
  
    'Record Profile Order
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysRecordProfileName" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " WHERE ASRSysRecordProfileName.OrderID = " & strID & _
        " OR ASRSYSRecordProfileTables.OrderID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Record Profile'," & _
        " ASRSysRecordProfileName.Name," & _
        " ASRSysRecordProfileName.UserName," & _
        " ASRSysRecordProfileAccess.Access" & _
        " FROM ASRSysRecordProfileName" & _
        " INNER JOIN ASRSysRecordProfileAccess ON ASRSysRecordProfileName.recordProfileID = ASRSysRecordProfileAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysRecordProfileAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " WHERE ASRSysRecordProfileName.OrderID = " & strID & _
        " OR ASRSYSRecordProfileTables.OrderID = " & strID)
    End If
  
    'Custom Report Order
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM AsrSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " WHERE (ASRSYSCustomReportsChildDetails.ChildOrder = " & strID & ")")
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " AsrSysCustomReportsName.Name," & _
        " AsrSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM AsrSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " INNER JOIN ASRSysCustomReportAccess ON AsrSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE (ASRSYSCustomReportsChildDetails.ChildOrder = " & strID & ")")
    End If
  
    'Workflow Record Selector Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Workflow'," & _
      " ASRSysWorkflows.Name," & _
      " '' AS [userName]," & _
      " '" & ACCESS_READWRITE & "' AS access" & _
      " FROM ASRSysWorkflows" & _
      " INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflows.ID = ASRSysWorkflowElements.WorkflowID" & _
      " INNER JOIN ASRSysWorkflowElementItems ON ASRSysWorkflowElements.ID = ASRSysWorkflowElementItems.elementID" & _
      " WHERE (ASRSysWorkflowElementItems.recordOrderID = " & strID & ")")
  
  Case utlBatchJob
    Call GetNameWhereUsed( _
      "SELECT DISTINCT JobType, JobID, JobOrder, " & _
      "'' as Username, '" & ACCESS_READONLY & "' as Access " & _
      "FROM ASRSYSBatchJobDetails " & _
      "WHERE ASRSYSBatchJobDetails.BatchJobNameID = " & CStr(plngItemID) & " ORDER BY JobOrder", True)

  Case utlReportPack
    Call GetNameWhereUsed( _
      "SELECT DISTINCT JobType, JobID, JobOrder, " & _
      "'' as Username, '" & ACCESS_READONLY & "' as Access " & _
      "FROM ASRSYSBatchJobDetails " & _
      "WHERE ASRSYSBatchJobDetails.BatchJobNameID = " & CStr(plngItemID) & " ORDER BY JobOrder", True)
      
  Case utlLabelType
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Envelopes & Labels'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.LabelTypeID = " & CStr(plngItemID))
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Envelopes & Labels'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeName.PickListID = " & CStr(plngItemID))
    End If

  Case utlDocumentMapping
  
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Mail Merge'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.DocumentMapID = " & CStr(plngItemID))
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Mail Merge'," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        "   AND b.name = '" & gsUserGroup & "'" & _
        " WHERE ASRSysMailMergeName.DocumentMapID = " & CStr(plngItemID))
    End If
  
  Case Else
    List1.AddItem "<Error Checking Usage>"    'Do not allow delete if not recognised

  End Select

  miUsageCount = List1.ListCount
  If piType = utlBatchJob Then
    CheckForUseage = (List1.ListCount > 0 And Not piType = utlBatchJob)
  Else
    CheckForUseage = (List1.ListCount > 0 And Not piType = utlReportPack)
  End If
  
  If miUsageCount = 0 Then
    List1.AddItem "<None>"
  End If

End Function


Private Sub GetNameWhereUsed(strSQL As String, Optional blnBatchJob As Boolean) 'As String
  
  Dim rsTemp As Recordset
  Dim blnHidden As Boolean
  Dim strName As String
  
  Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  Do While Not rsTemp.EOF
    
		blnHidden = (LCase(Trim(rsTemp!UserName)) <> LCase(Trim(datGeneral.UserNameForSQL)) _
								And rsTemp!Access = ACCESS_HIDDEN)
    
    If blnHidden Then
			strName = rsTemp(0) & ": <Hidden by " & StrConv(Trim(rsTemp!UserName), vbProperCase) & ">"
    Else
      If blnBatchJob Then
        strName = Right("  " & CStr(rsTemp.Fields("JobOrder").Value + 1), 3) & _
                  ". " & rsTemp.Fields("JobType").Value
        If rsTemp.Fields("JobID").Value > 0 Then
          strName = strName & ": '" & _
                  GetJobName(rsTemp.Fields("JobType").Value, rsTemp.Fields("JobID").Value) & _
                  "'"
        End If
      Else
        strName = rsTemp(0) & ": '" & rsTemp(1) & "'"
      End If
    End If
    
    List1.AddItem strName
    rsTemp.MoveNext
  Loop

  rsTemp.Close
  Set rsTemp = Nothing

End Sub


Private Sub CheckModuleSetup(piType As UtilityType, plngItemID As Long)
  
  Dim rsResults As ADODB.Recordset
  Dim strSQL As String
  Dim lngUtilType As Long
  
  lngUtilType = piType
  
  If lngUtilType < 0 Then Exit Sub
  
  strSQL = "SELECT COUNT(*) AS [result]" & _
    " FROM [ASRSysSSIntranetLinks]" & _
    " WHERE [ASRSysSSIntranetLinks].[utilityID] = " & CStr(plngItemID) & _
    "   AND [ASRSysSSIntranetLinks].[utilityType] = " & CStr(lngUtilType)
    
  Set rsResults = datGeneral.GetReadOnlyRecords(strSQL)

  If rsResults!Result > 0 Then
    List1.AddItem "Self-service intranet link"
  End If

  rsResults.Close
  Set rsResults = Nothing
  
End Sub

Private Sub CheckUserSettings(strSection As String, strValue As String, strName As String)
  
  Dim rsAllUserSetting As ADODB.Recordset
  Dim strSQL As String
  
  strSQL = vbNullString
  strSQL = strSQL & "SELECT DISTINCT [ASRSysUserSettings].[Username], "
  strSQL = strSQL & "                [ASRSysUserSettings].[Section]"
  strSQL = strSQL & "FROM   [ASRSysUserSettings] "
  strSQL = strSQL & "WHERE  [ASRSysUserSettings].[Section] = '" & strSection & "' "
  strSQL = strSQL & "   AND [ASRSysUserSettings].[SettingValue] = '" & strValue & "' "
  
  Set rsAllUserSetting = datGeneral.GetReadOnlyRecords(strSQL)
  
  With rsAllUserSetting
    Do While Not .EOF
      List1.AddItem strName & " for user '" & !userName & "'"
      .MoveNext
    Loop
    .Close
  End With

  Set rsAllUserSetting = Nothing
  
End Sub


Private Sub CheckSystemSettings(strIDs As String, strType As String, strName As String)
  
  Dim strReportType As String
  
  strReportType = Replace(strName, " ", "")
  If strName = "Stability" Then
    strName = "Stability Index"
  End If
  
  If strType = "X" Then   'Date Calc
    If GetSystemSetting(strReportType, "Custom Dates", "0") = "1" Then
      'MH20050411 Fault 9985
      'If InStr(strIDs, CStr(GetSystemSetting(strReportType, "Start Date", 0))) > 0 Or _
         InStr(strIDs, CStr(GetSystemSetting(strReportType, "End Date", 0))) > 0 Then
      If Trim(strIDs) = Trim(GetSystemSetting(strReportType, "Start Date", 0)) Or _
         Trim(strIDs) = Trim(GetSystemSetting(strReportType, "End Date", 0)) Then
            List1.AddItem strName
      End If
    End If
  Else
    If GetSystemSetting(strReportType, "Type", "A") = strType Then
      'MH20050411 Fault 9985
      'If InStr(strIDs, CStr(GetSystemSetting(strReportType, "ID", 0))) > 0 Then
      If Trim(strIDs) = Trim(GetSystemSetting(strReportType, "ID", 0)) Then
            List1.AddItem strName
      End If
    End If
  End If

End Sub



Private Function GetExprRootIDs(strID As String, Optional blnOrders As Boolean = False) As String

  Dim rsTemp As Recordset
  Dim objComp As clsExprComponent
  Dim strSQL As String


  If blnOrders Then
    strSQL = "SELECT ComponentID FROM ASRSysExprComponents " & _
             "WHERE FieldSelectionOrderID = " & strID
  Else
    'TM20010906 Fault 2778
    'Added clause to check for the Filters used Expressions.
    strSQL = "SELECT ComponentID FROM ASRSysExprComponents " & _
             "WHERE CalculationID = " & strID & " OR " & _
             "FilterID = " & strID & " OR " & _
             "(fieldSelectionFilter = " & strID & _
             " AND type = " & CStr(giCOMPONENT_FIELD) & ")"
  End If

  Set rsTemp = datGeneral.GetRecords(strSQL)
  With rsTemp
      
    GetExprRootIDs = vbNullString
    Do While Not .EOF
        
      Set objComp = New clsExprComponent
      objComp.ComponentID = !ComponentID

      GetExprRootIDs = GetExprRootIDs & _
          IIf(GetExprRootIDs <> vbNullString, ", ", vbNullString) & _
          CStr(objComp.RootExpressionID)

      .MoveNext
    Loop
    Set objComp = Nothing
    .Close
  
  End With
  Set rsTemp = Nothing
        
End Function


Private Sub GetRecordCount(utlType As UtilityType, lngID As Long)

  Dim rsTemp As Recordset
  Dim strSQL As String
  Dim objFilterExpr As clsExprExpression
  
  Dim strFilterCode As String
  Dim fOK As Boolean
  
  If utlType = utlFilter Then
    
    Set objFilterExpr = New clsExprExpression
    objFilterExpr.ExpressionID = lngID
    objFilterExpr.ConstructExpression
    
    fOK = objFilterExpr.RuntimeFilterCode(strFilterCode, True, False)
    If fOK = False Then
      txtRecCount = "<Access Denied>"
      Exit Sub
    End If
    
    strSQL = "SELECT COUNT(*) FROM " & _
             gcoTablePrivileges.Item(objFilterExpr.BaseTableName).RealSource & _
             " WHERE ID IN (" & strFilterCode & ")"
    Set rsTemp = datData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    
    txtRecCount = CStr(rsTemp(0).Value)
    
    rsTemp.Close
    Set rsTemp = Nothing
    Set objFilterExpr = Nothing
  
  
  ElseIf utlType = utlPicklist Then

    'This is wrong at the moment!!
    '(Need to check views and stuff)

    strSQL = "EXEC sp_ASRGetPickListRecords " & lngID
    Set rsTemp = datData.OpenRecordset(strSQL, adOpenKeyset, adLockReadOnly)

    txtRecCount = rsTemp.RecordCount
    
    rsTemp.Close
    Set rsTemp = Nothing

  
  End If

End Sub

Public Property Get UsageCount() As Integer
  'returns the amount of times this "thing" is used
  UsageCount = miUsageCount
End Property

Public Sub SetChangeTypeError()

' Used to display message saying that this "thing" cannot have it's returntype changed.
' Called from frmExpression
  mbChangeErrorMode = True

  frmCannotChangeType.Top = Frame1.Top
  frmCannotChangeType.Left = Frame1.Left
  frmCannotChangeType.Height = 1000
  frmCannotChangeType.Width = Frame1.Width
  
  frmCannotChangeType.Visible = True
  Frame1.Visible = False
  Frame2.Top = frmCannotChangeType.Top + frmCannotChangeType.Height + 500

  ' Set the caption
  lblCannotChangeType.Caption = "The return type cannot be changed, currently being used in... "

  ' Fix the border style
  Me.BorderStyle = vbFixedDialog

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

