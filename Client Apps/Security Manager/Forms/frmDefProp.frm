VERSION 5.00
Begin VB.Form frmDefProp 
   Caption         =   "Definition Properties"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8027
   Icon            =   "frmDefProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmCannotChangeType 
      Height          =   840
      Left            =   105
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   4785
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   195
         Picture         =   "frmDefProp.frx":000C
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblCannotChangeType 
         Caption         =   "Cannot Change Type Text"
         Height          =   375
         Left            =   750
         TabIndex        =   15
         Top             =   420
         Width           =   3930
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current Usage :"
      Height          =   2100
      Left            =   120
      TabIndex        =   10
      Top             =   2440
      Width           =   6045
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
      Left            =   4965
      TabIndex        =   8
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   2410
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   6045
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

Private mblnUsage As Boolean
Private miUsageCount As Integer
Private mbChangeErrorMode As Boolean

Private Const MIN_FORM_HEIGHT = 6000
Private Const MIN_FORM_WIDTH = 6500

Public Property Get UtilName() As String
  UtilName = txtName.Text
End Property

Public Property Let UtilName(ByVal strNewValue As String)
  txtName.Text = strNewValue
End Property


Private Function FormatText(varDate As Variant, varUser As Variant) As String
  If IsNull(varUser) Then
    FormatText = "<Unknown>"
  Else
    FormatText = Format(varDate, DateFormat & " hh:nn") & _
                 "  by  " & StrConv(varUser, vbProperCase)
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
  
  mbChangeErrorMode = False
    
  Hook Me.hWnd, MIN_FORM_WIDTH, MIN_FORM_HEIGHT
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
  
  If Me.Width < 5000 Then Me.Width = 5000
  
  If Frame2.Visible = False And mblnUsage = False Then
    If Me.Height <> 3120 Then Me.Height = 3120
  Else
    If Me.Height < 4500 Then Me.Height = 4500
  End If
  
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

Public Sub PopulateUtil(piExprType As ExpressionTypes, lngID As Long)
  
  Call GetData(piExprType, lngID)
  Call DrawControls(piExprType)

End Sub

Private Sub GetData(piExprType As ExpressionTypes, lngID As Long)

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String

  strSQL = "SELECT * FROM ASRSysUtilAccessLog " & _
           "WHERE UtilID = " & CStr(lngID)

  Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

  With rsTemp
    
    If Not .BOF And Not .EOF Then
      txtCreated = FormatText(!CreatedDate, !CreatedBy)
      txtSaved = FormatText(!SavedDate, !SavedBy)
      txtRun = FormatText(!RunDate, !RunBy)
    Else
      txtCreated = "<Unknown>"
      txtSaved = "<Unknown>"
      txtRun = "<Unknown>"
    End If

  End With
  
  rsTemp.Close
  Set rsTemp = Nothing

End Sub


Private Sub DrawControls(piExprType As ExpressionTypes)

  Dim blnUsage As Boolean
  Dim blnLastRun As Boolean
  Dim blnRecCount As Boolean
  Dim lngOffset As Long
  
  ' JDM - 26/11/01 - Fault 3201 - Disable the last run box
  blnLastRun = False
  blnRecCount = False
  
  blnUsage = True
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

Public Function CheckForUseage(psType As String, plngItemID As Long) As Boolean
  ' NEWACCESS - needs to be updated as each report/utility is updated for the new access.
  
  Dim strSQL As String
  Dim strID As String
  Dim strRootIDs As String
  
  strID = CStr(plngItemID)
  List1.Clear

  Select Case Trim(UCase(psType))

  Case "FILTER"

    strRootIDs = GetExprRootIDs(strID)
    
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
        " WHERE ASRSysCrossTab.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Cross Tab'," & _
        " ASRSysCrossTab.Name," & _
        " ASRSysCrossTab.UserName," & _
        " ASRSysCrossTabAccess.Access" & _
        " FROM ASRSysCrossTab" & _
        " INNER JOIN ASRSysCrossTabAccess ON ASRSysCrossTab.crossTabID = ASRSysCrossTabAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCrossTabAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysCrossTab.FilterID = " & strID)
    End If
    
    'Custom Report Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " ASRSysCustomReportsName.Name," & _
        " ASRSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " WHERE ASRSysCustomReportsName.Filter = " & strID & _
        " OR ASRSysCustomReportsName.Parent1Filter = " & strID & _
        " OR ASRSysCustomReportsName.Parent2Filter = " & strID & _
        " OR ASRSYSCustomReportsChildDetails.ChildFilter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " ASRSysCustomReportsName.Name," & _
        " ASRSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM ASRSysCustomReportsName" & _
        " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
        " INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysCustomReportsName.Filter = " & strID & _
        " OR ASRSysCustomReportsName.Parent1Filter = " & strID & _
        " OR ASRSysCustomReportsName.Parent2Filter = " & strID & _
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
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysDataTransferName.FilterID = " & strID)
    End If
    
    'Export Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " ASRSysExportName.Name," & _
        " ASRSysExportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysExportName" & _
        " WHERE ASRSysExportName.Filter = " & strID & " OR ASRSysExportName.ChildFilter = " & strID & _
        " OR ASRSysExportName.Parent1Filter = " & strID & " OR ASRSysExportName.Parent2Filter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Export'," & _
        " ASRSysExportName.Name," & _
        " ASRSysExportName.UserName," & _
        " ASRSysExportAccess.Access" & _
        " FROM ASRSysExportName" & _
        " INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysExportName.Filter = " & strID & " OR ASRSysExportName.ChildFilter = " & strID & _
        " OR ASRSysExportName.Parent1Filter = " & strID & " OR ASRSysExportName.Parent2Filter = " & strID)
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
    
    'Global Function Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN type = 'A' THEN 'Global Add'" & _
        "   WHEN type = 'D' THEN 'Global Delete'" & _
        "   WHEN type = 'U' THEN 'Global Update'" & _
        "   ELSE 'Global Function'" & _
        "   END," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysGlobalFunctions" & _
        " WHERE ASRSysGlobalFunctions.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN type = 'A' THEN 'Global Add'" & _
        "   WHEN type = 'D' THEN 'Global Delete'" & _
        "   WHEN type = 'U' THEN 'Global Update'" & _
        "   ELSE 'Global Function'" & _
        "   END," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " ASRSysGlobalAccess.Access" & _
        " FROM ASRSysGlobalFunctions" & _
        " INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysGlobalAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysGlobalFunctions.FilterID = " & strID)
    End If
    
    'Mail Merge/Envelopes & Labels Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN isLabel = 1 THEN 'Envelopes & Labels'" & _
        "   ELSE 'Mail Merge'" & _
        "   END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " WHERE ASRSysMailMergeName.FilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN isLabel = 1 THEN 'Envelopes & Labels'" & _
        "   ELSE 'Mail Merge'" & _
        "   END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysMailMergeName.FilterID = " & strID)
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
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
        " WHERE ASRSysRecordProfileName.FilterID = " & strID & _
        " OR ASRSYSRecordProfileTables.FilterID = " & strID)
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
        " INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE AsrSysCalendarReports.filter = " & strID & _
        " OR ASRSysCalendarReportEvents.filterID = " & strID)
    End If
  
    'Match Report/Succession Planning/Career Progression Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN matchReportType = 1 THEN 'Succession Planning'" & _
        "   WHEN matchReportType = 2 THEN 'Career Progression'" & _
        "   ELSE 'Match Report'" & _
        "   END," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMatchReportName" & _
        " WHERE ASRSysMatchReportName.Table1Filter = " & strID & _
        " OR ASRSysMatchReportName.Table2Filter = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN matchReportType = 1 THEN 'Succession Planning'" & _
        "   WHEN matchReportType = 2 THEN 'Career Progression'" & _
        "   ELSE 'Match Report'" & _
        "   END," & _
        " ASRSysMatchReportName.Name," & _
        " ASRSysMatchReportName.UserName," & _
        " ASRSysMatchReportAccess.Access" & _
        " FROM ASRSysMatchReportName" & _
        " INNER JOIN ASRSysMatchReportAccess ON ASRSysMatchReportName.matchReportID = ASRSysMatchReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMatchReportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysMatchReportName.Table1Filter = " & strID & _
        " OR ASRSysMatchReportName.Table2Filter = " & strID)
    End If
  
    'Talent Reports
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Talent Report'," & _
        " n.Name, n.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysTalentReports n" & _
        " WHERE n.BaseFilterID = " & strID & " OR n.MatchFilterID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Talent Report'," & _
        " n.Name, n.UserName, a.Access" & _
        " FROM ASRSysCrossTab n" & _
        " INNER JOIN ASRSysCrossTabAccess a ON n.ID = a.ID" & _
        " INNER JOIN sysusers b ON a.groupname = b.name AND b.name = '" & gsUserGroup & "'" & _
        " WHERE n.BaseFilterID = " & strID & " OR n.MatchFilterID = " & strID)
    End If
  
  Case "CALCULATION"
    
    strRootIDs = GetExprRootIDs(strID)

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
        " INNER JOIN ASRSysCalendarReportAccess ON ASRSysCalendarReportAccess.ID = ASRSysCalendarReports.ID" & _
        " INNER JOIN sysusers b ON ASRSysCalendarReportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE AsrSysCalendarReports.DescriptionExpr = " & strID & _
        "   OR AsrSysCalendarReports.StartDateExpr = " & strID & _
        "   OR AsrSysCalendarReports.EndDateExpr = " & strID)
    End If
    
    'Custom Report Calculation
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " ASRSysCustomReportsName.Name," & _
        " ASRSysCustomReportsName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = AsrSysCustomReportsDetails.CustomReportID" & _
        " WHERE UPPER(ASRSysCustomReportsDetails.type) = 'E' AND ASRSysCustomReportsDetails.colExprID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT 'Custom Report'," & _
        " ASRSysCustomReportsName.Name," & _
        " ASRSysCustomReportsName.UserName," & _
        " ASRSysCustomReportAccess.Access" & _
        " FROM ASRSysCustomReportsName" & _
        " INNER JOIN ASRSysCustomReportsDetails ON ASRSysCustomReportsName.ID = AsrSysCustomReportsDetails.CustomReportID" & _
        " INNER JOIN ASRSysCustomReportAccess ON ASRSysCustomReportsName.ID = ASRSysCustomReportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysCustomReportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
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
        " INNER JOIN ASRSysExportName ON ASRSysExportDetails.exportID = ASRSysExportName.ID " & _
        " INNER JOIN ASRSysExportAccess ON ASRSysExportName.ID = ASRSysExportAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysExportAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
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

    'Global Function calc
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN type = 'A' THEN 'Global Add'" & _
        "   WHEN type = 'D' THEN 'Global Delete'" & _
        "   WHEN type = 'U' THEN 'Global Update'" & _
        "   ELSE 'Global Function'" & _
        "   END," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysGlobalItems" & _
        " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID " & _
        " WHERE ASRSysGlobalItems.ValueType = 4 AND ASRSysGlobalItems.ExprID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN type = 'A' THEN 'Global Add'" & _
        "   WHEN type = 'D' THEN 'Global Delete'" & _
        "   WHEN type = 'U' THEN 'Global Update'" & _
        "   ELSE 'Global Function'" & _
        "   END," & _
        " ASRSysGlobalFunctions.Name," & _
        " ASRSysGlobalFunctions.UserName," & _
        " ASRSysGlobalAccess.Access" & _
        " FROM ASRSysGlobalItems" & _
        " INNER JOIN ASRSysGlobalFunctions ON ASRSysGlobalItems.functionID = ASRSysGlobalFunctions.functionID " & _
        " INNER JOIN ASRSysGlobalAccess ON ASRSysGlobalFunctions.functionID = ASRSysGlobalAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysGlobalAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE ASRSysGlobalItems.ValueType = 4 AND ASRSysGlobalItems.ExprID = " & strID)
    End If
    
    'Mail Merge/Envelopes & Labels Filter
    If gfCurrentUserIsSysSecMgr Then
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN isLabel = 1 THEN 'Envelopes & Labels'" & _
        "   ELSE 'Mail Merge'" & _
        "   END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " '" & ACCESS_READWRITE & "' AS access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " WHERE upper(AsrSysMailMergeColumns.Type) = 'E' " & _
        "   AND AsrSysMailMergeColumns.ColumnID = " & strID)
    Else
      Call GetNameWhereUsed( _
        "SELECT DISTINCT CASE WHEN isLabel = 1 THEN 'Envelopes & Labels'" & _
        "   ELSE 'Mail Merge'" & _
        "   END," & _
        " ASRSysMailMergeName.Name," & _
        " ASRSysMailMergeName.UserName," & _
        " ASRSysMailMergeAccess.Access" & _
        " FROM ASRSysMailMergeName" & _
        " INNER JOIN AsrSysMailMergeColumns ON AsrSysMailMergeName.mailMergeID = AsrSysMailMergeColumns.mailMergeID" & _
        " INNER JOIN ASRSysMailMergeAccess ON ASRSysMailMergeName.mailMergeID = ASRSysMailMergeAccess.ID" & _
        " INNER JOIN sysusers b ON ASRSysMailMergeAccess.groupname = b.name" & _
        " INNER JOIN sysusers a ON ( b.uid = a.gid" & _
        "   AND a.Name = system_user)" & _
        " WHERE upper(AsrSysMailMergeColumns.Type) = 'E' " & _
        "   AND AsrSysMailMergeColumns.ColumnID = " & strID)
    End If
    
  Case "ORDER"

    strRootIDs = GetExprRootIDs(strID, True)

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

    'Default Order
    'JPD 20050812 Fault 10166
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Default Order', tableName " & _
      ",'RW' As Access, '" & UCase(LTrim(Replace(gsUserName, "'", "''"))) & "' As Username " & _
      "FROM ASRSysTables " & _
      " WHERE defaultOrderID = " & strID)

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
    
    'Screen Order
    'JPD 20050812 Fault 10166
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Screen Order', ASRSysScreens.name " & _
      ", 'RW' As Access, '" & UCase(LTrim(Replace(gsUserName, "'", "''"))) & "' As Username " & _
      "FROM ASRSysScreens " & _
      " WHERE ASRSysScreens.OrderID = " & strID)
    
    'Module Setup Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Module Setup', " & _
        "CASE WHEN ASRSysModuleSetup.ModuleKey = 'MODULE_TRAININGBOOKING' " & "THEN 'Training Booking'" & _
          "WHEN ASRSysModuleSetup.ModuleKey = 'MODULE_PERSONNEL' " & "THEN 'Personnel'" & _
          "WHEN ASRSysModuleSetup.ModuleKey = 'MODULE_ABSENCE' " & "THEN 'Absence'" & _
          "ELSE '<unknown>' END" & _
        ",'" & ACCESS_READWRITE & "' As Access, '' As Username " & _
      "FROM ASRSysModuleSetup " & _
      " WHERE parameterType = '" & gsPARAMETERTYPE_ORDERID & "'" & _
      " AND parameterValue = '" & strID & "'")

    'Record Profile Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Record Profile'," & _
      " ASRSysRecordProfileName.Name," & _
      " ASRSysRecordProfileName.UserName," & _
      " '" & ACCESS_READWRITE & "' AS access" & _
      " FROM ASRSysRecordProfileName" & _
      " LEFT OUTER JOIN ASRSYSRecordProfileTables ON ASRSysRecordProfileName.recordProfileID = ASRSYSRecordProfileTables.recordProfileID" & _
      " WHERE ASRSysRecordProfileName.OrderID = " & strID & _
      " OR ASRSYSRecordProfileTables.OrderID = " & strID)

    'Custom Report Order
    Call GetNameWhereUsed( _
      "SELECT DISTINCT 'Custom Report'," & _
      " AsrSysCustomReportsName.Name," & _
      " AsrSysCustomReportsName.UserName," & _
      " '" & ACCESS_READWRITE & "' AS access" & _
      " FROM AsrSysCustomReportsName" & _
      " LEFT OUTER JOIN ASRSYSCustomReportsChildDetails ON ASRSysCustomReportsName.ID = ASRSYSCustomReportsChildDetails.customReportID" & _
      " WHERE (ASRSYSCustomReportsChildDetails.ChildOrder = " & strID & ")")

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
    
  Case "BATCH JOB"
    'Do nothing - no usage !
    
  Case Else
    List1.AddItem "<Error Checking Usage>"    'Do not allow delete if not recognised

  End Select

  CheckForUseage = (List1.ListCount > 0)
  miUsageCount = List1.ListCount

  If CheckForUseage = False Then
    List1.AddItem "<None>"
  End If

End Function


Private Sub GetNameWhereUsed(strSQL As String) 'As String
  
  Dim rsTemp As ADODB.Recordset
  Dim blnHidden As Boolean
  Dim strName As String
  
  Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
  
  Do While Not rsTemp.EOF
    
    blnHidden = (LCase(Trim(rsTemp!UserName)) <> LCase(Trim(gsUserName)) _
                And rsTemp!Access = ACCESS_HIDDEN)
    
    If blnHidden Then
      strName = "<Hidden by " & StrConv(Trim(rsTemp!UserName), vbProperCase) & ">"
    Else
      strName = "'" & rsTemp(1) & "'"
    End If
    
    List1.AddItem rsTemp(0) & ": " & strName
    rsTemp.MoveNext
  Loop

  rsTemp.Close
  Set rsTemp = Nothing

End Sub


Private Function GetExprRootIDs(strID As String, Optional blnOrders As Boolean = False) As String

  Dim rsTemp As ADODB.Recordset
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

  Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
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


Private Sub GetRecordCount(piExpressionType As ExpressionTypes, lngID As Long)

  Dim rsTemp As ADODB.Recordset
  Dim strSQL As String
  Dim objFilterExpr As clsExprExpression
  
  Dim strFilterCode As String
  Dim fOK As Boolean
  
  If piExpressionType = giEXPR_RUNTIMEFILTER Then
    
    Set objFilterExpr = New clsExprExpression
    objFilterExpr.ExpressionID = lngID
    objFilterExpr.ConstructExpression
    
    fOK = objFilterExpr.RuntimeFilterCode(strFilterCode, True, False)
    If fOK = False Then
      txtRecCount = "<Access Denied>"
      Exit Sub
    End If
  
    strSQL = "SELECT COUNT(*) FROM " & _
             objFilterExpr.BaseTableName & _
             " WHERE ID IN (" & strFilterCode & ")"
    Set rsTemp = modExpression.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)
    
    txtRecCount = CStr(rsTemp(0).Value)
    
    rsTemp.Close
    Set rsTemp = Nothing
    Set objFilterExpr = Nothing
 
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
