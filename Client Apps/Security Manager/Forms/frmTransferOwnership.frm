VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmTransferOwnership 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ownership Transfer"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8022
   Icon            =   "frmTransferOwnership.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "&Transfer"
      Height          =   405
      Left            =   4290
      TabIndex        =   5
      Top             =   3180
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   5595
      TabIndex        =   7
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Frame fraOwnerOf 
      Caption         =   "Owner of : "
      Height          =   1995
      Left            =   90
      TabIndex        =   6
      Top             =   1080
      Width           =   6705
      Begin VB.TextBox txtUtil 
         Height          =   1560
         Index           =   0
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmTransferOwnership.frx":000C
         Top             =   285
         Width           =   6435
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdUtil 
         Height          =   1560
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   2820
         Width           =   4935
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         GroupHeadLines  =   0
         Col.Count       =   0
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
         UseExactRowCount=   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         RowNavigation   =   1
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         TabNavigation   =   1
         _ExtentX        =   8705
         _ExtentY        =   2752
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
   Begin VB.Frame fraUtilityOwnership 
      Caption         =   "Transfer ownership of all reports, tools and utilities :"
      Height          =   825
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   6720
      Begin VB.ComboBox cboTo 
         Height          =   315
         Left            =   3975
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2640
      End
      Begin VB.ComboBox cboFrom 
         Height          =   315
         Left            =   765
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   195
         Left            =   3525
         TabIndex        =   4
         Top             =   390
         Width           =   285
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   390
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmTransferOwnership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnLoading As Boolean
Private mbCancelled As Boolean
Private marySelectedUsers() As Variant
Public gfCanDeleteAll As Boolean
Private mstrConnectString As String

Public Property Let Cancelled(bNewValue As Boolean)
  mbCancelled = bNewValue
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Function Initialise(vAry() As Variant) As Boolean
Dim fNothingToTransfer As Boolean
  On Error GoTo ErrorTrap

  Initialise = True
  mblnLoading = True
  
  marySelectedUsers = vAry
  
  Me.Cancelled = Not PopulateToCombo

  If Not Me.Cancelled Then
    'TM20011107 Fault 3105 - Change mouse pointer as this could be a lengthy process.
    Screen.MousePointer = vbHourglass
    Me.Cancelled = Not PopulateReports
    Screen.MousePointer = vbNormal
    
    'TM06092004 Fault 9125 fixed
    If cboTo.ListCount > 0 Then
      cboTo.ListIndex = 0
    End If
    If cboFrom.ListCount > 0 Then
      cboFrom.ListIndex = 0
    End If
    
    fNothingToTransfer = TransferedStatus
  End If
  
  ' NHRD16072003 Fault 4369
  ' No point displaying the TransferOwnership form if is has
  ' been cancelled or the users have no utilities to transfer
  Initialise = (Me.Cancelled = False) And (fNothingToTransfer = False)

TidyUpAndExit:
  mblnLoading = False
  Exit Function

ErrorTrap:
  Initialise = False
  MsgBox "Error whilst initialising the Utility Ownership form." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Function Selected(sUser As String) As Boolean

  Dim i As Integer
  
  On Error GoTo ErrorTrap
  
  Selected = False
  
  For i = 1 To UBound(marySelectedUsers) Step 1
    If sUser = marySelectedUsers(i) Then
      Selected = True
    End If
  Next i

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error whilst validating selected users." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Function TransferedStatus() As Boolean
' NHRD16072003 Fault 4369
' Changed this Sub to a function as we can use this boolean
' as a flag to show the TransferOwnership form (or not)
  Dim i As Integer
  Dim bAllTransfered As Boolean

  On Error GoTo ErrorTrap
  
  TransferedStatus = True
  TransferedStatus = NothingToTransfer
   
  If TransferedStatus Then
    cboTo.Enabled = False
    cboTo.Visible = False
    lblTo.Visible = False
    lblTo.Enabled = False
    cmdTransfer.Enabled = False
    cmdTransfer.Visible = False
    cmdCancel.Caption = "&Finish"
    cmdCancel.Visible = True
    cmdCancel.Enabled = True
    cmdCancel.Tag = "F"
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error whilst validating transfer status." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Function PopulateReports() As Boolean

  Dim i As Integer
  Dim iFromCount As Integer
  Dim sTempUser As String

  On Error GoTo ErrorTrap

  iFromCount = 0
  
  If UBound(marySelectedUsers) > 1 Then
    If WriteReport("<All>", iFromCount) Then
      If mblnLoading Then
        cboFrom.AddItem "<All>", iFromCount
      End If
      iFromCount = iFromCount + 1
    End If
  End If
  
  ' Now add the selected users to the From combo box.
  For i = 1 To UBound(marySelectedUsers) Step 1
    sTempUser = marySelectedUsers(i)
    If WriteReport(sTempUser, iFromCount) Then
      If mblnLoading Then
        cboFrom.AddItem sTempUser, iFromCount
      End If
      iFromCount = iFromCount + 1
    End If
  Next i

  If iFromCount = 1 Then
    cboFrom.Enabled = False
    'TM20020215 Fault 3515
    cboFrom.BackColor = vbButtonFace
  Else
    cboFrom.Enabled = True
    cboFrom.BackColor = vbWindowBackground
  End If
  
  If iFromCount >= UBound(marySelectedUsers) Then
    PopulateReports = True
  Else
    PopulateReports = False
  End If
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error whilst populating the utility reports." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Function

Private Function WriteReport(sUserName As String, iFromIndex As Integer) As Boolean

  Dim rsTemp As New ADODB.Recordset
  Dim sCurrentType As String
  Dim bAll As Boolean
  Dim i As Integer
  Dim s As String
  Dim tStr As String
  
  On Error GoTo ErrorTrap

  sCurrentType = vbNullString
  bAll = (sUserName = "<All>")
  
  If mblnLoading Then
    If iFromIndex > 0 Then
      Load Me.txtUtil(iFromIndex)
      
      '******************************
      ' Grid Control Code.          *
      '******************************
      'Load Me.grdUtil(iFromIndex)
      '******************************
    End If
  End If

  With Me.txtUtil(iFromIndex)
    'Connect mstrConnectString
    rsTemp.Open BuildReportSQL(sUserName), gADOCon, adOpenStatic, adLockReadOnly
      .Enabled = True
      .Locked = True
      .Text = vbNullString
    
    If rsTemp.RecordCount > 0 Then
      Do While Not rsTemp.EOF
        If rsTemp!Type <> sCurrentType And sCurrentType <> vbNullString Then
          .Text = .Text & vbCrLf
          .Text = .Text & vbCrLf
        ElseIf sCurrentType <> vbNullString Then
          .Text = .Text & vbCrLf
        End If
        sCurrentType = IIf(IsNull(rsTemp!Type), "Other", rsTemp!Type)
        
        tStr = StrConv(sCurrentType, VbStrConv.vbProperCase) & ":  "
        i = InStr(i + 1, tStr, "-")
        Do While i > 0 And i < Len(tStr)
          Mid$(tStr, i + 1, 1) = UCase(Mid$(tStr, i + 1, 1))
          i = InStr(i + 1, tStr, "-")
        Loop
        .Text = .Text & tStr
        
        .Text = .Text & "'" & StrConv(rsTemp!Name, vbProperCase) & "'"
        If bAll Then .Text = .Text & " (" & rsTemp!UserName & ")"
        rsTemp.MoveNext
      Loop
    Else
      .Text = "<None>"
    End If
    
    rsTemp.Close
  End With

'********************************************************************************
' The following code is for displaying the reports in grid form.                *
' Un-comment this code and reposition the grid control grdUtil to display the   *
' reports in a grid format.                                                     *
'********************************************************************************
'  With Me.grdUtil(iFromIndex)
'      Connect ("Driver={SQL Server};Server=" & gsSQLServerName & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";Database=" & gsDatabaseName & ";")
'
'      rsTemp.Open BuildReportSQL(sUserName), gADOCon, adOpenStatic, adLockReadOnly
'
'      .Enabled = True
'      .Redraw = False
'      .Columns.RemoveAll
'      .FieldSeparator = vbTab
'
'      For i = 0 To (rsTemp.Fields.Count - 1)
'        .Columns.Add i
'        .Columns(i).Name = rsTemp.Fields(i).Name
'
'        If (UCase(rsTemp.Fields(i).Name) <> "ID") And (Left(rsTemp.Fields(i).Name, 1) <> "?") Then
'          If (UCase(rsTemp.Fields(i).Name) = "USERNAME") And Not bAll Then
'            .Columns(i).Visible = False
'          Else
'            .Columns(i).Visible = True
'          End If
'        Else
'          .Columns(i).Visible = False
'        End If
'
'        .Columns(i).Caption = rsTemp.Fields(i).Name
'      Next i
'
'      If rsTemp.RecordCount > 0 Then
'        rsTemp.MoveFirst
'        Do While Not rsTemp.EOF
'          s = s & rsTemp!ID & vbTab
'          s = s & rsTemp!Type & vbTab
'          s = s & rsTemp!Name & vbTab
'          s = s & rsTemp!UserName & vbTab
'          .AddItem s
'          s = vbNullString
'          rsTemp.MoveNext
'        Loop
'
'      End If
'      .ReBind
'      .Redraw = True
'
'      rsTemp.Close
'  End With
'
'********************************************************************************
  
  WriteReport = True

TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
ErrorTrap:
  WriteReport = False
  MsgBox "Error writing report for " & sUserName & "." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Function BuildReportSQL(sUserName As String) As String

  Dim sSQL As String
  Dim sWhere As String
  Dim bAll As Boolean
  Dim i As Integer
  Dim sWhereExtra As String
  
  On Error GoTo ErrorTrap
  
  sSQL = vbNullString
  bAll = (sUserName = "<All>")
 
  If bAll Then
    sWhere = vbNullString
    For i = 1 To UBound(marySelectedUsers) Step 1
      
      'MH20061207 Fault 11767
      'sWhere = sWhere & _
        IIf(sWhere <> vbNullString, " OR ", vbNullString) & _
        "LOWER(Username) = '" & LCase(marySelectedUsers(i)) & "'" & vbCrLf
      sWhere = sWhere & _
        IIf(sWhere <> vbNullString, " OR ", vbNullString) & _
        "LOWER(Username) = '" & Replace(LCase(marySelectedUsers(i)), "'", "''") & "'" & vbCrLf
    
    Next i
    sWhere = "WHERE (" & sWhere & ")"
  Else
    sWhere = "WHERE LOWER(Username) = '" & Replace(LCase(sUserName), "'", "''") & "' " & vbCrLf
  End If
  
  sSQL = sSQL & "SELECT  ID, 'ORGANISATION REPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysOrganisationReport " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  ID, 'TALENT REPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysTalentReports " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  CrossTabID AS ID, '9-BOX GRID REPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysCrossTab " & vbCrLf
  sSQL = sSQL & sWhere & " AND CrossTabType = " & ctt9GridBox & vbCrLf
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  ID, CASE IsBatch WHEN 1 THEN 'BATCH JOB' WHEN 0 THEN 'REPORT PACK' END AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From  ASRSysBatchJobName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  ID, 'CUSTOM REPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysCustomReportsName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  CrossTabID AS ID, 'CROSS TAB' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysCrossTab " & vbCrLf
  sSQL = sSQL & sWhere & " AND CrossTabType <> " & ctt9GridBox & vbCrLf
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  MailMergeID AS ID, CASE IsLabel WHEN 1 THEN 'ENVELOPE & LABEL' ELSE 'MAIL MERGE' END AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysMailMergeName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  FunctionID AS ID, CASE Type WHEN 'A' THEN 'GLOBAL ADD' WHEN 'D' THEN 'GLOBAL DELETE' WHEN 'U' THEN 'GLOBAL UPDATE' END AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysGlobalFunctions " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  DataTransferID AS ID, 'DATA TRANSFER' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysDataTransferName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  ID, 'IMPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysImportName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  ID, 'EXPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysExportName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  
  'NPG20071122 Fault 12441
  sSQL = sSQL & "SELECT  ExprID AS ID, CASE Type WHEN " & ExpressionTypes.giEXPR_RUNTIMEFILTER & " THEN 'FILTER' " & _
                                               " WHEN " & ExpressionTypes.giEXPR_RUNTIMECALCULATION & " THEN 'CALCULATION' " & _
                                               " WHEN " & ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC & " THEN 'REPORT CALCULATION' " & _
                                     " END AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysExpressions " & vbCrLf
  
  'TM20010904 Fault 1808
  'Only select expressions that are used in the Data Manager.
  'i.e. RUNTIME_FILTERS OR RUNTIME_CALCULATIONS.
  'NPG20071122 Fault 12441 (include record independant calcs too)
  sSQL = sSQL & sWhere & " AND (Type = " & ExpressionTypes.giEXPR_RUNTIMEFILTER & _
                           " OR Type = " & ExpressionTypes.giEXPR_RUNTIMECALCULATION & _
                           " OR Type = " & ExpressionTypes.giEXPR_RECORDINDEPENDANTCALC & ") " & vbCrLf
  
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  PickListID AS ID, 'PICKLIST' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysPickListName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  recordProfileID, 'RECORD PROFILE' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysRecordProfileName " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT ID, 'CALENDAR REPORT' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysCalendarReports " & vbCrLf
  sSQL = sSQL & sWhere
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT MatchReportID AS ID, CASE matchReportType WHEN 0 THEN 'MATCH REPORT' WHEN 1 THEN 'SUCCESSION PLANNING' WHEN 2 THEN 'CAREER PROGRESSION' END AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysMatchReportName " & vbCrLf
  sSQL = sSQL & sWhere
  
  'MH20030529 Fault 5705
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  EmailGroupID, 'EMAIL GROUP' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysEmailGroupName " & vbCrLf
  sSQL = sSQL & sWhere

  'JPD 20030730 Fault 6420
  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  LabelTypeID, 'ENVELOPE & LABEL TEMPLATE' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysLabelTypes " & vbCrLf
  sSQL = sSQL & sWhere

  sSQL = sSQL & "Union " & vbCrLf
  sSQL = sSQL & "SELECT  DocumentMapID, 'DOCUMENT TYPE' AS Type, Name, Username " & vbCrLf
  sSQL = sSQL & "From ASRSysDocumentManagementTypes " & vbCrLf
  sSQL = sSQL & sWhere

  sSQL = sSQL & "ORDER BY Type, Name " & vbCrLf

  BuildReportSQL = sSQL

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  MsgBox "Error whilst creating report SQL string." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Function

Private Sub ShowUtilities(Index As Integer)

  Dim iCount  As Integer

  On Error GoTo ErrorTrap

  For iCount = 0 To txtUtil.Count - 1 Step 1
    If Index = iCount Then
      txtUtil(iCount).Visible = True
    Else
      txtUtil(iCount).Visible = False
    End If
  Next iCount

'**************************************************************
' Grid Control Code.                                          *
'**************************************************************
'  For iCount = 0 To grdUtil.Count - 1 Step 1
'    If Index = iCount Then
'      grdUtil(iCount).Visible = True
'    Else
'      grdUtil(iCount).Visible = False
'    End If
'  Next iCount
'**************************************************************

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error whilst showing the utility reports." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit

End Sub

Private Sub cboFrom_Click()
  cboFrom.Tag = cboFrom.Text
  ShowUtilities cboFrom.ListIndex
  If cboFrom.Text <> "<All>" Then
    Me.fraOwnerOf.Caption = "'" & cboFrom.Text & "'" & " is owner of : "
  Else
    Me.fraOwnerOf.Caption = "Reports, tools && utilities to be transferred : "
  End If
End Sub

Private Sub cboTo_Click()
  cboTo.Tag = cboTo.Text
End Sub

Private Sub cmdCancel_Click()
  
  Unload Me
  
End Sub

Private Sub cmdTransfer_Click()

  If txtUtil(cboFrom.ListIndex).Text <> "<None>" Then
    If ValidateSelection Then
      DoTransfer
      PopulateReports
    End If
    
    TransferedStatus
  Else
    MsgBox "The user '" & cboFrom.Text & "' does not own any utilities.", vbInformation + vbOKOnly, App.Title
  End If
  Me.SetFocus

End Sub

Private Function ValidateSelection() As Boolean

  Dim sMessage As String
  Dim bAll As String
  
  On Error GoTo ErrorTrap

  bAll = (cboFrom.Text = "<All>")
  
  If bAll Then
    sMessage = "You are about to transfer the ownership of all the utilities owned by all the selected users to '" & cboTo.Text & "'." & vbCrLf & _
                  "This action cannot be undone. Are you sure you wish to continue ?"
  Else
    sMessage = "You are about to transfer the ownership of all the utilities owned by '" & cboFrom.Text & "' to '" & cboTo.Text & "'." & vbCrLf & _
                "This action cannot be undone. Are you sure you wish to continue ?"
  End If
  
  If MsgBox(sMessage, vbYesNo + vbQuestion, App.Title) = vbNo Then
    ValidateSelection = False
    Me.Visible = True
    GoTo TidyUpAndExit
  End If

  ValidateSelection = True

TidyUpAndExit:
  Exit Function

ErrorTrap:
  MsgBox "Error whilst validating selection." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  ValidateSelection = False
  Resume TidyUpAndExit
  
End Function

Private Sub Progress(strCaption As String)

  On Error GoTo ErrorTrap

  With gobjProgress
    .Bar1Caption = strCaption
    .UpdateProgress False
  End With

TidyUpAndExit:
  Exit Sub
  
ErrorTrap:
  MsgBox "Error setting the status of the progress bar." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Sub

Private Function DoTransfer() As Boolean

  On Error GoTo ErrorTrap

  Dim strCommand As String
  Dim strFrom As String
  Dim strTo As String
  Dim blnAll As Boolean
  Dim sWhere As String
  Dim sWhereExtra As String
  Dim i As Integer
  
  With gobjProgress
    '.AviFile = "" 'App.Path & "\videos\table.Avi"
    .AVI = dbXferOwnership
    .Caption = "Ownership Transfer..."
    .NumberOfBars = 1
    .Bar1MaxValue = 14
    .Time = False
    .Cancel = True
    .OpenProgress
  End With
  
  ' Set variables
  Progress "Setting Variables..."
  strFrom = Replace(cboFrom.Text, "'", "''")
  strTo = Replace(cboTo.Text, "'", "''")
  blnAll = (strFrom = "<All>")
  If blnAll Then
    sWhere = " WHERE  "
    For i = 1 To UBound(marySelectedUsers) Step 1
      If i = 1 Then
        sWhere = sWhere & " LOWER(Username) = '" & LCase(Replace(marySelectedUsers(i), "'", "''")) & "' " & vbCrLf
      Else
        sWhere = sWhere & " OR LOWER(Username) = '" & LCase(Replace(marySelectedUsers(i), "'", "''")) & "' " & vbCrLf
      End If
    Next i
  Else
    sWhere = " WHERE   LOWER(Username) = '" & strFrom & "' " & vbCrLf
  End If
  DoEvents
  
  ' Initialise the gADOCon connection
  Progress "Initialising Connection..."
  
  'Connect mstrConnectString
  DoEvents

  ' Batch Jobs
  Progress "Transferring Batch Jobs..."
  strCommand = "UPDATE ASRSysBatchJobName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Report Packs
  Progress "Transferring Report Packs..."
  strCommand = "UPDATE ASRSysBatchJobName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Calendar Reports
  Progress "Transferring Calendar Reports..."
  strCommand = "UPDATE ASRSysCalendarReports SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents

  ' Cross Tabs
  Progress "Transferring CrossTabs..."
  strCommand = "UPDATE ASRSysCrossTab SET Username = '" & strTo & "'"
  sWhereExtra = sWhere & " AND CrossTabType <> " & ctt9GridBox
  
  strCommand = strCommand & sWhereExtra
  gADOCon.Execute strCommand
  DoEvents
  
  ' Talent Reports
  Progress "Transferring Talent Reports..."
  strCommand = "UPDATE ASRSysTalentReports SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' 9-Box Grid Reports
  If IsModuleEnabled(modNineBoxGrid) Then
    Progress "Transferring 9-Box Grid Reports..."
    strCommand = "UPDATE ASRSysCrossTab SET Username = '" & strTo & "'"
    sWhereExtra = sWhere & " AND CrossTabType = " & ctt9GridBox
    strCommand = strCommand & sWhereExtra
    gADOCon.Execute strCommand
    DoEvents
  End If
   
  ' Custom Reports
  Progress "Transferring Custom Reports..."
  strCommand = "UPDATE ASRSysCustomReportsName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Data Transfer
  Progress "Transferring Data Transfers..."
  strCommand = "UPDATE ASRSysDataTransferName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Diary Events
  Progress "Transferring Diary Events..."
  strCommand = "UPDATE ASRSysDiaryEvents SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Email Groups
  Progress "Transferring Email Groups..."
  strCommand = "UPDATE ASRSysEmailGroupName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents

  ' Exports
  Progress "Transferring Exports..."
  strCommand = "UPDATE ASRSysExportName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Expressions
  Progress "Transferring Expressions..."
  strCommand = "UPDATE ASRSysExpressions SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Global Functions
  Progress "Transferring Global Functions..."
  strCommand = "UPDATE ASRSysGlobalFunctions SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Import
  Progress "Transferring Imports..."
  strCommand = "UPDATE ASRSysImportName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Mail Merge
  Progress "Transferring Mail Merges..."
  strCommand = "UPDATE ASRSysMailMergeName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Match Reports
  Progress "Transferring Match Reports..."
  strCommand = "UPDATE ASRSysMatchReportName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents

  ' Talent Reports
  Progress "Transferring Organisation Reports..."
  strCommand = "UPDATE ASRSysOrganisationReport SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents

  ' Picklists
  Progress "Transferring Picklists..."
  strCommand = "UPDATE ASRSysPicklistName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Record Profiles
  Progress "Transferring Record Profiles..."
  strCommand = "UPDATE ASRSysRecordProfileName SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Envelope & Label Templates
  Progress "Transferring Envelope & Label Templates..."
  strCommand = "UPDATE ASRSysLabelTypes SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  ' Document management types
  Progress "Transferring Document Management Types..."
  strCommand = "UPDATE ASRSysDocumentManagementTypes SET Username = '" & strTo & "'"
  strCommand = strCommand & sWhere
  gADOCon.Execute strCommand
  DoEvents
  
  
  ' Close progress bar
  gobjProgress.CloseProgress

  ' Inform user
  MsgBox "Ownership transferred successfully.", vbInformation + vbOKOnly, App.Title
  
  DoTransfer = True
  
TidyUpAndExit:
  Exit Function

ErrorTrap:
  gobjProgress.Visible = False
  gobjProgress.CloseProgress
  MsgBox "Error whilst performing ownership transfer." & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  DoTransfer = False
  Resume TidyUpAndExit

End Function

Private Function PopulateToCombo() As Boolean

  Dim objGroup As SecurityGroup
  Dim objUser As SecurityUser
  Dim sMessage As String
  
  On Error GoTo ErrorTrap
  
  With cboTo
    ' Load the Users combo
    .Clear
  
    For Each objGroup In gObjGroups
      ' If the collections dont already exist, initialise them
      If Not gObjGroups(objGroup.Name).Users_Initialised Then
        InitialiseUsersCollection gObjGroups(objGroup.Name)
      End If
      
      'TM20030122 Fault 4954 - Don't add user groups to the list if they have been deleted.
      If Not objGroup.DeleteGroup Then
        gobjProgress.Bar1Caption = "Processing group '" & objGroup.Name & "'"
        
        ' Now add the users
        For Each objUser In gObjGroups(objGroup.Name).Users
          'Add users that have not been deleted or selected
          ' NPG20090205 Fault 11931
          ' If Not objUser.DeleteUser And Not Selected(objUser.UserName) And Not objUser.LoginType = iUSERTYPE_TRUSTEDGROUP Then
          If Not objUser.DeleteUser And Not Selected(objUser.UserName) _
            And Not objUser.LoginType = iUSERTYPE_TRUSTEDGROUP _
            And Not objUser.LoginType = iUSERTYPE_ORPHANUSER _
            And Not objUser.LoginType = iUSERTYPE_ORPHANGROUP Then
            .AddItem objUser.UserName
          End If
        Next objUser
        
        If gobjProgress.Visible Then
          gobjProgress.UpdateProgress (False)
        End If
      End If
      
    Next objGroup
    
    Select Case .ListCount
      Case 1
        .Enabled = False
        'TM20020215 Fault 3515
        .BackColor = vbButtonFace
      Case Is > 1
        .Enabled = True
        .BackColor = vbWindowBackground
      Case Is < 1
        If Not CanDeleteAllUsers Then
          sMessage = "You cannot delete all the users in the system! " & vbCrLf & vbCrLf & _
                "One or more of the selected users are owners of " & _
                "reports, tools or utilities."
                
          MsgBox sMessage, vbExclamation + vbOKOnly, App.Title
          PopulateToCombo = False
          GoTo TidyUpAndExit
        End If
    End Select
    
  End With

  PopulateToCombo = True

TidyUpAndExit:
  Set objGroup = Nothing
  Set objUser = Nothing
  Exit Function

ErrorTrap:
  MsgBox "Error whilst populating the available user list." & vbCrLf & vbCrLf & Err.Description, vbExclamation + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Function

Private Function CanDeleteAllUsers() As Boolean

  Dim rsTemp As New ADODB.Recordset

  On Error GoTo ErrorTrap
  
  'Connect mstrConnectString
  rsTemp.Open BuildReportSQL("<All>"), gADOCon, adOpenStatic, adLockReadOnly

  If rsTemp.RecordCount > 0 Then
    CanDeleteAllUsers = False
  Else
    CanDeleteAllUsers = True
  End If
   
  rsTemp.Close
  
TidyUpAndExit:
  Set rsTemp = Nothing
  Exit Function
  
ErrorTrap:
  MsgBox "Error checking if all users can be deleted."
  Resume TidyUpAndExit
  
End Function

Private Function NothingToTransfer() As Boolean

  Dim i As Integer
  Dim bTemp As Boolean
  
  bTemp = True
  
  For i = 1 To txtUtil.Count Step 1
    If txtUtil(i - 1).Text <> "<None>" Then bTemp = False
  Next i

'**************************************************************
' Grid Control Code.                                          *
'**************************************************************
'  For i = 1 To grdUtil.Count Step 1
'    If grdUtil(i - 1).Rows > 0 Then bAllTransfered = False
'  Next i
'**************************************************************
  
  NothingToTransfer = bTemp

TidyUpAndExit:
  Exit Function
  
ErrorTrap:
  MsgBox "Error checking if all required ownership has been transferred."
  Resume TidyUpAndExit
  
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

  If gbUseWindowsAuthentication Then
    mstrConnectString = "Driver={SQL Server};Server=" & gsSQLServerName & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";Database=" & gsDatabaseName & ";Integrated Security=SSPI;"
  Else
    mstrConnectString = "Driver={SQL Server};Server=" & gsSQLServerName & ";UID=" & gsUserName & ";PWD=" & gsPassword & ";Database=" & gsDatabaseName & ";"
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim iAnswer As Integer
  
  If (Me.cmdCancel.Tag <> "F") And (Not Me.Cancelled) Then
    iAnswer = MsgBox("To complete the deletion of the selected users, " & _
            "ownership of all the utilities must be transferred " & _
            "to remaining users." & vbCrLf & vbCrLf & "Do you wish to " & _
            "cancel without deleting the selected users?", vbExclamation + vbYesNo, App.Title)
    If (iAnswer = vbYes) Then
      Me.Cancelled = True
    Else
      Cancel = 1
      Me.Cancelled = False
    End If
  ElseIf (Me.cmdCancel.Tag = "F") And (UnloadMode <> vbFormCode) Then
    iAnswer = MsgBox("The ownership of utilities for all the selected users " & _
                "has been transferred." & vbCrLf & vbCrLf & _
                "Do you wish to continue deleting the selected users?", vbExclamation + vbYesNo, App.Title)
    If (iAnswer = vbYes) Then
      Me.Cancelled = False
    Else
      Me.Cancelled = True
    End If
  End If

End Sub


Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


