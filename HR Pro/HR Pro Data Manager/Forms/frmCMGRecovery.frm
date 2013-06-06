VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCMGRecovery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CMG Recovery"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1071
   Icon            =   "frmCMGRecovery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid dbCommitDates 
      Height          =   2835
      Left            =   180
      TabIndex        =   2
      Top             =   165
      Width           =   4515
      _Version        =   196617
      DataMode        =   2
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      Col.Count       =   2
      DividerType     =   2
      BeveColorScheme =   0
      CheckBox3D      =   0   'False
      AllowUpdate     =   0   'False
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorOdd    =   16777215
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   7964
      Columns(0).Caption=   "Commit Date"
      Columns(0).Name =   "Commit Date"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).HasHeadBackColor=   -1  'True
      Columns(0).HasBackColor=   -1  'True
      Columns(0).HeadBackColor=   -2147483633
      Columns(0).BackColor=   16777215
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "ActualCommitDate"
      Columns(1).Name =   "ActualCommitDate"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   7964
      _ExtentY        =   5001
      _StockProps     =   79
      ForeColor       =   0
      BackColor       =   16777215
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
   Begin VB.CommandButton cmdRollback 
      Caption         =   "&Rollback"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5085
      TabIndex        =   1
      Top             =   2085
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5085
      TabIndex        =   0
      Top             =   2595
      Width           =   1200
   End
End
Attribute VB_Name = "frmCMGRecovery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function Initialise() As Boolean

  LoadCMGDates

End Function

Private Function LoadCMGDates() As Boolean

  ' Populates the list with all the dates when a commit has taken place

  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  Dim strDates As String

  ' Clear the list
  dbCommitDates.RemoveAll
   
  sSQL = "Select Distinct cmgcommitdate AS CMGCommitDate " _
    & "From ASRSysAuditTrail " _
    & "Where CMGCommitDate Is Not Null Order By 1 Desc"
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If Not rsTemp.EOF Then
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
      strDates = Format(rsTemp!CMGCommitDate, "dddd, DD Mmmm YYYY") & "  (" & Format(rsTemp!CMGCommitDate, "Long Time") & ")" _
        & vbTab & rsTemp!CMGCommitDate
      dbCommitDates.AddItem strDates
      rsTemp.MoveNext
    Loop
  Else
    cmdRollback.Enabled = False
  End If

  rsTemp.Close
  Set rsTemp = Nothing

End Function

Private Sub cmdOK_Click()

  Unload Me

End Sub

Private Sub cmdRollback_Click()

  ' Rollback the changes to the cmg audit log
  Dim sSQL As String
  Dim bOK As Boolean
  Dim strErrorMessage As String
  Dim strRollbackDate As String

  On Error GoTo ErrTrap
  bOK = True
  
  If COAMsgBox("Are you sure you want to rollback. This process cannot be reversed.", vbQuestion + vbYesNo, "Warning") = vbYes Then
  
    strRollbackDate = Format(dbCommitDates.Columns("ActualCommitDate").Value, "YYYY/MM/DD hh:mm:ss")
  
    ' Update the audit log to say these have been exported
    sSQL = "Update asrsysAuditTrail Set CMGCommitDate = NULL " _
      & "Where CMGCommitDate >= '" & strRollbackDate & "'"
    datGeneral.ExecuteSql sSQL, strErrorMessage
    
    ' Refresh the dates list
    LoadCMGDates
  End If
  
ExitPoint:
  Exit Sub
  
ErrTrap:
  bOK = False
  Resume ExitPoint

End Sub

Private Sub dbCommitDates_Click()

  If dbCommitDates.Rows > 0 And Not dbCommitDates.SelBookmarks Is Nothing Then
    cmdRollback.Enabled = True
  End If

End Sub

Private Sub Form_Initialize()

  Initialise

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub



