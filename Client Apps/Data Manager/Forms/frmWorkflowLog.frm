VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmWorkflowLog 
   Caption         =   "Workflow Log"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1139
   Icon            =   "frmWorkflowLog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFilters 
      Caption         =   "Filters :"
      Height          =   945
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   11910
      Begin VB.ComboBox cboTargetName 
         Height          =   315
         Left            =   6315
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   500
         Width           =   3300
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmWorkflowLog.frx":000C
         Left            =   3270
         List            =   "frmWorkflowLog.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   500
         Width           =   2625
      End
      Begin VB.ComboBox cboUser 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   500
         Width           =   2000
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmWorkflowLog.frx":0010
         Left            =   10200
         List            =   "frmWorkflowLog.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   500
         Width           =   1545
      End
      Begin VB.Label lblTargetName 
         Caption         =   "Identified Record :"
         Height          =   210
         Left            =   6315
         TabIndex        =   15
         Top             =   255
         Width           =   1680
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Workflow Name :"
         Height          =   195
         Left            =   3270
         TabIndex        =   3
         Top             =   250
         Width           =   1485
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initiator :"
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   250
         Width           =   810
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         Height          =   195
         Left            =   10200
         TabIndex        =   5
         Top             =   255
         Width           =   705
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2650
      Left            =   12180
      TabIndex        =   8
      Top             =   200
      Width           =   1200
      Begin VB.CommandButton cmdRebuild 
         Caption         =   "Re&build..."
         Height          =   400
         Left            =   0
         TabIndex        =   12
         Top             =   1650
         Width           =   1200
      End
      Begin VB.CommandButton cmdPurge 
         Caption         =   "&Purge..."
         Height          =   400
         Left            =   0
         TabIndex        =   13
         Top             =   2200
         Width           =   1200
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete..."
         Height          =   400
         Left            =   0
         TabIndex        =   11
         Top             =   1100
         Width           =   1200
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View..."
         Height          =   400
         Left            =   0
         TabIndex        =   10
         Top             =   525
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1200
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   14
      Top             =   4845
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   23230
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B.SSDBGrid grdWorkflowLog 
      Height          =   3525
      Left            =   90
      TabIndex        =   7
      Top             =   1185
      Width           =   11910
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      stylesets.count =   2
      stylesets(0).Name=   "ssetDormant"
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
      stylesets(0).Picture=   "frmWorkflowLog.frx":0014
      stylesets(1).Name=   "ssetActive"
      stylesets(1).ForeColor=   16777215
      stylesets(1).BackColor=   -2147483646
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
      stylesets(1).Picture=   "frmWorkflowLog.frx":0030
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
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   0
      StyleSet        =   "ssetDormant"
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ActiveRowStyleSet=   "ssetActive"
      CaptionAlignment=   0
      Columns.Count   =   8
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "ID"
      Columns(0).Name =   "ID"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2990
      Columns(1).Caption=   "Initiation Time"
      Columns(1).Name =   "StartTime"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2990
      Columns(2).Caption=   "Completion Time"
      Columns(2).Name =   "EndTime"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1667
      Columns(3).Caption=   "Duration"
      Columns(3).Name =   "Duration"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3519
      Columns(4).Caption=   "Workflow Name"
      Columns(4).Name =   "Utility Name"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2302
      Columns(5).Caption=   "Status"
      Columns(5).Name =   "Status"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3254
      Columns(6).Caption=   "Initiator"
      Columns(6).Name =   "Username"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   5292
      Columns(7).Caption=   "Identified Record"
      Columns(7).Name =   "TargetName"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      TabNavigation   =   1
      _ExtentX        =   21008
      _ExtentY        =   6218
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
   Begin ActiveBarLibraryCtl.ActiveBar abWorkflowLog 
      Left            =   12810
      Top             =   4185
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
      Bands           =   "frmWorkflowLog.frx":004C
   End
End
Attribute VB_Name = "frmWorkflowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Flag to prevent grid refreshing when combos are being populated and set initially
Private mblnLoading As Boolean

' Flag showing if user is viewing all or just his own entries
Private mblnViewAllEntries As Boolean

' Must be public so the details form can change the bookmark of the recordset
Public mrstHeaders As Recordset

' Data access class
Private mclsData As New clsDataAccess

' Variables to hold the column clicked on, its field and the order to sort the grid
Private pstrOrderField As String
Private pstrOrderOrder As String
Private mintSortColumnIndex As Integer

' Flag showing if user has delete permission
Private mblnDeleteEnabled As Boolean



Private Sub abWorkflowLog_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.abWorkflowLog_Click(Tool)", Array(Tool)

  Select Case Tool.Name
    Case "View"
      ViewWorkflow
      
    Case "Delete"
      DeleteWorkflow
  End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub DeleteWorkflow()
  
  Dim strWorkflowIDs As String
  Dim strQueueIDs As String
  Dim plngLoop As Long
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim arrayBookmarks() As Variant
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.DeleteWorkflow()"
  
  With frmSelection
    .Source = "Workflow Log"
    .HelpContextID = Me.HelpContextID
    .Caption = Me.Caption & " Selection"
    .Show vbModal
  End With

  Select Case frmSelection.Answer
  
    Case -1 ' Cancelled
    
      Unload frmSelection
      GoTo TidyUpAndExit
      
    Case 0 ' Delete the currently highlighted entries
    
      Screen.MousePointer = vbHourglass

      'Workout how many records have been selected
      nTotalSelRows = grdWorkflowLog.SelBookmarks.Count
        
      'Redimension the arrays to the count of the bookmarks
      ReDim arrayBookmarks(nTotalSelRows)
      
      For intCount = 1 To nTotalSelRows
        arrayBookmarks(intCount) = grdWorkflowLog.SelBookmarks.item(intCount - 1)
      Next intCount
      
      For intCount = 1 To nTotalSelRows
        grdWorkflowLog.Bookmark = arrayBookmarks(intCount)
        
        If grdWorkflowLog.Columns("Status").value = WorkflowStatusDescription(giWFSTATUS_SCHEDULED) _
          And grdWorkflowLog.Columns("Username").value = "<Triggered>" Then
                
          strQueueIDs = strQueueIDs & _
            IIf(Len(strQueueIDs) > 0, ",", "") & _
            grdWorkflowLog.Columns("ID").value
        Else
          strWorkflowIDs = strWorkflowIDs & _
            IIf(Len(strWorkflowIDs) > 0, ",", "") & _
            grdWorkflowLog.Columns("ID").value
        End If
      Next intCount
      grdWorkflowLog.SelBookmarks.RemoveAll

      If Len(strWorkflowIDs) > 0 Then
        gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances WHERE ID IN (" & strWorkflowIDs & ")"
      End If
      If Len(strQueueIDs) > 0 Then
        gADOCon.Execute "DELETE FROM ASRSysWorkflowQueue WHERE queueID IN (" & strQueueIDs & ")"
      End If
      
    Case 1 ' Delete entries currently displayed in the grid
    
      Screen.MousePointer = vbHourglass
      
      mrstHeaders.MoveFirst
      Do Until mrstHeaders.EOF
        
        If mrstHeaders.Fields("Status") = WorkflowStatusDescription(giWFSTATUS_SCHEDULED) _
          And mrstHeaders.Fields("Username") = "<Triggered>" Then
        
          strQueueIDs = strQueueIDs & _
            IIf(Len(strQueueIDs) > 0, ",", "") & _
            mrstHeaders.Fields("ID")
        Else
          strWorkflowIDs = strWorkflowIDs & _
            IIf(Len(strWorkflowIDs) > 0, ",", "") & _
            mrstHeaders.Fields("ID")
        End If
        
        mrstHeaders.MoveNext
      Loop

      If Len(strWorkflowIDs) > 0 Then
        gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances WHERE ID IN (" & strWorkflowIDs & ")"
      End If
      If Len(strQueueIDs) > 0 Then
        gADOCon.Execute "DELETE FROM ASRSysWorkflowQueue WHERE queueID IN (" & strQueueIDs & ")"
      End If
      

    Case 2 ' Delete all entries visible to the user
    
      Screen.MousePointer = vbHourglass
    
      If mblnViewAllEntries = True Then
        gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances"
        gADOCon.Execute "DELETE FROM ASRSysWorkflowQueue"
      Else
        gADOCon.Execute "DELETE FROM ASRSysWorkflowInstances WHERE username = '" & datGeneral.UserNameForSQL & "'"
      End If
      
  End Select
  
'  If Me.grdWorkflowLog.Rows > 0 Then
'    Me.grdWorkflowLog.SelBookmarks.RemoveAll
'    Me.grdWorkflowLog.MoveFirst
'    Me.grdWorkflowLog.SelBookmarks.Add Me.grdWorkflowLog.Bookmark
'  End If

  Unload frmSelection
  Screen.MousePointer = vbDefault
  RefreshGrid
  RefreshButtons

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
  
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub RefreshButtons()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.RefreshButtons()"

  cmdView.Enabled = Me.grdWorkflowLog.Rows > 0
  cmdDelete.Enabled = (Me.grdWorkflowLog.Rows > 0) And (mblnDeleteEnabled)
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub



Private Function ViewWorkflow()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.ViewWorkflow()"

  Dim frmDetails As frmWorkflowLogDetails
  Dim frmQueueDetails As frmWorkflowQueueDetails
  Dim varBookmark As Variant
  
  If IsNumeric(grdWorkflowLog.Columns("ID").value) Then
    If grdWorkflowLog.Columns("Status").value = WorkflowStatusDescription(giWFSTATUS_SCHEDULED) _
      And grdWorkflowLog.Columns("Username").value = "<Triggered>" Then
    
      Set frmQueueDetails = New frmWorkflowQueueDetails
      
      If frmQueueDetails.Initialise(CLng(grdWorkflowLog.Columns("ID").value)) Then
        frmQueueDetails.Show vbModal

        If grdWorkflowLog.SelBookmarks.Count = 1 Then
          varBookmark = grdWorkflowLog.SelBookmarks.item(0)
        End If

        RefreshGrid

        If (Not IsEmpty(varBookmark)) _
          And (grdWorkflowLog.Rows > 0) Then
          
          grdWorkflowLog.SelBookmarks.RemoveAll
          grdWorkflowLog.SelBookmarks.Add varBookmark
          grdWorkflowLog.Bookmark = varBookmark
        End If
      Else
        RefreshGrid
        Unload frmQueueDetails
      End If
    Else
      Set frmDetails = New frmWorkflowLogDetails
      If frmDetails.Initialise(CLng(grdWorkflowLog.Columns("ID").value), Me) Then
  
        frmDetails.Caption = Me.Caption & " Details"
        frmDetails.Show vbModal
      
        If grdWorkflowLog.SelBookmarks.Count = 1 Then
          varBookmark = grdWorkflowLog.SelBookmarks.item(0)
        End If
        
        RefreshGrid
      
        If (Not IsEmpty(varBookmark)) _
          And (grdWorkflowLog.Rows > 0) Then
          
          grdWorkflowLog.SelBookmarks.RemoveAll
          grdWorkflowLog.SelBookmarks.Add varBookmark
          grdWorkflowLog.Bookmark = varBookmark
        End If
      Else
        RefreshGrid
        Unload frmDetails
      End If
    End If
  End If

TidyUpAndExit:
  Set frmDetails = Nothing
  gobjErrorStack.PopStack
  Exit Function
  
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function


Private Sub cboStatus_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cboStatus_Click()"

  RefreshGrid
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cboTargetName_Click()
  RefreshGrid
End Sub

Private Sub cboType_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cboStatus_Click()"

  RefreshGrid
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cboUser_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cboUser_Click()"

  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cmdDelete_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cmdDelete_Click()"

  DeleteWorkflow
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cmdOK_Click()"

  Unload Me

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cmdPurge_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cmdPurge_Click()"

  frmWorkflowLogPurge.Show vbModal
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub cmdRebuild_Click()

  Dim strMBText As String
  Dim intMBButtons As Long
  Dim strMBCaption As String
  Dim intMBResponse As Integer
  Dim sSQL As String
  Dim fTriggersExist As Boolean
  Dim rsCount As ADODB.Recordset
  
  On Error GoTo LocalErr

  'JPD 20061116 Fault 11695 Only perform the 'rebuild' option if there are triggered Workflows to rebuild.
  sSQL = "SELECT COUNT(*) AS result" & _
    " FROM ASRSysWorkflowTriggeredLinks" & _
    " WHERE type = 2" ' 2 = date
  Set rsCount = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  fTriggersExist = (rsCount!Result > 0)
  rsCount.Close
  Set rsCount = Nothing

  If Not fTriggersExist Then
    COAMsgBox "No 'Date' type triggered workflow links exist. Triggered workflow queue will not be rebuilt.", vbInformation + vbOKOnly, Me.Caption
  Else
    strMBText = "Are you sure that you would like to rebuild the triggered workflow queue?"
    intMBButtons = vbQuestion + vbYesNo
    strMBCaption = Me.Caption
    intMBResponse = COAMsgBox(strMBText, intMBButtons, strMBCaption)
  
    If intMBResponse <> vbYes Then
      Exit Sub
    End If
  
    Screen.MousePointer = vbHourglass
    With gobjProgress
      '.AviFile = App.Path & "\videos\diary.avi"
      .AVI = dbWorkflow
      .MainCaption = "Workflow"
      .NumberOfBars = 0
      .Caption = "Workflow Queue Rebuild"
      .Time = False
      .Cancel = False
      .OpenProgress
      .Bar1Caption = "Workflow Queue Rebuild"
    End With
  
    ' Event Log Header
    gobjEventLog.AddHeader eltWorkflowRebuild, "Workflow Rebuild"
  
    datGeneral.ExecuteSql "EXEC spASRWorkflowRebuild", ""
    Call RefreshGrid
  
    gobjProgress.CloseProgress
    Screen.MousePointer = vbDefault
  
    ' Event Log Header
    gobjEventLog.ChangeHeaderStatus elsSuccessful
  End If
  
Exit Sub

LocalErr:
  COAMsgBox "Error rebuilding Workflow queue", vbCritical, Me.Caption
  ' Event Log Header
  gobjEventLog.ChangeHeaderStatus elsFailed

End Sub

Private Sub cmdView_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.cmdView_Click()"

  ViewWorkflow

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_Activate()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.Form_Activate()"

  DoColumnSizes
  
  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub DoColumnSizes()
  Dim lngAvailableWidth As Long
    
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.DoColumnSizes()"

  With grdWorkflowLog
    lngAvailableWidth = grdWorkflowLog.Width - (270 + .Columns(1).Width + .Columns(2).Width + .Columns(3).Width + .Columns(5).Width)

    .Columns(4).Width = (lngAvailableWidth * 0.399)
    .Columns(6).Width = (lngAvailableWidth * 0.2)
    .Columns(7).Width = (lngAvailableWidth * 0.399)
    
  End With
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.Form_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  Case KeyCode = vbKeyEscape
    Unload Me
  Case KeyCode = vbKeyF5
    RefreshGrid
  Case KeyCode = vbKeyDelete
    If cmdDelete.Enabled Then
      cmdDelete_Click
    End If
End Select

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub

Private Function RefreshGrid() As Boolean
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.RefreshGrid()"
  
  ' Populate the grid using filter/sort criteria as set by the user
  Dim pstrSQL As String
  Dim strStatusBarText As String
  Dim sWhereClause As String
  
  If mblnLoading = True Then GoTo TidyUpAndExit
  
  Screen.MousePointer = vbHourglass
  
  pstrSQL = "SELECT * FROM (SELECT WI.ID AS [ID]," & _
    "   WI.InitiationDateTime AS [InitiationDateTime]," & _
    "   WI.CompletionDateTime AS [CompletionDateTime]," & _
    "   CASE" & _
    "     WHEN WI.InitiationDateTime IS null OR WI.CompletionDateTime IS null THEN 0" & _
    "     ELSE datediff(s, WI.InitiationDateTime, WI.CompletionDateTime)" & _
    "   END AS [Duration]," & _
    "   WF.name AS [Name]," & _
    "   CASE WI.Status" & _
    "     WHEN " & CStr(giWFSTATUS_INPROGRESS) & " THEN '" & WorkflowStatusDescription(giWFSTATUS_INPROGRESS) & "'" & _
    "     WHEN " & CStr(giWFSTATUS_CANCELLED) & " THEN '" & WorkflowStatusDescription(giWFSTATUS_CANCELLED) & "'" & _
    "     WHEN " & CStr(giWFSTATUS_ERROR) & " THEN '" & WorkflowStatusDescription(giWFSTATUS_ERROR) & "'" & _
    "     WHEN " & CStr(giWFSTATUS_COMPLETED) & " THEN '" & WorkflowStatusDescription(giWFSTATUS_COMPLETED) & "'" & _
    "   END AS [Status]," & _
    "   CASE WF.initiationType" & _
    "     WHEN 2 THEN '<External>'" & _
    "     ELSE WI.Username" & _
    "   END AS [Username]," & _
    "   ISNULL([TargetName], '') AS [TargetName]" & _
    " FROM ASRSysWorkflowInstances WI" & _
    " INNER JOIN ASRSysWorkflows WF ON WI.workflowID = WF.ID"

  sWhereClause = vbNullString

  If cboType.ListIndex > 0 Then
    sWhereClause = sWhereClause & " WHERE [Name] = '" & Replace(cboType.Text, "'", "''") & "'"
  End If

  If cboStatus.ListIndex > 0 Then
    sWhereClause = sWhereClause & IIf(InStr(sWhereClause, "WHERE") > 0, " AND ", " WHERE ") & _
      "status = '" & cboStatus.Text & "'"
  End If

  If cboTargetName.ListIndex > 0 Then
    sWhereClause = sWhereClause & IIf(InStr(sWhereClause, "WHERE") > 0, " AND ", " WHERE ") & _
      "TargetName = '" & Replace(cboTargetName.Text, "'", "''") & "'"
  End If

  If mblnViewAllEntries = False Then
    sWhereClause = sWhereClause & IIf(InStr(sWhereClause, "WHERE") > 0, " AND ", " WHERE ") & _
      "Username = '" & Replace(gsUserName, "'", "''") & "'"
  Else
    If cboUser.Text <> "<All>" Then
      sWhereClause = sWhereClause & IIf(InStr(sWhereClause, "WHERE") > 0, " AND ", " WHERE ") & _
        "Username = '" & Replace(cboUser.Text, "'", "''") & "'"
    End If
  End If

  If mblnViewAllEntries And _
    ((cboStatus.ItemData(cboStatus.ListIndex) = giWFSTATUS_ALL) Or (cboStatus.ItemData(cboStatus.ListIndex) = giWFSTATUS_SCHEDULED)) And _
    ((cboUser.Text = "<All>") Or (cboUser.Text = "<Triggered>")) Then
    ' All or Scheduled to be shown.
    
    pstrSQL = pstrSQL & _
      " UNION " & _
      " SELECT Q.queueID AS [ID]," & _
      " Q.dateDue AS [InitiationDateTime]," & _
      " null AS [CompletionDateTime]," & _
      " 0 AS [Duration]," & _
      " ISNULL(WF.Name,'<Deleted>') AS [Name]," & _
      " '" & WorkflowStatusDescription(giWFSTATUS_SCHEDULED) & "' AS [Status]," & _
      " '<Triggered>' AS [Username]," & _
      " '' AS [TargetName]" & _
      " FROM [ASRSysWorkflowQueue] Q" & _
      " INNER JOIN [ASRSysWorkflowTriggeredLinks] TL ON Q.linkID = TL.linkID" & _
      " INNER JOIN [ASRSysWorkflows] WF ON TL.workflowID = WF.ID" & _
      "     AND WF.enabled = 1" & _
      " WHERE Q.dateInitiated IS null"
  
    If cboType.ListIndex > 0 Then
      pstrSQL = pstrSQL & " AND [Name] = '" & Replace(cboType.Text, "'", "''") & "'"
    End If
  End If

  pstrSQL = pstrSQL & ") tableAllRecords " & _
    sWhereClause & _
    " ORDER BY [" & pstrOrderField & "] " & pstrOrderOrder

  Set mrstHeaders = mclsData.OpenPersistentRecordset(pstrSQL, adOpenKeyset, adLockReadOnly)

  With grdWorkflowLog
    .Redraw = False
    .Rebind
    .Rows = mrstHeaders.RecordCount
    .Redraw = True

    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With

  strStatusBarText = vbNullString
  strStatusBarText = strStatusBarText & " " & mrstHeaders.RecordCount & " record" & IIf(mrstHeaders.RecordCount > 1 Or mrstHeaders.RecordCount = 0, "s", "")
  If mrstHeaders.RecordCount > 1 Then
    strStatusBarText = strStatusBarText & " sorted by "
    If pstrOrderField = "InitiationDateTime" Then
      strStatusBarText = strStatusBarText & "Initiation Time "
    ElseIf pstrOrderField = "CompletionDateTime" Then
      strStatusBarText = strStatusBarText & "Completion Time "
    ElseIf pstrOrderField = "Username" Then
      strStatusBarText = strStatusBarText & "Initiator "
    Else
      strStatusBarText = strStatusBarText & pstrOrderField & " "
    End If
    
    strStatusBarText = strStatusBarText & "in "
    strStatusBarText = strStatusBarText & IIf(pstrOrderOrder = "ASC", "ascending", "descending")
    strStatusBarText = strStatusBarText & " order"
  End If

  StatusBar1.SimpleText = strStatusBarText

  RefreshButtons
  
  DoColumnSizes

  Screen.MousePointer = vbDefault

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Function
ErrorTrap:
  gobjErrorStack.HandleError
  
End Function


Private Sub Form_Load()

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.Form_Load()"

  Hook Me.hWnd, 11700, 5550

  Dim rstUsers As Recordset
  Dim rstWorkflows As Recordset
  Set mclsData = New clsDataAccess

  mblnLoading = True

  fraButtons.BackColor = Me.BackColor

  ' Get rid of the icon off the form
  RemoveIcon Me

  'If user does not have Workflow Administer permission, hide the delete, rebuild and purge buttons
  If datGeneral.SystemPermission("WORKFLOW", "ADMINISTER") = False Then
    cmdDelete.Enabled = False
    mblnDeleteEnabled = False
  
    cmdRebuild.Enabled = False
    cmdPurge.Enabled = False
  Else
    mblnDeleteEnabled = True
  End If

  'If user can see all entries, populate and enable the users combo, else
  'populate it with the users name only and disable it
  If datGeneral.SystemPermission("WORKFLOW", "ADMINISTER") = True Then
    mblnViewAllEntries = True
    cboUser.AddItem "<All>"
    cboUser.AddItem "<External>"
    cboUser.AddItem "<Triggered>"
    
    Set rstUsers = mclsData.OpenRecordset("SELECT DISTINCT Username from ASRSysWorkflowInstances WHERE (NOT userName IS null) AND (userName <> '<Triggered>') AND (userName <> '<External>') ORDER BY Username", adOpenForwardOnly, adLockReadOnly)
    Do Until rstUsers.EOF
      cboUser.AddItem rstUsers.Fields("Username")
      rstUsers.MoveNext
    Loop
    Set rstUsers = Nothing
  Else
    Me.Caption = Me.Caption & " [Viewing own entries only]"
    cboUser.AddItem gsUserName
    cboUser.Enabled = False
    cboUser.BackColor = &H8000000F
  End If

  cboUser.ListIndex = 0

  'Purge the event log before we display it
  gADOCon.Execute "EXEC spASRWorkflowLogPurge"

  'Add all available workflow names to the Type combo
  With cboType
    .AddItem "<All>"
    
    Set rstWorkflows = mclsData.OpenRecordset("SELECT DISTINCT name from ASRSysWorkflows ORDER BY name", adOpenForwardOnly, adLockReadOnly)
    Do Until rstWorkflows.EOF
      .AddItem rstWorkflows.Fields("name")
      rstWorkflows.MoveNext
    Loop
    Set rstWorkflows = Nothing
    
    .ListIndex = 0
  End With

  ' Add all distinct target names
  With cboTargetName
    .AddItem "<All>"
       
    Set rstWorkflows = mclsData.OpenRecordset("SELECT DISTINCT TargetName FROM ASRSysWorkflowInstances ORDER BY TargetName", adOpenForwardOnly, adLockReadOnly)
    Do Until rstWorkflows.EOF
      .AddItem IIf(IsNull(rstWorkflows.Fields("TargetName")), "", rstWorkflows.Fields("TargetName"))
      rstWorkflows.MoveNext
    Loop
    Set rstWorkflows = Nothing
    
    .ListIndex = 0
  End With



  'Add all available statuses to the Status combo
  With cboStatus
    .AddItem WorkflowStatusDescription(giWFSTATUS_ALL)
    .ItemData(.NewIndex) = giWFSTATUS_ALL
    
'    .AddItem WorkflowStatusDescription(giWFSTATUS_CANCELLED)
'    .ItemData(.NewIndex) = giWFSTATUS_CANCELLED
    
    .AddItem WorkflowStatusDescription(giWFSTATUS_COMPLETED)
    .ItemData(.NewIndex) = giWFSTATUS_COMPLETED
    
    .AddItem WorkflowStatusDescription(giWFSTATUS_ERROR)
    .ItemData(.NewIndex) = giWFSTATUS_ERROR
    
    .AddItem WorkflowStatusDescription(giWFSTATUS_INPROGRESS)
    .ItemData(.NewIndex) = giWFSTATUS_INPROGRESS
        
    .AddItem WorkflowStatusDescription(giWFSTATUS_SCHEDULED)
    .ItemData(.NewIndex) = giWFSTATUS_SCHEDULED
        
    .ListIndex = 0
  End With

  mblnLoading = False

  'Set height and width to last saved. Form is centred on screen
  Me.Height = GetPCSetting("WorkflowLog", "Height", Me.Height)
  Me.Width = GetPCSetting("WorkflowLog", "Width", Me.Width)

  'Set default sort order to be date desc
  pstrOrderField = "InitiationDateTime"
  mintSortColumnIndex = 1
  pstrOrderOrder = "DESC"
  
  'Populate the grid
  RefreshGrid

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)

  ' Save the window size ready to recall next time user views the workflow log
  SavePCSetting "WorkflowLog", "Height", Me.Height
  SavePCSetting "WorkflowLog", "Width", Me.Width
  UI.RemoveClipping

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub Form_Resize()

  Const COMBO_GAP As Integer = 170
  Const lngGap As Long = 120
  
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.Form_Resize()"

  'JPD 20030908 Fault 5756
  DisplayApplication
  
  ' Ensure form does not get too small/big. Also reposition controls as necessary
  fraButtons.Left = Me.ScaleWidth - (fraButtons.Width + lngGap)
  fraFilters.Width = fraButtons.Left - (lngGap * 2)
  
  cboTargetName.Left = fraFilters.Width - (cboTargetName.Width + COMBO_GAP)
  lblTargetName.Left = cboTargetName.Left

  cboStatus.Left = cboTargetName.Left - (cboStatus.Width + COMBO_GAP)
  lblStatus.Left = cboStatus.Left

  cboType.Left = cboStatus.Left - (cboType.Width + COMBO_GAP)
  lblName.Left = cboType.Left

  cboUser.Width = cboType.Left - (cboUser.Left + COMBO_GAP)

  grdWorkflowLog.Width = fraFilters.Width
  grdWorkflowLog.Height = Me.ScaleHeight - (fraFilters.Height + StatusBar1.Height + (lngGap * 3))
  
  DoColumnSizes
  
  Me.Refresh
  grdWorkflowLog.Refresh
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

Private Sub grdWorkflowLog_Click()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_Click()"

  If (Me.grdWorkflowLog.SelBookmarks.Count > 1) Or (Me.grdWorkflowLog.Rows = 0) Then
    Me.cmdView.Enabled = False
  Else
    Me.cmdView.Enabled = True
  End If
  
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub grdWorkflowLog_DblClick()
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_DblClick()"
  
  If (Me.grdWorkflowLog.Rows > 0) And Me.grdWorkflowLog.SelBookmarks.Count = 1 Then
    ViewWorkflow
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub grdWorkflowLog_HeadClick(ByVal ColIndex As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_HeadClick(ColIndex)", Array(ColIndex)

  ' Set the sort criteria depending on the column header clicked and refresh the grid
  Select Case ColIndex
    Case 1: pstrOrderField = "InitiationDateTime"
    Case 2: pstrOrderField = "CompletionDateTime"
    Case 3: pstrOrderField = "Duration"
    Case 4: pstrOrderField = "Name"
    Case 5: pstrOrderField = "Status"
    Case 6: pstrOrderField = "Username"
    Case 7: pstrOrderField = "TargetName"
  End Select
  
  If ColIndex = mintSortColumnIndex Then
    If pstrOrderOrder = "ASC" Then pstrOrderOrder = "DESC" Else pstrOrderOrder = "ASC"
  End If
  
  mintSortColumnIndex = ColIndex
  
  RefreshGrid
    
TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub grdWorkflowLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_MouseUp(Button,Shift,X,Y)", Array(Button, Shift, X, Y)

 If (Button = vbRightButton) And (Y > Me.grdWorkflowLog.RowHeight) Then
    ' Enable/disable the required tools.
    With Me.abWorkflowLog.Bands("bndWorkflowLog")
      .Tools("View").Enabled = Me.cmdView.Enabled
      .Tools("Delete").Enabled = Me.cmdDelete.Enabled
      .TrackPopup -1, -1
    End With
    
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub grdWorkflowLog_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol)
    
  If Me.grdWorkflowLog.SelBookmarks.Count > 1 Then
    Me.cmdView.Enabled = False
  ElseIf Me.grdWorkflowLog.SelBookmarks.Count = 1 Then
    Me.grdWorkflowLog.SelBookmarks.RemoveAll
    Me.grdWorkflowLog.SelBookmarks.Add Me.grdWorkflowLog.Bookmark
    Me.cmdView.Enabled = True
  End If

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub grdWorkflowLog_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_UnboundPositionData(StartLocation,NumberOfRowsToMove,NewLocation)", Array(StartLocation, NumberOfRowsToMove, NewLocation)
  
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

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub


Private Sub grdWorkflowLog_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "frmWorkflowLog.grdWorkflowLog_UnboundReadData(RowBuf,StartLocation,ReadPriorRows)", Array(RowBuf, StartLocation, ReadPriorRows)
  
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim sDateFormat As String
  
  sDateFormat = DateFormat
  
  iRowsRead = 0
  
  ' Do nothing if we a re just formatting the grid,
  ' or if there a re no records to display.
  If mrstHeaders Is Nothing Then GoTo TidyUpAndExit
  If mrstHeaders.State = adStateClosed Then GoTo TidyUpAndExit

  ' Do nothing if we are loading or if there are no records to display
  If mblnLoading = True And mrstHeaders.RecordCount = 0 Then GoTo TidyUpAndExit
  
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
            Case "InitiationDateTime", "CompletionDateTime"
              RowBuf.value(iRowIndex, iFieldIndex) = Format(mrstHeaders.Fields(iFieldIndex), sDateFormat & " hh:nn")
            Case "Duration"
              RowBuf.value(iRowIndex, iFieldIndex) = IIf(mrstHeaders.Fields("Duration").value = 0, "", FormatEventDuration(mrstHeaders.Fields("Duration").value))
            Case Else
              RowBuf.value(iRowIndex, iFieldIndex) = CStr(IIf(IsNull(mrstHeaders.Fields(iFieldIndex)), "", mrstHeaders.Fields(iFieldIndex)))
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark

      Case 1
        RowBuf.Bookmark(iRowIndex) = mrstHeaders.Bookmark

    End Select

    If ReadPriorRows Then
      mrstHeaders.MovePrevious
    Else
      mrstHeaders.MoveNext
    End If

    iRowsRead = iRowsRead + 1
  Next iRowIndex

  RowBuf.RowCount = iRowsRead

TidyUpAndExit:
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub



