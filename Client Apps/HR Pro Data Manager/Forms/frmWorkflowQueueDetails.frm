VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{1EE59219-BC23-4BDF-BB08-D545C8A38D6D}#1.1#0"; "COA_Line.ocx"
Begin VB.Form frmWorkflowQueueDetails 
   Caption         =   "Workflow Queue Details"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1150
   Icon            =   "frmWorkflowQueueDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDetails 
      Caption         =   "Details :"
      Height          =   4700
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7750
      Begin SSDataWidgets_B.SSDBGrid grdColumnValues 
         Height          =   2325
         Left            =   200
         TabIndex        =   2
         Top             =   2150
         Width           =   7350
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         Col.Count       =   2
         AllowUpdate     =   0   'False
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
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         CellNavigation  =   1
         MaxSelectedRows =   0
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   6165
         Columns(0).Caption=   "Column"
         Columns(0).Name =   "Column"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   4710
         Columns(1).Caption=   "Value"
         Columns(1).Name =   "Value"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         UseDefaults     =   0   'False
         TabNavigation   =   1
         _ExtentX        =   12965
         _ExtentY        =   4101
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin COALine.COA_Line linColumns 
         Height          =   30
         Left            =   200
         Top             =   1900
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   53
      End
      Begin VB.Label lblRecordDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Record Description :"
         Height          =   195
         Left            =   195
         TabIndex        =   14
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Label lblRecordDescriptionValue 
         BackStyle       =   0  'Transparent
         Caption         =   "RecordDescriptionValue"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1980
         TabIndex        =   13
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblLinkType 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Type :"
         Height          =   195
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label lblLinkTypeValue 
         BackStyle       =   0  'Transparent
         Caption         =   "LinkTypeValue"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5235
         TabIndex        =   11
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label lblTable 
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Left            =   4200
         TabIndex        =   10
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblTableValue 
         BackStyle       =   0  'Transparent
         Caption         =   "TableValue"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   5235
         TabIndex        =   9
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label lblInitiator 
         BackStyle       =   0  'Transparent
         Caption         =   "Initiator :"
         Height          =   195
         Left            =   195
         TabIndex        =   8
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblInitiatorValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Initiator"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1980
         TabIndex        =   7
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblInitiationTimeValue 
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999  00:00"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1980
         TabIndex        =   6
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label lblInitiationTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Initiation Time :"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label lblWorkflowName 
         BackStyle       =   0  'Transparent
         Caption         =   "Workflow Name :"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lblWorkflowNameValue 
         BackStyle       =   0  'Transparent
         Caption         =   "WorkflowName"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1980
         TabIndex        =   3
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   6670
      TabIndex        =   0
      Top             =   4950
      Width           =   1200
   End
End
Attribute VB_Name = "frmWorkflowQueueDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsData As New clsDataAccess

Private mblnSizing As Boolean
Private mfSizeable As Boolean

Private miLinkType As Integer

Private Const MINFORMWIDTH = 8300
Private Const MINFORMHEIGHT = 5000
  
Private Sub cmdOK_Click()
  Unload Me

End Sub

Private Sub Form_Load()
  Hook Me.hWnd, MINFORMWIDTH, MINFORMHEIGHT

  RemoveIcon Me
  
  Set mclsData = New clsDataAccess
  grdColumnValues.RowHeight = 239

  ' Retrieve the size of the form when last viewed
  Me.Height = GetPCSetting("WorkflowLogQueueDetails", "Height", Me.Height)
  Me.Width = GetPCSetting("WorkflowLogQueueDetails", "Width", Me.Width)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' Store the size of the form for retrieval when next viewed
  SavePCSetting "WorkflowLogQueueDetails", "Height", Me.Height
  SavePCSetting "WorkflowLogQueueDetails", "Width", Me.Width

End Sub


Private Sub Form_Resize()
  Dim lngColumn1Left As Long
  Dim lngColumn2Left As Long
  Dim lngColumn1DataLeft As Long
  Dim lngColumn2DataLeft As Long
  Dim sngMinFormHeight As Single

  Const GAPOVERBUTTONS = 100
  Const GAPUNDERBUTTONS = 620
  
  Const YGAP = 260

  Const COLUMN1LABELLEFT = 200
  Const COLUMN1LABELWIDTH = 1900
  Const COLUMN2LABELWIDTH = 1100
  Const GAPBETWEENCOLUMNS = 240

  DisplayApplication

  If mblnSizing Then Exit Sub

  mblnSizing = True

  Select Case miLinkType
    Case 0, 2
      ' Column, Date
      sngMinFormHeight = MINFORMHEIGHT

    Case 1
      ' Record
      fraDetails.Height = lblRecordDescription.Top _
        + lblRecordDescription.Height _
        + YGAP
        
  End Select

'  If Me.Height < sngMinFormHeight Then Me.Height = sngMinFormHeight
'  If Me.Height > Screen.Height Then Me.Height = (Screen.Height - 200)

  If mfSizeable Then
    fraDetails.Height = Me.Height _
      - fraDetails.Top _
      - cmdOK.Height _
      - GAPOVERBUTTONS _
      - GAPUNDERBUTTONS

    grdColumnValues.Height = fraDetails.Height _
      - grdColumnValues.Top _
      - 200
  End If

'  If Me.Width < MINFORMWIDTH Then Me.Width = MINFORMWIDTH
'  If Me.Width > Screen.Width Then Me.Width = (Screen.Width - 200)

  fraDetails.Width = Me.Width _
    - 350

  lngColumn1Left = COLUMN1LABELLEFT
  lngColumn1DataLeft = COLUMN1LABELLEFT + COLUMN1LABELWIDTH

  lngColumn2Left = (fraDetails.Width / 1.75)
  lngColumn2DataLeft = lngColumn2Left + COLUMN2LABELWIDTH

  ' FIRST COLUMN
  lblWorkflowName.Left = lngColumn1Left
  lblWorkflowNameValue.Left = lngColumn1DataLeft
  lblWorkflowNameValue.Width = lngColumn2Left - lngColumn1DataLeft - GAPBETWEENCOLUMNS

  lblInitiationTime.Left = lblWorkflowName.Left
  lblInitiationTimeValue.Left = lblWorkflowNameValue.Left
  lblInitiationTimeValue.Width = lblWorkflowNameValue.Width

  lblInitiator.Left = lblWorkflowName.Left
  lblInitiatorValue.Left = lblWorkflowNameValue.Left
  lblInitiatorValue.Width = lblWorkflowNameValue.Width

  lblRecordDescription.Left = lblWorkflowName.Left
  lblRecordDescriptionValue.Left = lblWorkflowNameValue.Left
  lblRecordDescriptionValue.Width = fraDetails.Width - lngColumn1DataLeft - lngColumn1Left

  If mfSizeable Then
    linColumns.Left = lngColumn1Left
    linColumns.Width = fraDetails.Width - (2 * lngColumn1Left)
    
    grdColumnValues.Left = linColumns.Left
    grdColumnValues.Width = linColumns.Width

    ResizeGridColumns grdColumnValues
  End If
  
  ' SECOND COLUMN
  lblTable.Left = lngColumn2Left
  lblTableValue.Left = lngColumn2DataLeft
  lblTableValue.Width = fraDetails.Width - lngColumn2DataLeft - lngColumn1Left

  lblLinkType.Left = lblTable.Left
  lblLinkTypeValue.Left = lblTableValue.Left
  lblLinkTypeValue.Width = lblTableValue.Width

  cmdOK.Top = fraDetails.Top _
    + fraDetails.Height _
    + GAPOVERBUTTONS
  cmdOK.Left = fraDetails.Left _
    + fraDetails.Width _
    - cmdOK.Width

  If Not mfSizeable Then
    Me.Height = cmdOK.Top _
      + cmdOK.Height _
      + GAPUNDERBUTTONS
  End If

  mblnSizing = False

End Sub


Private Sub ResizeGridColumns(pctlGrid As SSDBGrid)
  ' Size the visible columns in the given grid to fit the text.
  ' If the columns are then not as wide as the grid, stretch out the last visible column.

  Dim iLastVisibleColumn As Integer
  Dim iColumn As Integer
  Dim iRow As Integer
  Dim lngTextWidth As Long
  Dim varBookmark As Variant
  Dim varOriginalPos As Variant
  Dim fVerticalScrollRequired As Boolean
  Dim fHorizontalScrollRequired As Boolean
  
  Const SCROLLWIDTH = 255
  
  iLastVisibleColumn = -1
  lngTextWidth = 0
  
  With pctlGrid
    varOriginalPos = .Bookmark

    .Redraw = False
    .MoveFirst
    
    For iColumn = 0 To .Columns.Count - 1 Step 1
      lngTextWidth = Me.TextWidth(.Columns(iColumn).Caption)

      If .Columns(iColumn).Visible Then
        iLastVisibleColumn = iColumn
        
        For iRow = 0 To .Rows - 1 Step 1
          varBookmark = .AddItemBookmark(iRow)

          If Me.TextWidth(Trim(.Columns(iColumn).CellText(varBookmark))) > lngTextWidth Then
            lngTextWidth = Me.TextWidth(Trim(.Columns(iColumn).CellText(varBookmark)))
          End If
        Next iRow

        .Columns(iColumn).Width = lngTextWidth + 195
      End If
      lngTextWidth = 0
    Next iColumn

    If iLastVisibleColumn >= 0 Then
      ' Stretch out the last column if required
      fVerticalScrollRequired = (.Rows > .VisibleRows)
      
      If .Columns(iLastVisibleColumn).Left + .Columns(iLastVisibleColumn).Width _
        < (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) Then
      
        .Columns(iLastVisibleColumn).Width = _
          (.Width - IIf(fVerticalScrollRequired, SCROLLWIDTH, 0)) - .Columns(iLastVisibleColumn).Left - 25
      End If
    End If
    
    .Bookmark = varOriginalPos
    .Redraw = True
  End With

End Sub


Public Function Initialise(plngQueueID As Long) As Boolean

  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim rstDetails As Recordset
  Dim rstColumns As Recordset
  Dim sSQL As String
  Dim sDateFormat As String
  Dim sRecDesc As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim sAddString As String
  
  fOK = True
  sDateFormat = DateFormat

  ' Get the queue item details
  sSQL = "SELECT wf.name," & _
    "    wfq.dateDue," & _
    "    '<Triggered>' AS [Username]," & _
    "    isnull(wfq.recordDesc, '') AS [recordDesc]," & _
    "    wfq.recalculateRecordDesc," & _
    "    isnull(wfq.recordID, 0) AS [recordID]," & _
    "    REPLACE(t.tableName, '_', ' ') AS [tableName]," & _
    "    isnull(t.recordDescExprID, 0) AS [recordDescExprID]," & _
    "    wftl.type," & _
    "    CASE" & _
    "      WHEN wftl.type = 0 THEN 'Column'" & _
    "      WHEN wftl.type = 1 THEN 'Record'" & _
    "      WHEN wftl.type = 2 THEN 'Date'" & _
    "      ELSE '<unknown>'" & _
    "    END AS [linkType]" & _
    "  FROM asrsysworkflowqueue wfq" & _
    "  INNER JOIN ASRSysWorkflowTriggeredLinks wftl ON wfq.linkID = wftl.linkID" & _
    "  INNER JOIN ASRSysWorkflows wf ON wftl.workflowID = wf.ID" & _
    "  INNER JOIN ASRSysTables t ON wftl.tableID = t.tableID" & _
    "  WHERE wfq.queueID = " & CStr(plngQueueID)

  Set rstDetails = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  fOK = Not (rstDetails.EOF And rstDetails.BOF)
  
  If fOK Then
    ' Populate the screen controls
    With rstDetails
      lblWorkflowNameValue.Caption = Replace(.Fields("name"), "&", "&&")
      lblInitiationTimeValue.Caption = Format(.Fields("dateDue"), sDateFormat & " hh:nn")
      lblInitiatorValue.Caption = .Fields("userName")

      sRecDesc = .Fields("recordDesc")
      If .Fields("recalculateRecordDesc") _
        And (.Fields("recordDescExprID") > 0) _
        And (.Fields("recordID") > 0) Then
        
        Set cmADO = New ADODB.Command
        With cmADO
          .CommandText = "sp_ASRExpr_" & Trim(Str(rstDetails.Fields("recordDescExprID")))
          .CommandType = adCmdStoredProc
          .CommandTimeout = 0
          Set .ActiveConnection = gADOCon
                
          Set pmADO = .CreateParameter("recordDescription", adVarChar, adParamOutput, VARCHAR_MAX_Size)
          .Parameters.Append pmADO
        
          Set pmADO = .CreateParameter("currentID", adInteger, adParamInput)
          .Parameters.Append pmADO
          pmADO.Value = rstDetails.Fields("recordID")
          
          Set pmADO = Nothing
              
          cmADO.Execute
        
          If Len(.Parameters("recordDescription").Value) > 0 Then
            sRecDesc = .Parameters("recordDescription").Value
          End If
        End With
        Set cmADO = Nothing
      
      End If
      lblRecordDescriptionValue.Caption = sRecDesc
      lblTableValue.Caption = .Fields("tableName")
      lblLinkTypeValue.Caption = .Fields("linkType")
      miLinkType = .Fields("type")
    End With
    
    If (miLinkType = 0) _
      Or (miLinkType = 2) Then
      ' Column or Date type link. ie. has column info attached.
      
      sSQL = "SELECT WFQC.columnValue," & _
        "  replace(C.columnName, '_', ' ') AS [columnName]," & _
        "  C.dataType" & _
        "  FROM ASRSysWorkflowQueueColumns WFQC" & _
        "  INNER JOIN ASRSysColumns C ON WFQC.columnID = C.columnID" & _
        "  WHERE WFQC.queueID = " & CStr(plngQueueID) & _
        "  ORDER BY [columnName]"
      Set rstColumns = mclsData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
      Do While Not rstColumns.EOF
        sAddString = rstColumns.Fields("columnName") & vbTab
        
        Select Case rstColumns.Fields("dataType")
          Case sqlBoolean:
            ' Logic
            sAddString = sAddString & IIf(rstColumns.Fields("columnValue") = "1", "True", "False")

          Case sqlDate:
            ' sqlDate
            If Len(rstColumns.Fields("columnValue")) > 0 Then
              sAddString = sAddString & ConvertSQLDateToLocale(rstColumns.Fields("columnValue"))
            Else
              sAddString = sAddString & "<null>"
            End If
  
          Case sqlNumeric, sqlInteger:
            ' Numeric/Integer
            sAddString = sAddString & datGeneral.ConvertNumberForDisplay(rstColumns.Fields("columnValue"))

          Case Else
            sAddString = sAddString & rstColumns.Fields("columnValue")
        End Select
    
        grdColumnValues.AddItem sAddString
        
        rstColumns.MoveNext
      Loop
    
      rstColumns.Close
      Set rstColumns = Nothing
    
      ResizeGridColumns grdColumnValues
    End If
  End If
  
  rstDetails.Close
  Set rstDetails = Nothing
  
  If fOK Then
    mfSizeable = (miLinkType <> 1)
    linColumns.Visible = mfSizeable
    grdColumnValues.Visible = mfSizeable
    
    Form_Resize
  End If
  
TidyUpAndExit:
  Initialise = fOK
  
  Exit Function

ErrorTrap:
  fOK = False
  COAMsgBox "Error retrieving detail entries for this workflow." & vbCrLf & "(" & Err.Description & ")", vbExclamation + vbOKOnly, "Workflow Log"
  GoTo TidyUpAndExit

End Function




Private Sub Form_Unload(Cancel As Integer)
  Set mclsData = Nothing

  Unhook Me.hWnd
End Sub



