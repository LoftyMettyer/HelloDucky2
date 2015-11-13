VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWorkflowOpen 
   Caption         =   "Workflow Designer"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   5051
   Icon            =   "frmWorkflowOpen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   5535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export..."
      Height          =   400
      Left            =   4245
      TabIndex        =   15
      Top             =   3180
      Width           =   1200
   End
   Begin VB.Frame fraDetails 
      Height          =   1125
      Left            =   150
      TabIndex        =   0
      Top             =   50
      Width           =   4000
      Begin VB.ComboBox cboInitiationType 
         Height          =   315
         ItemData        =   "frmWorkflowOpen.frx":000C
         Left            =   1575
         List            =   "frmWorkflowOpen.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   255
         Width           =   2325
      End
      Begin VB.ComboBox cboBaseTable 
         Height          =   315
         Left            =   1575
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   2325
      End
      Begin VB.Label lblInitiationType 
         BackStyle       =   0  'Transparent
         Caption         =   "Initiation Type :"
         Height          =   195
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   1365
      End
      Begin VB.Label lblBaseTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Table :"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   705
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   4245
      TabIndex        =   13
      Top             =   5550
      Width           =   1200
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Proper&ties..."
      Height          =   400
      Left            =   4245
      TabIndex        =   11
      Top             =   2150
      Width           =   1200
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   400
      Left            =   4245
      TabIndex        =   10
      Top             =   1650
      Width           =   1200
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Edit..."
      Height          =   400
      Left            =   4245
      TabIndex        =   8
      Top             =   650
      Width           =   1200
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New..."
      Height          =   400
      Left            =   4245
      TabIndex        =   7
      Top             =   150
      Width           =   1200
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Cop&y..."
      Height          =   400
      Left            =   4245
      TabIndex        =   9
      Top             =   1150
      Width           =   1200
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print..."
      Height          =   400
      Left            =   4245
      TabIndex        =   12
      Top             =   2650
      Width           =   1200
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000F&
      Height          =   1000
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4950
      Width           =   4000
   End
   Begin ComctlLib.ListView lstItems 
      Height          =   3575
      Left            =   150
      TabIndex        =   5
      Top             =   1275
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   6297
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   "Column"
         Object.Tag             =   "Column"
         Text            =   "Column"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SortKey"
         Object.Width           =   0
      EndProperty
   End
   Begin ComctlLib.StatusBar sbScrOpen 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   5970
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9234
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWorkflowOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declare events
Event UnLoad()

Private gfLoading As Boolean
Private mblnReadOnly As Boolean

Private mavWorkflowInfo() As Variant

Private mlngSelectedWorkflowID As Long

Private miInitiationType As WorkflowInitiationTypes
Private mlngBaseTableID As Long

Private mlngPersModulePersonnelTableID As Long

Private mblnWorkflowEnabled As Boolean

Private Const MIN_FORM_HEIGHT = 5000
Private Const MIN_FORM_WIDTH = 6000

Private Function GetWorkflowForm() As Form
  ' Return the workflow designer form for the current workflow ID, if it exists.
  Dim iLoop As Integer
  
  For iLoop = 1 To Forms.Count - 1
    If Forms(iLoop).Name = "frmWorkflowDesigner" Then
      If Forms(iLoop).WorkflowID = WorkflowID Then
        Set GetWorkflowForm = Forms(iLoop)
        Exit For
      End If
    End If
  Next iLoop

End Function

Public Sub SelectWorkflow()
  Dim iCount As Integer
  Dim fFound As Boolean
  
  fFound = False
  
  For iCount = 1 To lstItems.ListItems.Count
    If lstItems.ListItems(iCount).Tag = mlngSelectedWorkflowID Then
      fFound = True
      Set lstItems.SelectedItem = lstItems.ListItems(iCount)
      Exit For
    End If
  Next iCount
  
  If (Not fFound) And (lstItems.ListItems.Count > 0) Then
    Set lstItems.SelectedItem = lstItems.ListItems(1)
  End If
  
  RefreshControls
  
End Sub

Public Function ValidateWorkflow(pfSaving As Boolean, _
  pfSilent As Boolean, _
  pfFix As Boolean) As Boolean
  ' Validate the workflow.
  On Error GoTo ErrorTrap

  Dim fValid As Boolean
  Dim frmWFDes As frmWorkflowDesigner

  fValid = True

  If (WorkflowID > 0) Then
    ' Instantiate (but hide) the Workflow designer form for the selected Workflow.
    Set frmWFDes = GetWorkflowForm
    If frmWFDes Is Nothing Then
      Set frmWFDes = New frmWorkflowDesigner
      With frmWFDes
        .IsNew = False
        .WorkflowID = WorkflowID
      End With
    End If

    fValid = frmWFDes.ValidateWorkflow(pfSaving, pfSilent, pfFix)
    
    UnLoad frmWFDes

    DoEvents
  End If

TidyUpAndExit:
  ValidateWorkflow = fValid
  
  Exit Function

ErrorTrap:
  Resume TidyUpAndExit

End Function

Private Sub cboBaseTable_Click()
  mlngBaseTableID = cboBaseTable.ItemData(cboBaseTable.ListIndex)

  RefreshWorkflows
  SelectWorkflow

End Sub


Private Sub cboInitiationType_Click()
  miInitiationType = cboInitiationType.ItemData(cboInitiationType.ListIndex)
  cboBaseTable_refresh

End Sub


Private Sub cmdCopy_Click()
  ' Copy the selected screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fGoodName As Boolean
  Dim iCounter As Integer
  Dim lngWorkflowID As Long
  Dim lngID As Long
  Dim sSQL As String
  Dim sWorkflowName As String
  Dim rsWorkflow As DAO.Recordset
  Dim rsElements As DAO.Recordset
  Dim rsLinks As DAO.Recordset
  Dim rsElementItems As DAO.Recordset
  Dim rsElementItemValues As DAO.Recordset
  Dim rsElementColumns As DAO.Recordset
  Dim rsElementValidations As DAO.Recordset
  Dim alngElementIDs() As Long
  Dim iLoop As Integer
  Dim sSQL2 As String
  Dim rsFilters As DAO.Recordset
  Dim rsExpressions As DAO.Recordset
  Dim objSourceExpr As CExpression
  Dim objNewExpr As CExpression
  Dim avCloneRegister() As Variant
  Dim iIndex As Integer
  
  fOK = True
  ReDim alngElementIDs(1, 0)
  ReDim avCloneRegister(3, 0)
  
  ' Show user the system is busy...this operation could take some time...
  Screen.MousePointer = vbHourglass

  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans

  ' Get the selected workflow's definitions from the database.
  sSQL = "SELECT *" & _
    " FROM tmpWorkflows" & _
    " WHERE tmpWorkflows.ID = " & Trim(Str(WorkflowID))
  Set rsWorkflow = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  With rsWorkflow
    ' Create a new unique workflow name.
    sWorkflowName = "Copy of " & .Fields("Name")
    iCounter = 1
    fGoodName = False
    Do While Not fGoodName
      With recWorkflowEdit
        .Index = "idxName"
        .Seek "=", sWorkflowName, False
        If Not .NoMatch Then
          iCounter = iCounter + 1
          sWorkflowName = "Copy_" & Trim(Str(iCounter)) & "_of_" & rsWorkflow.Fields("Name")
        Else
          fGoodName = True
        End If
      End With
    Loop

    ' Get a unique ID for the new record.
    lngWorkflowID = UniqueColumnValue("tmpWorkflows", "ID")

    ' Add a new record in the database for the copied screen definition.
    recWorkflowEdit.AddNew

    recWorkflowEdit!ID = lngWorkflowID
    recWorkflowEdit!Changed = False
    recWorkflowEdit!perge = False
    recWorkflowEdit!New = True
    recWorkflowEdit!Deleted = False
    recWorkflowEdit!Name = sWorkflowName
    recWorkflowEdit!Description = .Fields("description")
    recWorkflowEdit!PictureID = .Fields("PictureID")
    recWorkflowEdit!Enabled = False
    recWorkflowEdit!InitiationType = IIf(IsNull(.Fields("InitiationType")), WORKFLOWINITIATIONTYPE_MANUAL, .Fields("InitiationType"))
    recWorkflowEdit!BaseTable = IIf(IsNull(.Fields("BaseTable")), 0, .Fields("BaseTable"))
    recWorkflowEdit!queryString = IIf(IsNull(.Fields("InitiationType")), "", _
      IIf(.Fields("InitiationType") = WORKFLOWINITIATIONTYPE_EXTERNAL, GetWorkflowQueryString(lngWorkflowID * -1, -1), ""))

    recWorkflowEdit.Update

    sSQL = "SELECT tmpExpressions.exprID" & _
      " FROM tmpExpressions" & _
      " WHERE tmpExpressions.deleted = FALSE" & _
      " AND tmpExpressions.utilityID = " & Trim(Str(WorkflowID)) & _
      " AND tmpExpressions.parentComponentID = 0 " & _
      " AND (tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWCALCULATION) & _
      "   OR tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
      "   OR tmpExpressions.Type = " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER) & ")"
    Set rsExpressions = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
    With rsExpressions
      ' For each expression definition ...
      Do While (Not .EOF) And fOK
        ' Instantiate the original expression object.
        Set objSourceExpr = New CExpression
        objSourceExpr.ExpressionID = .Fields("exprID")
    
        Set objNewExpr = objSourceExpr.CloneExpression(avCloneRegister)
        fOK = Not objNewExpr Is Nothing
    
        If fOK Then
          ' Copy properties from the original expression to the copy.
          objNewExpr.UtilityID = lngWorkflowID
          ' Write the copied expession definition to the database.
          fOK = objNewExpr.WriteExpression
        End If
    
        ' Remember the IDs of the original and copied orders.
        'If fOK Then
        iIndex = UBound(avCloneRegister, 2) + 1
        ReDim Preserve avCloneRegister(3, iIndex)
        avCloneRegister(1, iIndex) = "EXPRESSION"
        avCloneRegister(2, iIndex) = objSourceExpr.ExpressionID
    
        If fOK Then
          avCloneRegister(3, iIndex) = objNewExpr.ExpressionID
        Else
          avCloneRegister(3, iIndex) = 0
    
          fOK = True
        End If
    
        ' Disassociate object variables.
        Set objSourceExpr = Nothing
        Set objNewExpr = Nothing
    
        .MoveNext
      Loop
    
      .Close
    End With
    ' Disassociate object variables.
    Set rsExpressions = Nothing
    
    If fOK Then
      ' Update the copied expression field components with the new IDs of their filter expressions.
      sSQL = "SELECT tmpComponents.componentID, tmpComponents.fieldSelectionFilter" & _
        " FROM tmpComponents, tmpExpressions " & _
        " WHERE tmpExpressions.utilityID = " & Trim(Str(lngWorkflowID)) & _
        " AND tmpComponents.exprID = tmpExpressions.exprID" & _
        " AND (tmpComponents.type = " & Trim(Str(giCOMPONENT_FIELD)) & _
        "   OR tmpComponents.type = " & Trim(Str(giCOMPONENT_WORKFLOWFIELD)) & ")" & _
        " AND tmpComponents.fieldSelectionFilter > 0"
      Set rsFilters = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    
      With rsFilters
        Do While (Not .EOF)
          For iIndex = 1 To UBound(avCloneRegister, 2)
            If avCloneRegister(1, iIndex) = "EXPRESSION" And _
              avCloneRegister(2, iIndex) = .Fields("fieldSelectionFilter") Then
    
              recCompEdit.Index = "idxCompID"
              recCompEdit.Seek "=", .Fields("componentID")
              If Not recCompEdit.NoMatch Then
                recCompEdit.Edit
                recCompEdit.Fields("fieldSelectionFilter") = avCloneRegister(3, iIndex)
                recCompEdit.Update
              End If
              Exit For
            End If
          Next iIndex
    
          .MoveNext
        Loop
    
        .Close
      End With
    
      ' Disassociate object variables.
      Set rsFilters = Nothing
    End If

    ' Copy the workflow element definitions.
    sSQL = "SELECT *" & _
      " FROM tmpWorkflowElements" & _
      " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(.Fields("ID")))
    Set rsElements = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElements
      ' For each workflow element definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied screen control definition.
        recWorkflowElementEdit.AddNew

        lngID = UniqueColumnValue("tmpWorkflowElements", "ID")
        recWorkflowElementEdit!ID = lngID

        recWorkflowElementEdit!WorkflowID = lngWorkflowID
        recWorkflowElementEdit!Type = .Fields("Type")
        recWorkflowElementEdit!Caption = .Fields("Caption")
        recWorkflowElementEdit!ConnectionPairID = .Fields("connectionPairID")
        recWorkflowElementEdit!LeftCoord = .Fields("leftCoord")
        recWorkflowElementEdit!TopCoord = .Fields("topCoord")
        recWorkflowElementEdit!Identifier = .Fields("Identifier")
        
        recWorkflowElementEdit!DecisionCaptionType = .Fields("DecisionCaptionType")
        recWorkflowElementEdit!TrueFlowType = .Fields("TrueFlowType")
        recWorkflowElementEdit!TrueFlowIdentifier = .Fields("TrueFlowIdentifier")
        
        recWorkflowElementEdit!TrueFlowExprID = .Fields("TrueFlowExprID")
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("TrueFlowExprID") Then
        
            recWorkflowElementEdit!TrueFlowExprID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        
        recWorkflowElementEdit!DescriptionExprID = .Fields("DescriptionExprID")
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("DescriptionExprID") Then
        
            recWorkflowElementEdit!DescriptionExprID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        
        recWorkflowElementEdit!DescHasWorkflowName = .Fields("DescHasWorkflowName")
        recWorkflowElementEdit!DescHasElementCaption = .Fields("DescHasElementCaption")
        
        recWorkflowElementEdit!DataAction = .Fields("DataAction")
        recWorkflowElementEdit!DataTableID = .Fields("DataTableID")
        recWorkflowElementEdit!DataRecord = .Fields("DataRecord")
        recWorkflowElementEdit!EmailID = .Fields("EmailID")
        recWorkflowElementEdit!EmailCCID = .Fields("EmailCCID")
        recWorkflowElementEdit!EmailRecord = .Fields("EmailRecord")
                
        recWorkflowElementEdit!WebFormFGColor = .Fields("WebFormFGColor")
        recWorkflowElementEdit!WebFormBGColor = .Fields("WebFormBGColor")
        recWorkflowElementEdit!WebFormBGImageID = .Fields("WebFormBGImageID")
        recWorkflowElementEdit!WebFormBGImageLocation = .Fields("WebFormBGImageLocation")
        recWorkflowElementEdit!WebFormDefaultFontName = .Fields("webFormDefaultFontName")
        recWorkflowElementEdit!WebFormDefaultFontSize = .Fields("webFormDefaultFontSize")
        recWorkflowElementEdit!WebFormDefaultFontBold = .Fields("webFormDefaultFontBold")
        recWorkflowElementEdit!WebFormDefaultFontItalic = .Fields("webFormDefaultFontItalic")
        recWorkflowElementEdit!WebFormDefaultFontStrikeThru = .Fields("webFormDefaultFontStrikeThru")
        recWorkflowElementEdit!WebFormDefaultFontUnderline = .Fields("webFormDefaultFontUnderline")
        recWorkflowElementEdit!WebFormWidth = .Fields("WebFormWidth")
        recWorkflowElementEdit!WebFormHeight = .Fields("WebFormHeight")
        recWorkflowElementEdit!RecSelWebFormIdentifier = .Fields("recSelWebFormIdentifier")
        recWorkflowElementEdit!RecSelIdentifier = .Fields("recSelIdentifier")
        
        recWorkflowElementEdit!SecondaryDataRecord = .Fields("SecondaryDataRecord")
        recWorkflowElementEdit!SecondaryRecSelWebFormIdentifier = .Fields("secondaryRecSelWebFormIdentifier")
        recWorkflowElementEdit!SecondaryRecSelIdentifier = .Fields("secondaryRecSelIdentifier")
        
        recWorkflowElementEdit!DataRecordTable = .Fields("DataRecordTable")
        recWorkflowElementEdit!SecondaryDataRecordTable = .Fields("secondaryDataRecordTable")
        recWorkflowElementEdit!UseAsTargetIdentifier = .Fields("UseAsTargetIdentifier")
        recWorkflowElementEdit!RequiresAuthentication = .Fields("RequiresAuthentication")
               
        'JPD 20060908 Fault 11482
        recWorkflowElementEdit!EMailSubject = .Fields("EmailSubject")
        recWorkflowElementEdit!TimeoutFrequency = .Fields("TimeoutFrequency")
        recWorkflowElementEdit!TimeoutPeriod = .Fields("TimeoutPeriod")
        recWorkflowElementEdit!TimeoutExcludeWeekend = .Fields("TimeoutExcludeWeekend")
        
        recWorkflowElementEdit!CompletionMessageType = .Fields("CompletionMessageType")
        recWorkflowElementEdit!CompletionMessage = .Fields("CompletionMessage")
        recWorkflowElementEdit!SavedForLaterMessageType = .Fields("SavedForLaterMessageType")
        recWorkflowElementEdit!SavedForLaterMessage = .Fields("SavedForLaterMessage")
        recWorkflowElementEdit!FollowOnFormsMessageType = .Fields("FollowOnFormsMessageType")
        recWorkflowElementEdit!FollowOnFormsMessage = .Fields("FollowOnFormsMessage")
        
        recWorkflowElementEdit!Attachment_Type = .Fields("Attachment_Type")
        recWorkflowElementEdit!Attachment_File = .Fields("Attachment_File")
        recWorkflowElementEdit!Attachment_WFElementIdentifier = .Fields("Attachment_WFElementIdentifier")
        recWorkflowElementEdit!Attachment_WFValueIdentifier = .Fields("Attachment_WFValueIdentifier")
        recWorkflowElementEdit!Attachment_DBColumnID = .Fields("Attachment_DBColumnID")
        recWorkflowElementEdit!Attachment_DBRecord = .Fields("Attachment_DBRecord")
        recWorkflowElementEdit!Attachment_DBElement = .Fields("Attachment_DBElement")
        recWorkflowElementEdit!Attachment_DBValue = .Fields("Attachment_DBValue")
                
        recWorkflowElementEdit.Update

        ReDim Preserve alngElementIDs(1, UBound(alngElementIDs, 2) + 1)
        alngElementIDs(0, UBound(alngElementIDs, 2)) = .Fields("ID")
        alngElementIDs(1, UBound(alngElementIDs, 2)) = lngID

        .MoveNext
      Loop
    End With
    Set rsElements = Nothing

    ' Ensure the connector elements have the new IDs.
    For iLoop = 1 To UBound(alngElementIDs, 2)
      sSQL = "UPDATE tmpWorkflowElements" & _
        " SET tmpWorkflowElements.connectionPairID = " & Trim(Str(alngElementIDs(1, iLoop))) & _
        " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(lngWorkflowID)) & _
        "   AND tmpWorkflowElements.connectionPairID = " & Trim(Str(alngElementIDs(0, iLoop)))
    
      daoDb.Execute sSQL, dbFailOnError
    Next iLoop
    
    ' Copy the workflow element item definitions.
    sSQL = "SELECT tmpWorkflowElementItems.*" & _
      " FROM tmpWorkflowElementItems" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementItems.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(.Fields("ID")))
    Set rsElementItems = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementItems
      ' For each element item definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element item definition.
        recWorkflowElementItemEdit.AddNew

        lngID = UniqueColumnValue("tmpWorkflowElementItems", "ID")
        recWorkflowElementItemEdit!ID = lngID
       
        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementItemEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop
        
        recWorkflowElementItemEdit!UseAsTargetIdentifier = .Fields("UseAsTargetIdentifier")
        recWorkflowElementItemEdit!Caption = .Fields("Caption")
        recWorkflowElementItemEdit!DBColumnID = .Fields("DBColumnID")
        recWorkflowElementItemEdit!DBRecord = .Fields("DBRecord")
        recWorkflowElementItemEdit!Identifier = .Fields("Identifier")
        recWorkflowElementItemEdit!InputType = .Fields("InputType")
        recWorkflowElementItemEdit!InputSize = .Fields("InputSize")
        recWorkflowElementItemEdit!InputDecimals = .Fields("InputDecimals")
        recWorkflowElementItemEdit!InputDefault = .Fields("InputDefault")
        recWorkflowElementItemEdit!WFFormIdentifier = .Fields("WFFormIdentifier")
        recWorkflowElementItemEdit!WFValueIdentifier = .Fields("WFValueIdentifier")
        recWorkflowElementItemEdit!ItemType = .Fields("ItemType")

        recWorkflowElementItemEdit!LeftCoord = .Fields("LeftCoord")
        recWorkflowElementItemEdit!TopCoord = .Fields("TopCoord")
        recWorkflowElementItemEdit!Width = .Fields("Width")
        recWorkflowElementItemEdit!Height = .Fields("Height")
        recWorkflowElementItemEdit!BackColor = .Fields("BackColor")
        recWorkflowElementItemEdit!ForeColor = .Fields("ForeColor")
        recWorkflowElementItemEdit!FontName = .Fields("FontName")
        recWorkflowElementItemEdit!FontSize = .Fields("FontSize")
        recWorkflowElementItemEdit!FontBold = .Fields("FontBold")
        recWorkflowElementItemEdit!FontItalic = .Fields("FontItalic")
        recWorkflowElementItemEdit!FontStrikethru = .Fields("FontStrikeThru")
        recWorkflowElementItemEdit!FontUnderline = .Fields("FontUnderline")
        recWorkflowElementItemEdit!PictureID = .Fields("PictureID")
        recWorkflowElementItemEdit!PictureBorder = .Fields("PictureBorder")
        recWorkflowElementItemEdit!Alignment = .Fields("Alignment")
        recWorkflowElementItemEdit!ZOrder = .Fields("ZOrder")
        recWorkflowElementItemEdit!TabIndex = .Fields("TabIndex")
        recWorkflowElementItemEdit!BackStyle = .Fields("BackStyle")
        recWorkflowElementItemEdit!BackColorEven = .Fields("BackColorEven")
        recWorkflowElementItemEdit!BackColorOdd = .Fields("BackColorOdd")
        recWorkflowElementItemEdit!ColumnHeaders = .Fields("ColumnHeaders")
        recWorkflowElementItemEdit!ForeColorEven = .Fields("ForeColorEven")
        recWorkflowElementItemEdit!ForeColorOdd = .Fields("ForeColorOdd")
        recWorkflowElementItemEdit!HeaderBackColor = .Fields("HeaderBackColor")
        recWorkflowElementItemEdit!HeadFontName = .Fields("HeadFontName")
        recWorkflowElementItemEdit!HeadFontSize = .Fields("HeadFontSize")
        recWorkflowElementItemEdit!HeadFontBold = .Fields("HeadFontBold")
        recWorkflowElementItemEdit!HeadFontItalic = .Fields("HeadFontItalic")
        recWorkflowElementItemEdit!HeadFontStrikeThru = .Fields("HeadFontStrikeThru")
        recWorkflowElementItemEdit!HeadFontUnderline = .Fields("HeadFontUnderline")
        recWorkflowElementItemEdit!HeadLines = .Fields("Headlines")
        recWorkflowElementItemEdit!TableID = .Fields("TableID")
        recWorkflowElementItemEdit!RecSelWebFormIdentifier = .Fields("recSelWebFormIdentifier")
        recWorkflowElementItemEdit!RecSelIdentifier = .Fields("recSelIdentifier")

        recWorkflowElementItemEdit!ForeColorHighlight = .Fields("ForeColorHighlight")
        recWorkflowElementItemEdit!BackColorHighlight = .Fields("BackColorHighlight")

        recWorkflowElementItemEdit!LookupTableID = .Fields("LookupTableID")
        recWorkflowElementItemEdit!LookupColumnID = .Fields("LookupColumnID")

        recWorkflowElementItemEdit!RecordTableID = .Fields("RecordTableID")

        recWorkflowElementItemEdit!Orientation = .Fields("Orientation")
        recWorkflowElementItemEdit!RecordOrderID = .Fields("RecordOrderID")
        recWorkflowElementItemEdit!RecordFilterID = .Fields("RecordFilterID")
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("RecordFilterID") Then
        
            recWorkflowElementItemEdit!RecordFilterID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex

        recWorkflowElementItemEdit!Behaviour = .Fields("behaviour")
        recWorkflowElementItemEdit!Mandatory = .Fields("mandatory")
        
        recWorkflowElementItemEdit!CalcID = .Fields("calcID")
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("calcID") Then
        
            recWorkflowElementItemEdit!CalcID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        
        recWorkflowElementItemEdit!CaptionType = .Fields("captionType")
        recWorkflowElementItemEdit!DefaultValueType = .Fields("defaultValueType")
        
        recWorkflowElementItemEdit!VerticalOffset = .Fields("VerticalOffset")
        recWorkflowElementItemEdit!VerticalOffsetBehaviour = .Fields("VerticalOffsetBehaviour")
        recWorkflowElementItemEdit!HorizontalOffset = .Fields("HorizontalOffset")
        recWorkflowElementItemEdit!HorizontalOffsetBehaviour = .Fields("HorizontalOffsetBehaviour")
        recWorkflowElementItemEdit!HeightBehaviour = .Fields("HeightBehaviour")
        recWorkflowElementItemEdit!WidthBehaviour = .Fields("WidthBehaviour")
        recWorkflowElementItemEdit!PasswordType = .Fields("PasswordType")

        recWorkflowElementItemEdit!LookupFilterColumnID = .Fields("LookupFilterColumnID")
        recWorkflowElementItemEdit!LookupFilterOperator = .Fields("LookupFilterOperator")
        recWorkflowElementItemEdit!LookupFilterValue = .Fields("LookupFilterValue")
        recWorkflowElementItemEdit!LookupOrderID = .Fields("LookupOrderID")
        recWorkflowElementItemEdit!HotSpotIdentifier = .Fields("HotSpotIdentifier")
        
        recWorkflowElementItemEdit!PageNo = .Fields("PageNo")
        recWorkflowElementItemEdit!ButtonStyle = .Fields("ButtonStyle")

        recWorkflowElementItemEdit.Update

        ' Copy the workflow element item definitions.
        sSQL2 = "SELECT tmpWorkflowElementItemValues.*" & _
          " FROM tmpWorkflowElementItemValues" & _
          " WHERE tmpWorkflowElementItemValues.itemID = " & Trim(Str(.Fields("ID")))
        Set rsElementItemValues = daoDb.OpenRecordset(sSQL2, dbOpenForwardOnly, dbReadOnly)

        With rsElementItemValues
          ' For each element item value definition ...
          Do While (Not .EOF)
            ' Add a new record in the database for the copied element item definition.
            recWorkflowElementItemValuesEdit.AddNew

            recWorkflowElementItemValuesEdit!itemID = lngID

            recWorkflowElementItemValuesEdit!value = .Fields("Value")
            recWorkflowElementItemValuesEdit!Sequence = .Fields("Sequence")

            recWorkflowElementItemValuesEdit.Update

            .MoveNext
          Loop
        End With
        ' Disassociate object variables.
        Set rsElementItemValues = Nothing

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementItems = Nothing

    ' Copy the workflow element column definitions.
    sSQL = "SELECT tmpWorkflowElementColumns.*" & _
      " FROM tmpWorkflowElementColumns" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementColumns.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(.Fields("ID")))
    Set rsElementColumns = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementColumns
      ' For each element column definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element column definition.
        recWorkflowElementColumnEdit.AddNew

        lngID = UniqueColumnValue("tmpWorkflowElementColumns", "ID")
        recWorkflowElementColumnEdit!ID = lngID
       
        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementColumnEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        recWorkflowElementColumnEdit!ColumnID = .Fields("columnID")
        recWorkflowElementColumnEdit!ValueType = .Fields("ValueType")
        recWorkflowElementColumnEdit!value = .Fields("Value")
        recWorkflowElementColumnEdit!WFFormIdentifier = .Fields("WFFormIdentifier")
        recWorkflowElementColumnEdit!WFValueIdentifier = .Fields("WFValueIdentifier")

        recWorkflowElementColumnEdit!DBColumnID = .Fields("DBColumnID")
        recWorkflowElementColumnEdit!DBRecord = .Fields("DBRecord")
        recWorkflowElementColumnEdit!CalcID = .Fields("CalcID")
        
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("calcID") Then

            recWorkflowElementColumnEdit!CalcID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex

        recWorkflowElementColumnEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementColumns = Nothing

    ' Copy the workflow element validation definitions.
    sSQL = "SELECT tmpWorkflowElementValidations.*" & _
      " FROM tmpWorkflowElementValidations" & _
      " INNER JOIN tmpWorkflowElements ON tmpWorkflowElementValidations.elementID = tmpWorkflowElements.ID" & _
      " WHERE tmpWorkflowElements.workflowID = " & Trim(Str(.Fields("ID")))
    Set rsElementValidations = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsElementValidations
      ' For each element validation definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied element validation definition.
        recWorkflowElementValidationEdit.AddNew

        lngID = UniqueColumnValue("tmpWorkflowElementValidations", "ID")
        recWorkflowElementValidationEdit!ID = lngID

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("ElementID") Then
            recWorkflowElementValidationEdit!elementid = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        recWorkflowElementValidationEdit!ExprID = .Fields("exprID")
        For iIndex = 1 To UBound(avCloneRegister, 2)
          If avCloneRegister(1, iIndex) = "EXPRESSION" And _
            avCloneRegister(2, iIndex) = .Fields("exprID") Then
        
            recWorkflowElementValidationEdit!ExprID = avCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
        
        recWorkflowElementValidationEdit!Type = .Fields("Type")
        recWorkflowElementValidationEdit!Message = .Fields("Message")

        recWorkflowElementValidationEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsElementValidations = Nothing

    ' Copy the workflow link definitions.
    sSQL = "SELECT *" & _
      " FROM tmpWorkflowLinks" & _
      " WHERE tmpWorkflowLinks.workflowID = " & Trim(Str(.Fields("ID")))
    Set rsLinks = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

    With rsLinks
      ' For each link definition ...
      Do While (Not .EOF)
        ' Add a new record in the database for the copied link definition.
        recWorkflowLinkEdit.AddNew

        lngID = UniqueColumnValue("tmpWorkflowLinks", "ID")
        recWorkflowLinkEdit!ID = lngID

        recWorkflowLinkEdit!WorkflowID = lngWorkflowID
        recWorkflowLinkEdit!StartOutboundFlowCode = .Fields("StartOutboundFlowCode")

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("StartElementID") Then
            recWorkflowLinkEdit!StartElementID = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        For iLoop = 1 To UBound(alngElementIDs, 2)
          If alngElementIDs(0, iLoop) = .Fields("EndElementID") Then
            recWorkflowLinkEdit!EndElementID = alngElementIDs(1, iLoop)
            Exit For
          End If
        Next iLoop

        recWorkflowLinkEdit.Update

        .MoveNext
      Loop
    End With
    ' Disassociate object variables.
    Set rsLinks = Nothing

    .Close
  End With
  ' Disassociate object variables.
  Set rsWorkflow = Nothing

TidyUpAndExit:
  ' Disassociate object variables.
  Set rsWorkflow = Nothing
  Set rsElements = Nothing
  Set rsLinks = Nothing
  Set rsElementItems = Nothing
  Set rsElementItemValues = Nothing
  Set rsElementColumns = Nothing
  Set rsElementValidations = Nothing

  ' Show user the system has finished working
  Screen.MousePointer = vbDefault

  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
    RefreshWorkflows
    
    mlngSelectedWorkflowID = lngWorkflowID
    SelectWorkflow
  Else
    daoWS.Rollback
    MsgBox "Unable to copy the workflow." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
End Sub

Private Sub cmdDelete_Click()
  ' Delete the selected workflow.
  ' Effectively do a 'WorkflowIsUsed' function here.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As New ADODB.Recordset
  Dim rsLocalInfo As DAO.Recordset
  Dim iRecCount As Integer
  Dim objExpression As CExpression
  Dim sMsg As String
  
  fOK = (WorkflowID > 0)

  If fOK Then
    ' Check if the workflow is in use.
    If WorkflowsWithStatus(WorkflowID, giWFSTATUS_INPROGRESS) Then
      
      MsgBox "The '" & lstItems.SelectedItem.Text & "' workflow cannot be deleted." & vbCr & _
        "There are instances of this workflow in progress.", _
        vbExclamation + vbOKOnly, Me.Caption
      fOK = False
    End If
  End If

  If fOK Then
    ' Check if the workflow is used in SSI.
    sSQL = "SELECT COUNT(*) AS recCount" & _
      " FROM tmpSSIntranetLinks" & _
      " WHERE utilityType = " & CStr(utlWorkflow) & _
      "   AND utilityID = " & CStr(WorkflowID)
    Set rsLocalInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    If rsLocalInfo!reccount > 0 Then
      MsgBox "The '" & lstItems.SelectedItem.Text & "' workflow cannot be deleted." & vbCr & "It is used in the Self-service Intranet module.", _
        vbExclamation + vbOKOnly, Me.Caption
      fOK = False
    End If
    rsLocalInfo.Close
    Set rsLocalInfo = Nothing
  End If

  If fOK Then
    ' NPG20120222 Fault HRPRO-2027
    ' Check if the workflow is used in Mobile Designer.
    sSQL = "SELECT COUNT(*) AS recCount" & _
      " FROM tmpmobilegroupworkflows" & _
      " WHERE WorkflowID = " & CStr(WorkflowID)
    Set rsLocalInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    If rsLocalInfo!reccount > 0 Then
      MsgBox "The '" & lstItems.SelectedItem.Text & "' workflow cannot be deleted." & vbCr & "It is used in the Mobile Workflow module.", _
        vbExclamation + vbOKOnly, Me.Caption
      fOK = False
    End If
    rsLocalInfo.Close
    Set rsLocalInfo = Nothing
  End If
  
  
  If fOK Then
    ' Check if the workflow is used in a table's triggered link.
    sSQL = "SELECT COUNT(*) AS recCount" & _
      " FROM tmpWorkflowTriggeredLinks, tmpTables" & _
      " WHERE tmpWorkflowTriggeredLinks.deleted = FALSE" & _
      " AND tmpWorkflowTriggeredLinks.workflowID = " & CStr(WorkflowID)
    Set rsLocalInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
    If rsLocalInfo!reccount > 0 Then
      MsgBox "The '" & lstItems.SelectedItem.Text & "' workflow cannot be deleted." & vbCr & "It is used by a Workflow Link.", _
        vbExclamation + vbOKOnly, Me.Caption
      fOK = False
    End If
    rsLocalInfo.Close
    Set rsLocalInfo = Nothing
  End If

  If fOK Then
    If miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
      sMsg = "The '" & lstItems.SelectedItem.Text & "' workflow may be referenced externally." & vbNewLine & "Are you sure you want to delete it?"
    Else
      sMsg = "Delete workflow '" & lstItems.SelectedItem.Text & "', are you sure?"
    End If
    
    If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then

      daoDb.Execute "DELETE FROM tmpWorkflowElementItemValues WHERE itemID IN (SELECT ID FROM tmpWorkflowElementItems WHERE elementID IN (SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & CStr(WorkflowID) & "))"
      daoDb.Execute "DELETE FROM tmpWorkflowElementItems WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & CStr(WorkflowID) & ")"
      daoDb.Execute "DELETE FROM tmpWorkflowElementColumns WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & CStr(WorkflowID) & ")"
      daoDb.Execute "DELETE FROM tmpWorkflowElementValidations WHERE elementID IN(SELECT ID FROM tmpWorkflowElements WHERE workflowID=" & CStr(WorkflowID) & ")"
      daoDb.Execute "DELETE FROM tmpWorkflowElements WHERE workflowID=" & CStr(WorkflowID)
      daoDb.Execute "DELETE FROM tmpWorkflowLinks WHERE workflowID=" & CStr(WorkflowID)
      daoDb.Execute "UPDATE tmpWorkflows SET deleted = true WHERE ID =" & CStr(WorkflowID)

      ' Delete any expressions based on the Workflow.
      sSQL = "SELECT tmpExpressions.exprID" & _
        " FROM tmpExpressions" & _
        " WHERE tmpExpressions.deleted = FALSE" & _
        " AND (tmpExpressions.type = " & CStr(giEXPR_WORKFLOWCALCULATION) & _
        "   OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWSTATICFILTER) & _
        "   OR tmpExpressions.type = " & CStr(giEXPR_WORKFLOWRUNTIMEFILTER) & ")" & _
        " AND tmpExpressions.utilityID = " & CStr(WorkflowID)
      
      Set rsLocalInfo = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
      
      With rsLocalInfo
        Do While (Not .EOF) And fOK
        
          Set objExpression = New CExpression
          objExpression.ExpressionID = .Fields("exprID")
          fOK = objExpression.DeleteExpression(False)
          Set objExpression = Nothing
          
          .MoveNext
        Loop
      
        .Close
      End With
      Set rsLocalInfo = Nothing
      
      If fOK Then
        RefreshWorkflows
      End If
      
      Application.Changed = True
    End If
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub cmdExport_Click()

  Dim fOK As Boolean
  Dim iWorkflowId As Integer
  Dim objTestToLive As New OpenHRTestToLive.Repository
  Dim sXML As String
  Dim sOutputFileName As String
  
  iWorkflowId = WorkflowID
  fOK = (iWorkflowId > 0)

  If fOK Then
    Dim WorkflowName As String
    WorkflowName = "Exported Workflow_" + lstItems.SelectedItem + ".xml"
  
    sOutputFileName = WorkflowName
  
    With CommonDialog1
      .FileName = sOutputFileName
      .CancelError = False
      .DialogTitle = "Select a filename for your XML workflow export..."
      .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNOverwritePrompt
      .Filter = "XML File (*.xml)|*.xml"
      .DefaultExt = ".xml"
      
      .ShowSave
        
      If .FileName <> vbNullString Then
        If Len(.FileName) > 255 Then
          MsgBox "Path and file name must not exceed 255 characters in length", vbExclamation, Me.Caption
          fOK = False
        Else
          sOutputFileName = .FileName
        End If
      End If
      
    End With
  End If

  If fOK And Len(sOutputFileName) > 0 Then
    objTestToLive.Connection gsUserName, gsPassword, gsDatabaseName, gsServerName
    sXML = objTestToLive.ExportDefinition(iWorkflowId, sOutputFileName)
  End If

End Sub

Private Sub cmdModify_Click()
  ' Open the selected workflow.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim frmWFDes As frmWorkflowDesigner

  fOK = (WorkflowID > 0)

  If fOK Then
    ' Check if the workflow is in use.
    If Not mblnReadOnly Then
      If WorkflowsWithStatus(WorkflowID, giWFSTATUS_INPROGRESS) Then
        
        If MsgBox("The elements of the '" & lstItems.SelectedItem.Text & "' workflow cannot be modified." & vbCr & _
          "There are instances of this workflow in progress." & vbCrLf & vbCrLf & "Do you wish to continue?", _
          vbQuestion + vbYesNo, Me.Caption) = vbNo Then
          
          Exit Sub
        End If
      End If
     End If
     
    'Define a progress bar
    With gobjProgress
      .Caption = "Workflow Designer"
      .NumberOfBars = 1
      .Bar1Value = 1
      .Bar1MaxValue = 3
      .Bar1Caption = "Opening Workflow..."
      ' NPG Fault 13329
      '.AviFile = ""
      .AVI = dbWorkflow
      .MainCaption = "Workflow"
      .Cancel = False
      .Time = False
      .OpenProgress
    End With

    ' Display the Workflow designer form for the selected Workflow.
    Set frmWFDes = GetWorkflowForm
    If frmWFDes Is Nothing Then
      Set frmWFDes = New frmWorkflowDesigner
      With frmWFDes
        .IsNew = False
        .WorkflowID = WorkflowID

        'Update the progress bar
        gobjProgress.UpdateProgress

      End With
    End If

    Me.Hide

    frmWFDes.Show
    
    'Update the progress bar
    gobjProgress.UpdateProgress

    DoEvents
    frmWFDes.IsChanged = False
    frmWFDes.Start
  End If

TidyUpAndExit:
  'Close the progress bar
  gobjProgress.CloseProgress
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdNew_Click()
  ' Create a new screen.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim lngWorkflowID As Long
  Dim frmWorkflowDes As frmWorkflowDesigner

  ' Display the workflow properties form.
  With frmWorkflowEdit
    .WorkflowID = 0
    .ExternalInitiationQueryString = ""
    .InitiationType = miInitiationType
    Set .CallingForm = Me
    .Show vbModal
    fOK = Not .Cancelled
    lngWorkflowID = .WorkflowID
  End With
  Set frmWorkflowEdit = Nothing

  'Define a progress bar
  With gobjProgress
    .Caption = "Workflow Designer"
    .AVI = dbWorkflow
    .NumberOfBars = 1
    .Bar1Value = 1
    .Bar1MaxValue = 2
    .Bar1Caption = "Creating Workflow..."
    .Cancel = False
    .Time = False
    .OpenProgress
  End With

  ' If the workflow properties were confirmed then display the Workflow designer form.
  If fOK Then
    If lngWorkflowID > 0 Then
      With recWorkflowEdit
        .Index = "idxWorkflowID"
        .Seek "=", lngWorkflowID
      
        If Not .NoMatch Then
          .Edit
          
          !InitiationType = miInitiationType
          !BaseTable = mlngBaseTableID
          
          .Update
        End If
      End With

      Set frmWorkflowDes = New frmWorkflowDesigner
      With frmWorkflowDes
        .IsNew = True
        .WorkflowID = lngWorkflowID
        .Show
      End With

      'Update the progress bar
      gobjProgress.UpdateProgress
    
      frmWorkflowDes.Start
    End If

    UnLoad Me
  End If

TidyUpAndExit:
  'Close the progress bar
  gobjProgress.CloseProgress

  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdOK_Click()
  UnLoad Me
  
End Sub

Private Sub cmdPrint_Click()
  ' Print the selected workflow.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim frmWFDes As frmWorkflowDesigner

  fOK = (WorkflowID > 0)

  If fOK Then
    ' Instantiate (but hide) the Workflow designer form for the selected Workflow.
    Set frmWFDes = GetWorkflowForm
    If frmWFDes Is Nothing Then
      Set frmWFDes = New frmWorkflowDesigner
      With frmWFDes
        .IsNew = False
        .WorkflowID = WorkflowID
      End With
    End If

    frmWFDes.PrintWorkflow

    UnLoad frmWFDes

    DoEvents
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub cmdProperties_Click()
  ' Edit the selected workflow.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim frmWFDes As frmWorkflowDesigner

  If (WorkflowID > 0) Then
    ' Display the workflow properties form.
    With frmWorkflowEdit
      .WorkflowID = WorkflowID
      .InitiationType = miInitiationType
      Set .CallingForm = Me
      .WorkflowEnabled = mblnWorkflowEnabled
      
      .Show vbModal
      fOK = Not .Cancelled
    End With
    Set frmWorkflowEdit = Nothing

    If fOK Then
      ' The workflow name may have been changed so refresh the
      ' workflow list, and in the frmWorkflowDesigner caption if it is open.
      RefreshWorkflows

      ' The workflow name may have been changed in the properties window,
      ' so update the caption of the frmWorkflowDesigner form that displays this
      ' workflow, if it is open.
      Set frmWFDes = GetWorkflowForm
      If Not frmWFDes Is Nothing Then
        frmWFDes.Caption = "Workflow Designer - " & lstItems.SelectedItem.Text
      End If
    End If

    Me.SetFocus
    
    SelectWorkflow
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Resume TidyUpAndExit

End Sub

Private Sub Form_Activate()
  gfLoading = True
  
  RefreshWorkflows
  SelectWorkflow
  
  gfLoading = False

End Sub


Private Sub cboInitiationType_refresh()
  ' Initialise the InitiationType combo
  Dim iListIndex As Integer
  
  iListIndex = 1
  
  ' Clear the combo, and add the required items.
  With cboInitiationType
    .Clear
    
    .AddItem WorkflowInitiationTypeDescription(WORKFLOWINITIATIONTYPE_EXTERNAL)
    .ItemData(.NewIndex) = WORKFLOWINITIATIONTYPE_EXTERNAL
    If miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
      iListIndex = .NewIndex
    End If
    
    .AddItem WorkflowInitiationTypeDescription(WORKFLOWINITIATIONTYPE_MANUAL)
    .ItemData(.NewIndex) = WORKFLOWINITIATIONTYPE_MANUAL
    If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
      iListIndex = .NewIndex
    End If
    
    .AddItem WorkflowInitiationTypeDescription(WORKFLOWINITIATIONTYPE_TRIGGERED)
    .ItemData(.NewIndex) = WORKFLOWINITIATIONTYPE_TRIGGERED
    If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
      iListIndex = .NewIndex
    End If
    
    .ListIndex = iListIndex
  End With

End Sub

Private Sub cboBaseTable_refresh()
  ' Initialise the Base Table combo
  Dim iBaseTableListIndex As Integer
  
  iBaseTableListIndex = 0
  
  If mlngBaseTableID = 0 Then
    mlngBaseTableID = mlngPersModulePersonnelTableID
  End If
  
  ' Clear the combo, and add '<None>' items.
  With cboBaseTable
    .Clear
    
    If miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL Then
      .AddItem "<None>"
      .ItemData(.NewIndex) = mlngPersModulePersonnelTableID
    End If
    
    If miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
      .AddItem "<None>"
      .ItemData(.NewIndex) = 0
    End If
  End With

  If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
    ' Add items to the combo for each table that has not been deleted.
    With recTabEdit
      .Index = "idxName"
      If Not (.BOF And .EOF) Then
        .MoveFirst
      End If

      Do While Not .EOF
        If Not !Deleted Then
          cboBaseTable.AddItem !TableName
          cboBaseTable.ItemData(cboBaseTable.NewIndex) = !TableID
          If !TableID = mlngBaseTableID Then
            iBaseTableListIndex = cboBaseTable.NewIndex
          End If
        End If
        .MoveNext
      Loop
    End With
  End If
  
  With cboBaseTable
    If .ListCount > 0 Then
      .ListIndex = iBaseTableListIndex
    End If
    
    .Enabled = (miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED) And (.ListCount > 0)
    .ForeColor = IIf(.Enabled, vbWindowText, vbApplicationWorkspace)
    .BackColor = IIf(.Enabled, vbWindowBackground, vbButtonFace)
  End With

End Sub





Public Function RefreshWorkflows() As Boolean
  ' Populate the listbox with workflows.
  On Error GoTo ErrorTrap

  Dim iSelectedWorkflow As Integer
  Dim lngWorkflowID As Long

  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsRecords As DAO.Recordset
  Dim objListItem As ListItem
  Dim lngMaxTextWidth As Long
  Dim lngTextWidth As Long
  
  fOK = True
  lstItems.ListItems.Clear
  lngMaxTextWidth = 0
  
  ' Redimension an array to hold auxilliaryinformation for the workflows.
  ' Column 1 = workflow ID
  ' Column 2 = description
  ' Column 3 = enabled
  ReDim mavWorkflowInfo(5, 0)
  
  lngWorkflowID = WorkflowID

  iSelectedWorkflow = 1

  ' Define the selection string which determines
  ' what objects are displayed on the selection form.
  sSQL = "SELECT name, ID, description, enabled, locked, changed" & _
    " FROM tmpWorkflows" & _
    " WHERE deleted = FALSE"
  
  If miInitiationType = WORKFLOWINITIATIONTYPE_TRIGGERED Then
    sSQL = sSQL & _
    "   AND baseTable = " & CStr(mlngBaseTableID) & _
    "   AND initiationType = " & CStr(miInitiationType)
  ElseIf miInitiationType = WORKFLOWINITIATIONTYPE_EXTERNAL Then
    sSQL = sSQL & _
    "   AND initiationType = " & CStr(miInitiationType)
  Else
    sSQL = sSQL & _
    "   AND (initiationType = " & CStr(miInitiationType) & _
    "     OR initiationType IS null)"
  End If
  
  sSQL = sSQL & _
    " ORDER BY name"
    
  ' Populate the listbox with the required records.
  Set rsRecords = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)
  With rsRecords
    Do While Not .EOF

      Set objListItem = lstItems.ListItems.Add(, , !Name)
      objListItem.Tag = !ID
      
      ReDim Preserve mavWorkflowInfo(5, UBound(mavWorkflowInfo, 2) + 1)
      mavWorkflowInfo(1, UBound(mavWorkflowInfo, 2)) = !ID
      mavWorkflowInfo(2, UBound(mavWorkflowInfo, 2)) = !Description
      mavWorkflowInfo(3, UBound(mavWorkflowInfo, 2)) = !Enabled
      mavWorkflowInfo(4, UBound(mavWorkflowInfo, 2)) = !Locked
      mavWorkflowInfo(5, UBound(mavWorkflowInfo, 2)) = !Changed
      
      If !ID = lngWorkflowID Then
        iSelectedWorkflow = lstItems.ListItems.Count
      End If
      
      lngTextWidth = TextWidth(!Name)
      If lngMaxTextWidth < lngTextWidth Then
        lngMaxTextWidth = lngTextWidth
      End If
      
      .MoveNext
    Loop

    .Close
  End With
  Set rsRecords = Nothing

  If lstItems.ListItems.Count > 0 Then
    Set lstItems.SelectedItem = lstItems.ListItems(iSelectedWorkflow)
  End If
    
  lstItems.ColumnHeaders(1).Width = lngMaxTextWidth

  ' Update the status bar.
  sbScrOpen.Panels(1).Text = lstItems.ListItems.Count & " workflow" & _
    IIf(lstItems.ListItems.Count = 1, vbNullString, "s")

  RefreshControls

TidyUpAndExit:
  RefreshWorkflows = fOK
  Exit Function

ErrorTrap:
  fOK = False
  MsgBox Err.Description, vbExclamation + vbOKOnly, App.ProductName
  Err = False
  Resume TidyUpAndExit

End Function



Public Property Get WorkflowID() As Long
  ' Return the ID of the selected workflow.
  If Not lstItems.SelectedItem Is Nothing Then
    WorkflowID = CLng(lstItems.SelectedItem.Tag)
  Else
    WorkflowID = 0
  End If
  
End Property

Public Property Let WorkflowID(plngWorkflowID As Long)
  ' Set the ID of the selected workflow.
  mlngSelectedWorkflowID = plngWorkflowID
  
End Property
Private Sub RefreshControls()
  Dim fSelectionMade As Boolean
  Dim bLocked As Boolean
  Dim iLoop As Integer
  Dim sDescription As String
  Dim bModified As Boolean
  
  fSelectionMade = (Not (lstItems.SelectedItem Is Nothing))
 
  ' Refresh the 'description' textbox.
  sDescription = ""
  If fSelectionMade Then
    For iLoop = 1 To UBound(mavWorkflowInfo, 2)
      If CLng(mavWorkflowInfo(1, iLoop)) = CLng(lstItems.SelectedItem.Tag) Then
        sDescription = CStr(mavWorkflowInfo(2, iLoop))
        bLocked = mavWorkflowInfo(4, iLoop)
        mblnWorkflowEnabled = CBool(mavWorkflowInfo(3, iLoop))
        bModified = CBool(mavWorkflowInfo(5, iLoop))
        Exit For
      End If
    Next
  End If
  
  ' Enable/disable controls depending on the state of other.
  cmdNew.Enabled = Not mblnReadOnly
  cmdModify.Caption = IIf(bLocked, "&View...", "&Edit...")
  cmdModify.Enabled = fSelectionMade
  cmdCopy.Enabled = fSelectionMade And Not mblnReadOnly
  cmdDelete.Enabled = fSelectionMade And Not mblnReadOnly
  cmdProperties.Enabled = fSelectionMade
  cmdPrint.Enabled = fSelectionMade
  cmdExport.Enabled = fSelectionMade And Not bModified
  
  txtDesc.Text = sDescription

  ' Refresh the menu
  frmSysMgr.RefreshMenu
  
End Sub






Private Sub Form_Deactivate()
  
  ' Refresh the menu bar.
  frmSysMgr.RefreshMenu

End Sub


Private Sub Form_Initialize()
  miInitiationType = WORKFLOWINITIATIONTYPE_MANUAL
  mlngBaseTableID = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Pass any keystrokes onto the toolbar in the frmSysMgr form.
  'frmSysMgr.ActiveBar1.OnKeyDown KeyCode, Shift

  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
  End Select

  bHandled = frmSysMgr.tbMain.OnKeyDown(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  'TM20020102 Fault 2879
  Dim bHandled As Boolean
  
  bHandled = frmSysMgr.tbMain.OnKeyUp(KeyCode, Shift)
  If bHandled Then
    KeyCode = 0
    Shift = 0
  End If

End Sub


Private Sub Form_Load()
  
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  mblnReadOnly = (Application.AccessMode <> accFull And _
                  Application.AccessMode <> accSupportMode)
  
  If mblnReadOnly Then
    cmdNew.Enabled = False
    cmdModify.Caption = "&View"
    cmdCopy.Enabled = False
    cmdDelete.Enabled = False
    cmdProperties.Enabled = False
    cmdPrint.Enabled = False
    cmdExport.Enabled = False
  End If
  
  'JPD 20070615 Fault 12293
  mlngPersModulePersonnelTableID = GetModuleSetting(gsMODULEKEY_WORKFLOW, gsPARAMETERKEY_PERSONNELTABLE, 0)
  
  cboInitiationType_refresh
  
End Sub

Private Sub Form_Resize()
  
  On Error GoTo ErrorTrap
  
  Const XGAP = 150
  Const XGAP_RIGHT = 250
  
  Const YGAP = 100
  Const YGAP_BOTTOM = 650
  
  'JPD 20070927 Fault 12501
  DisplayApplication
  
  With fraDetails
    .Width = Me.Width - XGAP_RIGHT - cmdNew.Width - XGAP - .Left
    
    cboInitiationType.Width = .Width - cboInitiationType.Left - XGAP
    cboBaseTable.Width = cboInitiationType.Width
    
    lstItems.Width = .Width
    txtDesc.Width = .Width
    
    cmdNew.Left = .Left + .Width + (XGAP / 2)
    cmdModify.Left = cmdNew.Left
    cmdCopy.Left = cmdNew.Left
    cmdDelete.Left = cmdNew.Left
    cmdProperties.Left = cmdNew.Left
    cmdPrint.Left = cmdNew.Left
    cmdExport.Left = cmdNew.Left
    cmdOK.Left = cmdNew.Left
  End With
  
  With lstItems
    .Height = Me.Height - YGAP_BOTTOM - YGAP - sbScrOpen.Height - txtDesc.Height - YGAP - .Top
    txtDesc.Top = .Top + .Height + YGAP
  End With
    
  cmdOK.Top = Me.Height - YGAP_BOTTOM - sbScrOpen.Height - YGAP - cmdOK.Height
    
  ' Get rid of the icon off the form
  RemoveIcon Me
    
TidyUpAndExit:
  Exit Sub

ErrorTrap:
  Exit Sub
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'  Dim sAppName As String
'  Dim sSection As String
'
'  ' Save form size and position to the registry.
'  With Me
'
'    sAppName = App.ProductName
'    sSection = .Name
'
'    If .WindowState = vbNormal Then
'      SavePCSetting sSection, "Top", .Top
'      SavePCSetting sSection, "Left", .Left
'      SavePCSetting sSection, "Height", .Height
'      SavePCSetting sSection, "Width", .Width
'    End If
'
'    SavePCSetting Me.Name, "State", Me.WindowState
'  End With
  
  ' Update the menu.
  frmSysMgr.RefreshMenu True
  
End Sub


Private Sub lstItems_DblClick()
  If cmdModify.Enabled Then
    cmdModify_Click
  End If

End Sub


Private Sub lstItems_ItemClick(ByVal Item As ComctlLib.ListItem)
  ' Ensure that only the required command controls are enabled.
  If lstItems.SelectedItem Is Nothing Then
    mlngSelectedWorkflowID = 0
  Else
    mlngSelectedWorkflowID = lstItems.SelectedItem.Tag
  End If
  
  RefreshControls

End Sub

