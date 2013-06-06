VERSION 5.00
Begin VB.Form frmViewProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Properties"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1038
   Icon            =   "frmViewProp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraScreens 
      Caption         =   "Screens :"
      Height          =   2100
      Left            =   200
      TabIndex        =   10
      Top             =   2895
      Width           =   5310
      Begin VB.ListBox lstScreens 
         Height          =   1635
         Left            =   200
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   300
         Width           =   4920
      End
   End
   Begin VB.TextBox txtWhereClause 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   2400
      Width           =   3685
   End
   Begin VB.CommandButton cmdWhereClause 
Caption = "..."
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3000
      TabIndex        =   5
      Top             =   5115
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   4275
      TabIndex        =   6
      Top             =   5115
      Width           =   1200
   End
   Begin VB.TextBox txtViewName 
      Height          =   315
      Left            =   1470
      MaxLength       =   128
      TabIndex        =   0
      Top             =   200
      Width           =   4000
   End
   Begin VB.TextBox txtViewDescription 
      Height          =   1500
      Left            =   1470
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   700
      Width           =   4000
   End
   Begin VB.Label lblWhereClause 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter :"
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   2460
      Width           =   825
   End
   Begin VB.Label lblViewName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   255
      Width           =   870
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description :"
      Height          =   195
      Left            =   195
      TabIndex        =   7
      Top             =   765
      Width           =   1260
   End
End
Attribute VB_Name = "frmViewProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private glngViewID As Long
Private glngOriginalViewID As Long
'NPG20080206 Fault 12874
Private mstrCopyFromViewName As String
Private gLngTableID As Long
Private gLngExprID As Long
Private gblnCopy As Long
Private gfCancelled As Boolean
Private gfRefreshing As Boolean
' NPG20081114 Fault 13334
Private maScreenPrevChk() As Boolean

Private mblnReadOnly As Boolean

Private Function SaveChanges() As Boolean
  ' Save the changes.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim bNew As Boolean
  Dim frmPermissions As frmDefaultPermissions2
  
  fOK = True
  
  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans
  
  'NPG20080207 Fault 12874
  Set frmPermissions = New frmDefaultPermissions2
  
    ' Save the View details to the local database.
  With recViewEdit
    ' Get a unique number
    If glngViewID = 0 Or gblnCopy Then
      .AddNew
    
      'Apply permissions to existing security groups
      If Trim(mstrCopyFromViewName) = vbNullString Then
        frmPermissions.SetType "new", giVIEW, Me.Icon
      Else
        frmPermissions.SetType "copy", giVIEW, Me.Icon
      End If
      frmPermissions.Show vbModal
      If frmPermissions.OkCancel = vbOK Then
        'NPG20080206 Fault 12874
        If frmPermissions.CopyPermissions Then
          .Fields("OriginalViewName") = mstrCopyFromViewName
        Else
          .Fields("GrantRead") = frmPermissions.GrantRead
          .Fields("GrantEdit") = frmPermissions.GrantEdit
          .Fields("GrantNew") = frmPermissions.GrantNew
          .Fields("GrantDelete") = frmPermissions.GrantDelete
        End If

        glngViewID = Database.UniqueColumnValue("tmpViews", "ViewID")
        !ViewID = glngViewID
        .Fields("New") = True
        .Fields("Deleted") = False
        .Fields("Changed") = False
        'NPG20080206 Fault 12874
        '.Fields("OriginalViewName") = Trim(txtViewName.Text)
      
      Else
        fOK = False
      End If
    
    Else
      .Index = "idxViewID"
      .Seek "=", glngViewID
      fOK = Not .NoMatch
      bNew = False
      
      If fOK Then
        .Edit
        .Fields("Changed") = True
      End If
    End If
   
    If fOK Then
      
      'MH20020808 Set flag if view name has changed and
      'remove reference to "viewAlternativeName"
      '.Fields("ViewAlternativeName") = Trim(txtViewName.Text)
      If .Fields("ViewName") <> Trim(txtViewName.Text) Or IsNull(.Fields("ViewName")) Then
        Application.ChangedViewName = True
        .Fields("ViewName") = Trim(txtViewName.Text)
      End If
      
      .Fields("ViewDescription") = Trim(txtViewDescription.Text)
      .Fields("ViewTableID") = gLngTableID
      .Fields("ViewSQL") = "" ' No longer used.
      .Fields("ExpressionID") = gLngExprID
      .Update
    End If
  End With
  
  If fOK Then
    ' Update the recViewScreens Table.
    ' Skip the all column.
    For iCount = 1 To lstScreens.ListCount - 1
      With recViewScreens
        .Index = "idxViewScreen"
        
        If gblnCopy Then
          .Seek "=", glngOriginalViewID, lstScreens.ItemData(iCount)
        Else
          .Seek "=", glngViewID, lstScreens.ItemData(iCount)
        End If
        
        ' See if we have found it.
        ' If we have then decide whether we are to delete it or just reset the flags.
        If Not .NoMatch Then
          If lstScreens.Selected(iCount) Then
            ' See if it is marked for deletion and if it is remove the flag.
            If gblnCopy Then
              .AddNew
              .Fields("ScreenID") = lstScreens.ItemData(iCount)
              .Fields("ViewID") = glngViewID
              .Fields("New") = True
              .Fields("Deleted") = False
              .Update
            ElseIf .Fields("Deleted") Then
              .Edit
              .Fields("Deleted") = False
              .Update
            End If
          Else
            ' They have removed this view from being available to this screen
            ' See if it has a new flag and if it does then delete the row
            ' otherwise set the delete flag
            If .Fields("New") Then
              .Delete
            Else
              .Edit
              .Fields("Deleted") = True
              .Update
            End If
          End If
        Else
          ' The entry does not exist in the tmpViewScreens table so add it
          ' if we are granting the screen the view otherwise do nothing
          If lstScreens.Selected(iCount) Then
            .AddNew
            .Fields("ScreenID") = lstScreens.ItemData(iCount)
            .Fields("ViewID") = glngViewID
            .Fields("New") = True
            .Fields("Deleted") = False
            .Update
          End If
        End If
      End With
    Next iCount
  End If
  
TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  
  'NPG20080207 Fault 12874
  Set frmPermissions = Nothing
  
  SaveChanges = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Property Get Cancelled() As Boolean
  ' Return the 'cancelled' property.
  Cancelled = gfCancelled
  
End Property

Public Property Let ViewID(plngViewID As Long)
  ' Set the 'view ID' property, and read the view details from the local database.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sViewName As String
  Dim tmpViewName As String
  Dim sCaption As String
  Dim sDescription As String
  
  glngViewID = plngViewID
  glngOriginalViewID = plngViewID

  ' Find the record in the table.
  With recViewEdit
    .Index = "idxViewID"
    .Seek "=", glngViewID
    
    fOK = Not .NoMatch
    
    If fOK Then
      ' Get the existing details from the database.
      sViewName = Trim(.Fields("ViewName"))
      'NPG20080206 Fault 12874
      mstrCopyFromViewName = IIf(IsNull(.Fields("OriginalViewName")), "", .Fields("OriginalViewName"))

      If gblnCopy Then
        Dim varBookmark As Variant
        Dim iCounter As Integer, fGoodName As Boolean
        
        varBookmark = recViewEdit.Bookmark
       
        ' Create a new table name.
        tmpViewName = "Copy_of_" & sViewName
        
        ' Check that the view name is not already used.
        iCounter = 1
        fGoodName = False
        Do While Not fGoodName
            recTabEdit.Index = "idxName"
            recTabEdit.Seek "=", tmpViewName, False
            
            If Not recTabEdit.NoMatch Then
              iCounter = iCounter + 1
              tmpViewName = "Copy_" & Trim(Str(iCounter)) & "_of_" & sViewName
            Else
              ' Tablename is ok, now check views
              recViewEdit.Index = "idxViewName"
              recViewEdit.Seek "=", tmpViewName, False
              
              If Not recViewEdit.NoMatch Then
                iCounter = iCounter + 1
                tmpViewName = "Copy_" & Trim(Str(iCounter)) & "_of_" & sViewName
              Else
                fGoodName = True
                sViewName = tmpViewName
              End If
            End If
        Loop
        recViewEdit.Bookmark = varBookmark
      End If
      
      sDescription = Trim(.Fields("ViewDescription"))
      gLngExprID = IIf(IsNull(.Fields("expressionID")), 0, .Fields("expressionID"))
      gLngTableID = .Fields("viewTableID")
      
      ' Enable the view name.
'      txtViewName.Enabled = False
      txtViewName.Enabled = Not mblnReadOnly
      sCaption = "'" & sViewName & "' View Properties"
    Else
      ' New View.
      sViewName = ""
      sDescription = ""
      gLngExprID = 0

      ' Disable the view name.
      'txtViewName.Enabled = True
      txtViewName.Enabled = Not mblnReadOnly

      sCaption = "New View Properties"
    End If
  End With
  
  Me.Caption = sCaption
   
  ' Load the text boxes with the appropriate info
  txtViewName.Text = sViewName
  txtViewDescription.Text = sDescription
    
  ' Get the details of the expression.
  GetWhereClauseDetails
  
  ' Load the screen listbox.
  lstScreens_Refresh
    
TidyUpAndExit:
  Exit Property
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Property
Private Function GetWhereClauseDetails() As Boolean
  ' Get the 'where clause' expression details.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sExprName As String
  Dim objExpr As CExpression
  
  fOK = True
  
  ' Initialise the default values.
  sExprName = ""
    
  ' Instantiate the expression class.
  Set objExpr = New CExpression
    
  With objExpr
    ' Set the expression id.
    .ExpressionID = gLngExprID
      
    ' Read the required info from the expression.
    If .ReadExpressionDetails Then
      sExprName = .Name
    End If
  End With
  
TidyUpAndExit:
  ' Disassociate object variables.
  Set objExpr = Nothing
  If Not fOK Then
    sExprName = ""
  End If
  ' Update the clause controls properties.
  txtWhereClause.Text = sExprName
  GetWhereClauseDetails = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Property Get ViewID() As Long
  ' Return the 'view ID' property.
  ViewID = glngViewID

End Property

Public Property Let TableID(plngTableID As Long)
  ' Set the 'tableID' property.
  gLngTableID = plngTableID
  
End Property

Public Property Let Copy(pblnCopy As Boolean)
  gblnCopy = pblnCopy
End Property

Private Sub cmdCancel_Click()
  ' Flag that the changes have been cancelled..
  gfCancelled = True
  
  ' Unload the form.
  UnLoad Me
  
End Sub

Private Sub cmdOK_Click()
  ' Validate and save the View properties.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim sViewName As String
  
  sViewName = Trim(txtViewName.Text)
  
  ' Check that we have a view name.
  fOK = (Len(sViewName) > 0)
  If Not fOK Then
    MsgBox "A view name must be entered.", _
      vbOKOnly + vbExclamation, Application.Name
    If txtViewName.Enabled Then
      txtViewName.SetFocus
    End If
  End If
  
  
  ' Check that we have a unique view name.
  If fOK Then
    With recViewEdit
      If Not (.BOF And .EOF) Then
        ' Seek the views table for a view with this name.
        .Index = "idxViewName"
        .Seek "=", sViewName, False
          
        If Not .NoMatch Then
          fOK = (.Fields("viewID") = glngViewID) And Not gblnCopy
          
          If Not fOK Then
            ' Flag to the user if there already exists a view with this name.
            MsgBox "A view named '" & sViewName & "' already exists!", _
              vbOKOnly + vbExclamation, Application.Name
            txtViewName.SetFocus
          End If
        End If
      End If
    End With
  End If
    
  If fOK Then
    With recTabEdit
      .Index = "idxName"
      .Seek "=", sViewName, False
      
      fOK = .NoMatch
      
      If Not fOK Then
        MsgBox "There is a table named '" & sViewName & "' and therefore cannot be used as the view name.", _
          vbOKOnly + vbExclamation, Application.Name
        txtViewName.SetFocus
      End If
    End With
  End If

  If fOK Then
    ' Ensure that the table name is not a keyword.
    fOK = Not Database.IsKeyword(sViewName)
    If Not fOK Then
      MsgBox "'" & sViewName & "' cannot be used as a view name" & _
        vbCr & "as it is a reserved word.", _
        vbOKOnly + vbExclamation, Application.Name
      txtViewName.SetFocus
    End If
  End If
  
  If fOK Then
    ' Ensure that the table name is not a system database name.
    fOK = UCase(Left(sViewName, 6)) <> "ASRSYS"
    
    If Not fOK Then
      MsgBox "'" & sViewName & "' cannot be used as a view name" & _
        vbCr & "as the prefix 'ASRSys' is reserved for system views.", _
      vbOKOnly + vbExclamation, Application.Name
      txtViewName.SetFocus
    End If
  End If
 
  ' Save the changes.
  If fOK Then
    fOK = SaveChanges
  End If
  
  ' Unload the form only if everything was okay.
  If fOK Then
    gfCancelled = False
    UnLoad Me
  End If

TidyUpAndExit:
  Exit Sub

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub

Private Sub cmdWhereClause_Click()
  ' Display the 'Where Clause' expression selection form.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim objExpr As CExpression

  fOK = True
  
  ' Instantiate an expression object.
  Set objExpr = New CExpression
  
  With objExpr
    ' Set the properties of the expression object.
    .Initialise gLngTableID, gLngExprID, giEXPR_VIEWFILTER, giEXPRVALUE_LOGIC
    
    ' Instruct the expression object to display the
    ' expression selection form.
    If .SelectExpression Then
      gLngExprID = .ExpressionID
        
      ' Read the selected expression info.
      fOK = GetWhereClauseDetails
    Else
      ' Check in case the original expression has been deleted.
      With recExprEdit
        .Index = "idxExprID"
        .Seek "=", gLngExprID, False

        If .NoMatch Then
          ' Read the selected expression info.
          gLngExprID = 0
          fOK = GetWhereClauseDetails
        End If
      End With
    End If
  End With
  
TidyUpAndExit:
  Set objExpr = Nothing
  If Not fOK Then
    MsgBox "Error changing expression ID.", vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Sub


Private Sub Form_Initialize()

  'Allow access to view manager, even if they have limited
  'system access the view manager should be enabled.
  mblnReadOnly = (Application.AccessMode = accSystemReadOnly)

  If mblnReadOnly Then
    ControlsDisableAll Me
    cmdWhereClause.Enabled = True
  End If

End Sub

Private Sub Form_Load()
  ' Clear the menu shortcuts. This needs to be done so that some shortcut keys
  ' (eg. DEL) will function normally in textboxes instead of triggering menu options.
  frmSysMgr.ClearMenuShortcuts
  
  ' Position the form.
  UI.frmAtCenterOfParent Me, frmSysMgr
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub lstScreens_Click()
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fNewValue As Boolean
  Dim fAllSelected As Boolean
  Dim iCount As Integer
  Dim Item As Integer
  
  Item = lstScreens.ListIndex
  
  fOK = True
  
  If gfRefreshing Then
    Exit Sub
  End If

  gfRefreshing = True

  If mblnReadOnly Then
    With lstScreens
      gfRefreshing = True
      .Selected(.ListIndex) = Not .Selected(.ListIndex)
      gfRefreshing = False
    End With
    Exit Sub
  End If
  
  ' NPG
  If lstScreens.Selected(Item) = True Then
    If maScreenPrevChk(Item) Then
      lstScreens.Selected(Item) = False
    End If
  Else
    If Not maScreenPrevChk(Item) Then
      lstScreens.Selected(Item) = True
    End If
  End If
  
  fNewValue = lstScreens.Selected(Item)
  
  UI.LockWindow lstScreens.hWnd

  ' Check if the user has selected the all columns
  If Item = 0 Then
    For iCount = 1 To lstScreens.ListCount - 1
      ' Update all rows in the listbox.
      lstScreens.Selected(iCount) = fNewValue
      maScreenPrevChk(iCount) = fNewValue
    Next iCount
  Else
    ' Check for all columns now being selected or deselected
    fAllSelected = True
    For iCount = 1 To lstScreens.ListCount - 1
      If Not lstScreens.Selected(iCount) Then
        fAllSelected = False
        Exit For
      End If
    Next iCount
    ' Update the 'All' row in the listbox.
    lstScreens.Selected(0) = fAllSelected
    maScreenPrevChk(0) = fAllSelected
  End If

gfRefreshing = False

maScreenPrevChk(Item) = lstScreens.Selected(Item)

TidyUpAndExit:
  UI.UnlockWindow
  If Not fOK Then
    MsgBox "Error updating view." & vbCr & vbCr & _
      Err.Description, vbExclamation + vbOKOnly, App.ProductName
  End If
  Exit Sub
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Sub

Private Sub lstScreens_ItemCheck(Item As Integer)
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim fNewValue As Boolean
'  Dim fAllSelected As Boolean
'  Dim iCount As Integer
'
'  fOK = True
'
'  If gfRefreshing Then
'    Exit Sub
'  End If
'
'  If mblnReadOnly Then
'    With lstScreens
'      gfRefreshing = True
'      .Selected(.ListIndex) = Not .Selected(.ListIndex)
'      gfRefreshing = False
'    End With
'    Exit Sub
'  End If
'
'  fNewValue = lstScreens.Selected(Item)
'
'  UI.LockWindow lstScreens.hWnd
'
'  ' Check if the user has selected the all columns
'  If Item = 0 Then
'    For iCount = 1 To lstScreens.ListCount - 1
'      ' Update all rows in the listbox.
'      lstScreens.Selected(iCount) = fNewValue
'    Next iCount
'  Else
'    ' Check for all columns now being selected or deselected
'    fAllSelected = True
'    For iCount = 1 To lstScreens.ListCount - 1
'      If Not lstScreens.Selected(iCount) Then
'        fAllSelected = False
'        Exit For
'      End If
'    Next iCount
'    ' Update the 'All' row in the listbox.
'    lstScreens.Selected(0) = fAllSelected
'  End If
'
'TidyUpAndExit:
'  UI.UnlockWindow
'  If Not fOK Then
'    MsgBox "Error updating view." & vbCr & vbCr & _
'      Err.Description, vbExclamation + vbOKOnly, App.ProductName
'  End If
'  Exit Sub
'
'ErrorTrap:
'  fOK = False
'  Resume TidyUpAndExit

End Sub


Private Sub txtViewDescription_GotFocus()
  ' Select the whole string.
  
  UI.txtSelText
  
  cmdOK.Default = False
  
End Sub

Private Sub txtViewDescription_LostFocus()
  cmdOK.Default = True

End Sub


Private Sub txtViewName_Change()
  Dim sValidatedName As String
  Dim iSelStart As Integer
  Dim iSelLen As Integer
  
  'JPD 20090102 Fault 13484
  sValidatedName = Database.ValidateName(txtViewName.Text)
  
  If sValidatedName <> txtViewName.Text Then
    iSelStart = txtViewName.SelStart
    iSelLen = txtViewName.SelLength
    
    txtViewName.Text = sValidatedName
    
    txtViewName.SelStart = iSelStart
    txtViewName.SelLength = iSelLen
  End If

End Sub

Private Sub txtViewName_GotFocus()
  ' Select the whole string.
  UI.txtSelText

End Sub

Private Sub txtViewName_KeyPress(KeyAscii As Integer)
  ' Validate the character entered.
  KeyAscii = Database.ValidNameChar(KeyAscii, txtViewName.SelStart)
End Sub



Private Function lstScreens_Refresh() As Boolean
  ' Populate the listbox with the screens for the selected view.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fInView As Boolean
  Dim fAllScreens As Boolean
  Dim sSQL As String
  Dim rsScreens As dao.Recordset
  Dim iTmpCount As Integer

  gfRefreshing = True
  UI.LockWindow lstScreens.hWnd
  
  ' Remove all the screens from the listbox.
  lstScreens.Clear
  
  ReDim maScreenPrevChk(0)
  
  sSQL = "SELECT tmpScreens.*" & _
    " FROM tmpScreens" & _
    " WHERE tmpScreens.tableID = " & Trim(Str(gLngTableID)) & _
    " AND tmpScreens.deleted = FALSE" & _
    " AND (tmpScreens.QuickEntry = FALSE)" & _
    " AND (tmpScreens.SSIntranet = FALSE)" & _
    " ORDER BY name"
  Set rsScreens = daoDb.OpenRecordset(sSQL, dbOpenForwardOnly, dbReadOnly)

  ' Add the screens to the grid.
  fAllScreens = True
  With rsScreens
  iTmpCount = 1
  
    While Not .EOF
      
      ReDim Preserve maScreenPrevChk(iTmpCount)
      
      recViewScreens.Index = "idxViewScreen"
      recViewScreens.Seek "=", glngViewID, .Fields("screenID")
      fInView = Not recViewScreens.NoMatch
      If fInView Then
        fInView = Not recViewScreens("deleted")
      End If
      lstScreens.AddItem .Fields("name")
      lstScreens.ItemData(lstScreens.NewIndex) = .Fields("screenID")
      lstScreens.Selected(lstScreens.NewIndex) = fInView
        
      ' NPG20081114 Fault 13334
      maScreenPrevChk(iTmpCount) = lstScreens.Selected(lstScreens.NewIndex)
      iTmpCount = iTmpCount + 1
      
      If Not fInView Then fAllScreens = False
        
      .MoveNext
    Wend
    .Close
  End With
  Set rsScreens = Nothing

  With lstScreens
    fAllScreens = fAllScreens And (.ListCount > 0)
    ' Add the 'all screens' column.
    .AddItem "<All>", 0
    .ItemData(.NewIndex) = 0
    .Selected(.NewIndex) = fAllScreens
    ' See if all the screens are all selected.
    .Enabled = (.ListCount > 1)
  
    ' Select the first item.
    If .Enabled Then
      .ListIndex = 0
    End If
  End With
  
  ' NPG
  maScreenPrevChk(0) = fAllScreens
  
TidyUpAndExit:
  UI.UnlockWindow
  gfRefreshing = False
  ' Disassociate object variables.
  Set rsScreens = Nothing
  lstScreens_Refresh = fOK
  Exit Function
  
ErrorTrap:
  ' indicate that the function has failed
  fOK = False
  MsgBox Err.Description, vbCritical + vbOKOnly, App.Title
  Resume TidyUpAndExit
  
End Function

