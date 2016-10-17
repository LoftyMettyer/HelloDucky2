VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "Codejock.Controls.v13.1.0.ocx"
Begin VB.Form frmDocumentMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Document Management Type"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocumentMap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6630
      Left            =   90
      TabIndex        =   19
      Top             =   45
      Width           =   10545
      _Version        =   851969
      _ExtentX        =   18600
      _ExtentY        =   11695
      _StockProps     =   68
      ItemCount       =   1
      Item(0).Caption =   "Definition"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "fraDefinition"
      Item(0).Control(1)=   "fraClassification"
      Item(0).Control(2)=   "fraDestination"
      Begin VB.Frame fraDestination 
         Caption         =   "Record Identification : "
         Height          =   2355
         Left            =   90
         TabIndex        =   28
         Top             =   4095
         Width           =   10335
         Begin VB.ComboBox cboTargetTitle 
            Height          =   315
            Left            =   7050
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1080
            Width           =   3090
         End
         Begin VB.ComboBox cboParent2Table 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2295
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.ComboBox cboParent2Keyfield 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   2295
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetCategory 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   675
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetType 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1080
            Width           =   3090
         End
         Begin VB.ComboBox cboParent1Keyfield 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1890
            Width           =   3090
         End
         Begin VB.ComboBox cboParent1Table 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1890
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetKeyField 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   270
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetColumn 
            Height          =   315
            Left            =   7065
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   675
            Width           =   3090
         End
         Begin VB.ComboBox cboTargetTable 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   270
            Width           =   3090
         End
         Begin VB.Label lblTitle 
            Caption         =   "Title :"
            Height          =   225
            Left            =   5310
            TabIndex        =   38
            Top             =   1125
            Width           =   1290
         End
         Begin VB.Label lblParent2Table 
            Caption         =   "Parent 2 Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   37
            Top             =   2340
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Label lblParent2Keyfield 
            Caption         =   "Parent 2 Key Field :"
            Height          =   375
            Left            =   5310
            TabIndex        =   36
            Top             =   2340
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblDestinationType 
            Caption         =   "Type : "
            Height          =   240
            Left            =   225
            TabIndex        =   35
            Top             =   1125
            Width           =   1095
         End
         Begin VB.Label lblDestinationCategory 
            Caption         =   "Category :"
            Height          =   285
            Left            =   225
            TabIndex        =   34
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label lblParent1Keyfield 
            Caption         =   "Parent Key Field :"
            Height          =   375
            Left            =   5310
            TabIndex        =   33
            Top             =   1935
            Width           =   1815
         End
         Begin VB.Label lblParent1Table 
            Caption         =   "Parent Table :"
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   1935
            Width           =   1725
         End
         Begin VB.Label lblTargetTable 
            Caption         =   "Table :"
            Height          =   285
            Left            =   225
            TabIndex        =   31
            Top             =   315
            Width           =   1365
         End
         Begin VB.Label lblTargetColumn 
            Caption         =   "Document URL : "
            Height          =   330
            Left            =   5310
            TabIndex        =   30
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblTargetKeyfield 
            Caption         =   "Key Field :"
            Height          =   240
            Left            =   5310
            TabIndex        =   29
            Top             =   315
            Width           =   1230
         End
      End
      Begin VB.Frame fraClassification 
         Caption         =   "Document Description : "
         Height          =   1410
         Left            =   90
         TabIndex        =   25
         Top             =   2520
         Width           =   10335
         Begin VB.ComboBox cboTypes 
            Height          =   315
            ItemData        =   "frmDocumentMap.frx":000C
            Left            =   1845
            List            =   "frmDocumentMap.frx":000E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   765
            Width           =   3090
         End
         Begin VB.ComboBox cboCategories 
            Height          =   315
            Left            =   1845
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   3090
         End
         Begin VB.Label Label2 
            Caption         =   "Type Value :"
            Height          =   285
            Left            =   270
            TabIndex        =   27
            Top             =   810
            Width           =   1185
         End
         Begin VB.Label Label1 
            Caption         =   "Category Value :"
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   26
            Top             =   420
            Width           =   1500
         End
      End
      Begin VB.Frame fraDefinition 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   90
         TabIndex        =   20
         Top             =   405
         Width           =   10335
         Begin VB.TextBox txtUserName 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7065
            MaxLength       =   30
            TabIndex        =   2
            Top             =   315
            Width           =   3045
         End
         Begin VB.OptionButton optReadWrite 
            Caption         =   "Read / &Write"
            Height          =   195
            Left            =   7065
            TabIndex        =   4
            Top             =   765
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.OptionButton optReadOnly 
            Caption         =   "&Read Only"
            Height          =   195
            Left            =   7065
            TabIndex        =   5
            Top             =   1155
            Width           =   1470
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1845
            MaxLength       =   50
            TabIndex        =   1
            Top             =   315
            Width           =   3090
         End
         Begin VB.TextBox txtDesc 
            Height          =   1080
            Left            =   1845
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   705
            Width           =   3090
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Owner :"
            Height          =   195
            Index           =   4
            Left            =   5280
            TabIndex        =   24
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name :"
            Height          =   195
            Index           =   2
            Left            =   315
            TabIndex        =   23
            Top             =   365
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description :"
            Height          =   195
            Index           =   1
            Left            =   315
            TabIndex        =   22
            Top             =   765
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Access :"
            Height          =   195
            Index           =   3
            Left            =   5280
            TabIndex        =   21
            Top             =   765
            Width           =   600
         End
      End
   End
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9450
      TabIndex        =   18
      Top             =   6750
      Width           =   1200
      _Version        =   851969
      _ExtentX        =   2117
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancel"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdOK 
      Height          =   375
      Left            =   8190
      TabIndex        =   0
      Top             =   6750
      Width           =   1200
      _Version        =   851969
      _ExtentX        =   2117
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&OK"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "frmDocumentMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngDocumentMapID As Long

Private mblnCancelled As Boolean
Private mblnReadOnly As Boolean
Private mblnFromCopy As Boolean
Private mblnDefinitionCreator As Boolean
Private mbChanged As Boolean

Private mlngLabelDefinitionID As Long
Private mlngDocumentDefinitionID As Long

Private mdatData As DataMgr.clsDataAccess
Private mclsGeneral As DataMgr.clsGeneral

Private Const SQLTableDef = "ASRSysDocumentManagementTypes"

Public Property Let Changed(mbChanged As Boolean)
  cmdOK.Enabled = mbChanged
End Property

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mblnCancelled = bCancel
End Property

Public Property Get SelectedID() As Long
  SelectedID = mlngDocumentMapID
End Property

Public Sub PrintDefinition(plngDocumentMapID)
  Dim objPrintDef As clsPrintDef
  Dim rsTemp As Recordset
  Dim strDocumentTypeName As String
  
  Set mdatData = New DataMgr.clsDataAccess
  Set rsTemp = GetDefinition
  
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation, "Document Definition"
    Exit Sub
  End If
   
  Set objPrintDef = New DataMgr.clsPrintDef

  If objPrintDef.IsOK Then
    With objPrintDef
      If .PrintStart(False) Then
        .PrintHeader strDocumentTypeName & " Document Type Definition: " & rsTemp!Name
        .PrintNormal "Name : " & rsTemp!Name
        .PrintNormal "Description : " & rsTemp!Description
        .PrintNormal
        Select Case rsTemp!Access
          Case "RW": .PrintNormal "Owner : " & rsTemp!userName & vbTab & "Access : Read / Write"
          Case "RO": .PrintNormal "Owner : " & rsTemp!userName & vbTab & "Access : Read only"
        End Select
        .PrintTitle "Document Description"
        .PrintNormal "Category Value : " & Me.cboCategories
        .PrintNormal "Type Value : " & Me.cboTypes
        .PrintTitle "Record Identification"
        .PrintNormal "Table : " & Me.cboTargetTable & vbTab & "Key Field : " & Me.cboTargetKeyField
        .PrintNormal "Category : " & Me.cboTargetCategory & vbTab & "Document URL : " & Me.cboTargetColumn
        .PrintNormal "Type : " & Me.cboTargetType
        .PrintNormal "Title : " & Me.cboTargetTitle
        .PrintNormal
        .PrintNormal "Parent Table : " & Me.cboParent1Table & vbTab & "Parent Key Field : " & Me.cboParent1Keyfield
        .PrintEnd
      End If
    End With
  End If
  
  Set mdatData = Nothing
Exit Sub

LocalErr:
  COAMsgBox "Printing Document Typed Definition Failed"
End Sub



Public Function Initialise(bNew As Boolean, bCopy As Boolean, Optional lngDocumentMapID As Long, Optional bPrint As Boolean) As Boolean

  Set mdatData = New DataMgr.clsDataAccess
  Set mclsGeneral = New DataMgr.clsGeneral
  Dim sAccess As String

  PopulateCategoriesCombo
  LoadTableCombo cboTargetTable

  If bNew Then
    mlngDocumentMapID = 0
    sAccess = GetUserSetting("utils&reports", "dfltaccess version1", ACCESS_READWRITE)
    Select Case sAccess
      Case ACCESS_READWRITE
        optReadWrite.Value = True
      Case Else
        optReadOnly.Value = True
    End Select
  Else
    mblnFromCopy = bCopy
    mlngDocumentMapID = lngDocumentMapID
    RetreiveDefinition
  End If
  
  mblnReadOnly = Not datGeneral.SystemPermission("VERSION1", "EDIT")
  If mblnReadOnly Then
    ControlsDisableAll Me, False
    EnableControl cmdOK, True
    EnableControl cmdCancel, True
  End If
  
  'Reset pointer so copy will be saved as new
  If mblnFromCopy Then
    mlngDocumentMapID = 0
    mbChanged = True
  Else
    mbChanged = False
  End If
  
  cmdOK.Enabled = mbChanged
  
End Function

Private Function GetDefinition() As ADODB.Recordset

  Dim strSQL As String

  strSQL = "SELECT * FROM " & SQLTableDef & " " & _
           "WHERE [DocumentMapID] = " & CStr(mlngDocumentMapID)
  Set GetDefinition = mdatData.OpenRecordset(strSQL, adOpenForwardOnly, adLockReadOnly)

End Function

Private Sub RetreiveDefinition()

  Dim bOK As Boolean
  Dim rsTemp As ADODB.Recordset

  Set rsTemp = GetDefinition
  If rsTemp.BOF And rsTemp.EOF Then
    COAMsgBox "This definition has been deleted by another user.", vbExclamation + vbOKOnly, "Label Definition"
    bOK = False
    Exit Sub
  End If

  txtDesc.Text = IIf(rsTemp!Description <> vbNullString, rsTemp!Description, vbNullString)
  
  If mblnFromCopy Then
    txtName.Text = "Copy of " & rsTemp!Name
    txtUserName = gsUserName
    mblnDefinitionCreator = True
  Else
    txtName.Text = rsTemp!Name
    txtUserName = StrConv(rsTemp!userName, vbProperCase)
    mblnDefinitionCreator = (LCase$(rsTemp!userName) = LCase$(gsUserName))
  End If

  mblnReadOnly = Not datGeneral.SystemPermission("VERSION1", "EDIT")

  ' Set the access type
  Select Case rsTemp!Access
  Case "RW"
    optReadWrite = True
    optReadWrite.Enabled = mblnDefinitionCreator
    optReadOnly.Enabled = mblnDefinitionCreator
  Case "RO"
    optReadOnly = True
    optReadWrite.Enabled = mblnDefinitionCreator
    optReadOnly.Enabled = mblnDefinitionCreator
    mblnReadOnly = ((mblnReadOnly Or Not mblnDefinitionCreator) And Not gfCurrentUserIsSysSecMgr)
  End Select
    
  SetComboItem cboCategories, IIf(IsNull(rsTemp.Fields("CategoryRecordID").Value), 0, rsTemp.Fields("CategoryRecordID").Value)
  SetComboItem cboTypes, IIf(IsNull(rsTemp.Fields("TypeRecordID").Value), 0, rsTemp.Fields("TypeRecordID").Value)
    
  SetComboItem cboTargetTable, IIf(IsNull(rsTemp.Fields("TargetTableID").Value), 0, rsTemp.Fields("TargetTableID").Value)
  SetComboItem cboTargetKeyField, IIf(IsNull(rsTemp.Fields("TargetKeyFieldColumnID").Value), 0, rsTemp.Fields("TargetKeyFieldColumnID").Value)
  SetComboItem cboTargetColumn, IIf(IsNull(rsTemp.Fields("TargetColumnID").Value), 0, rsTemp.Fields("TargetColumnID").Value)
  SetComboItem cboTargetCategory, IIf(IsNull(rsTemp.Fields("TargetCategoryColumnID").Value), 0, rsTemp.Fields("TargetCategoryColumnID").Value)
  SetComboItem cboTargetType, IIf(IsNull(rsTemp.Fields("TargetTypeColumnID").Value), 0, rsTemp.Fields("TargetTypeColumnID").Value)
  SetComboItem cboTargetTitle, IIf(IsNull(rsTemp.Fields("TargetTitleColumnID").Value), 0, rsTemp.Fields("TargetTitleColumnID").Value)

  SetComboItem cboParent1Table, IIf(IsNull(rsTemp.Fields("Parent1TableID").Value), 0, rsTemp.Fields("Parent1TableID").Value)
  SetComboItem cboParent1Keyfield, IIf(IsNull(rsTemp.Fields("Parent1KeyFieldColumnID").Value), 0, rsTemp.Fields("Parent1KeyFieldColumnID").Value)

  SetComboItem cboParent2Table, IIf(IsNull(rsTemp.Fields("Parent2TableID").Value), 0, rsTemp.Fields("Parent2TableID").Value)
  SetComboItem cboParent2Keyfield, IIf(IsNull(rsTemp.Fields("Parent2KeyFieldColumnID").Value), 0, rsTemp.Fields("Parent2KeyFieldColumnID").Value)

  If mblnReadOnly Then
    ControlsDisableAll Me
    EnableControl cmdCancel, True
  End If

  ' Tidy Up
  rsTemp.Close
  Set rsTemp = Nothing

End Sub

Private Function ValidateDefinition() As Boolean
  
  Dim bOK As Boolean
  Dim strErrorMessage As String
  
  bOK = True
  strErrorMessage = vbNullString
  
  ' General validation
  If Len(txtName.Text) = 0 Then
    strErrorMessage = strErrorMessage & "You must give this definition a name." & vbNewLine
  End If
    
' Check the name is unique
If Not CheckUniqueName(Trim(txtName.Text), mlngDocumentMapID) Then
  'tabImport.Tab = 0
  COAMsgBox "A Document Type definition called '" & Trim(txtName.Text) & "' already exists.", vbExclamation, Me.Caption
  txtName.SelStart = 0
  txtName.SelLength = Len(txtName.Text)
  ValidateDefinition = False
  Exit Function
End If
    
  ' Category type
  If GetComboItem(cboCategories) = 0 Then
    strErrorMessage = strErrorMessage & "Category value must be specified." & vbNewLine
  End If
  
  ' Target table
  If GetComboItem(cboTargetTable) = 0 Then
    strErrorMessage = strErrorMessage & "Target table must be specified." & vbNewLine
  End If
  
  ' Target column
  If GetComboItem(cboTargetColumn) = 0 Then
    strErrorMessage = strErrorMessage & "Document URL column must be specified." & vbNewLine
  End If
  
  ' Target keyfield
  If GetComboItem(cboTargetKeyField) = 0 Then
    strErrorMessage = strErrorMessage & "Key field column must be specified." & vbNewLine
  End If
   
   
  ' Keyfield of parent 1
  If Me.cboParent1Table.Enabled Then
    If GetComboItem(cboParent1Table) = 0 Then
      strErrorMessage = strErrorMessage & "Parent table must be specified." & vbNewLine
    End If
     
    If GetComboItem(cboParent1Keyfield) = 0 Then
      strErrorMessage = strErrorMessage & "Parent key field column must be specified." & vbNewLine
    End If
  End If
   
  
'  ' Keyfield of parent 2
'  If Me.cboParent2Table.Enabled Then
'    If GetComboItem(cboParent2Table) = 0 Then
'      strErrorMessage = strErrorMessage & "Parent 2 table must be specified." & vbNewLine
'    End If
'
'    If GetComboItem(cboParent2Keyfield) = 0 Then
'      strErrorMessage = strErrorMessage & "Parent 2 key field column must be specified." & vbNewLine
'    End If
'  End If
   
   
  If Len(strErrorMessage) > 0 Then
    bOK = False
    COAMsgBox strErrorMessage, vbExclamation + vbOKOnly, Me.Caption
  End If
  
  ValidateDefinition = bOK
  
End Function


Private Function SaveDefinition() As Boolean

  Dim rsTemp As ADODB.Recordset
  Dim sSQL As String
  Dim bOK As Boolean
  
  bOK = True
 
  If mlngDocumentMapID > 0 Then
    sSQL = "UPDATE dbo.[" & SQLTableDef & "] SET " _
              & "[Name] = '" & Replace(txtName.Text, "'", "''") & "', " _
              & "[Description] = '" & Replace(txtDesc.Text, "'", "''") & "', " _
              & "[Access] = " & IIf(optReadOnly.Value = True, "'RO'", "'RW'") & ", " _
              & "[TargetTableID] = " & GetComboItem(cboTargetTable) & ", " _
              & "[TargetKeyFieldColumnID] = " & GetComboItem(cboTargetKeyField) & ", " _
              & "[TargetColumnID] = " & GetComboItem(cboTargetColumn) & ", " _
              & "[TargetCategoryColumnID] = " & GetComboItem(cboTargetCategory) & ", " _
              & "[TargetTypeColumnID] = " & GetComboItem(cboTargetType) & ", " _
              & "[TargetTitleColumnID] = " & GetComboItem(cboTargetTitle) & ", " _
              & "[TargetGUIDColumnID] = 0, " _
              & "[Parent1TableID] = " & GetComboItem(cboParent1Table) & ", " _
              & "[Parent1KeyFieldColumnID] = " & GetComboItem(cboParent1Keyfield) & ", " _
              & "[Parent2TableID] = " & GetComboItem(cboParent2Table) & ", " _
              & "[Parent2KeyFieldColumnID] = " & GetComboItem(cboParent2Keyfield) & ", " _
              & "[CategoryRecordID] = " & GetComboItem(cboCategories) & ", " _
              & "[TypeRecordID] = " & GetComboItem(cboTypes) _
              & " WHERE [DocumentMapID] = " & CStr(mlngDocumentMapID) & ";"
    gADOCon.Execute sSQL, , adCmdText
    Call UtilUpdateLastSaved(utlDocumentMapping, mlngDocumentMapID)
               
  Else
  
    sSQL = "INSERT " & SQLTableDef & " (" _
              & " [Name], [Description]," _
              & " [UserName], [Access], " _
              & " [TargetTableID], [TargetKeyFieldColumnID], [TargetColumnID], [Parent1TableID], [Parent1KeyFieldColumnID], [Parent2TableID], [Parent2KeyFieldColumnID]," _
              & " [TargetCategoryColumnID], [TargetTypeColumnID], [TargetGUIDColumnID], [CategoryRecordID], [TypeRecordID], [TargetTitleColumnID]) " _
              & " VALUES('" _
              & Replace(txtName.Text, "'", "''") & "', '" & Replace(txtDesc.Text, "'", "''") _
              & "', '" & datGeneral.UserNameForSQL & "', " & IIf(optReadOnly.Value = True, "'RO'", "'RW'") _
              & ", " & GetComboItem(cboTargetTable) & ", " & GetComboItem(cboTargetKeyField) & ", " & GetComboItem(cboTargetColumn) _
              & ", " & GetComboItem(cboParent1Table) & ", " & GetComboItem(cboParent1Keyfield) _
              & ", " & GetComboItem(cboParent2Table) & ", " & GetComboItem(cboParent2Keyfield) _
              & ", " & GetComboItem(cboTargetCategory) & ", " & GetComboItem(cboTargetType) & ", 0" _
              & ", " & GetComboItem(cboCategories) & ", " & GetComboItem(cboTypes) _
              & ", " & GetComboItem(cboTargetTitle) & ");"

    mlngDocumentMapID = InsertDocumentMap(sSQL)
    Call UtilCreated(utlDocumentMapping, mlngDocumentMapID)
  
  End If
  
  SaveDefinition = bOK

End Function

Private Sub cboCategories_Click()
  PopulateTypesCombo
  Changed = True
End Sub

Private Sub cboParent1Keyfield_Change()
  Changed = True
End Sub

Private Sub cboParent1Keyfield_Click()
  Changed = True
End Sub

Private Sub cboParent1Table_Change()
  Changed = True
End Sub

Private Sub cboParent1Table_Click()

  Dim rsCols As ADODB.Recordset

  ' Columns for the parent table
  If cboParent1Table.Enabled Then
    Set rsCols = mclsGeneral.GetColumnNames(GetComboItem(cboParent1Table), False)
    Do While Not rsCols.EOF
    
      With cboParent1Keyfield
        .AddItem rsCols!ColumnName
        .ItemData(.NewIndex) = rsCols!ColumnID
      End With
      
      rsCols.MoveNext
    Loop
    rsCols.Close
  End If
  
  Changed = True
  
End Sub

Private Sub cboParent2Keyfield_Change()
  Changed = True
End Sub

Private Sub cboParent2Table_Change()
  Changed = True
End Sub

Private Sub cboParent2Table_Click()

  Dim rsCols As ADODB.Recordset

  ' Columns for the parent table
  If cboParent2Table.Enabled Then
    Set rsCols = mclsGeneral.GetColumnNames(GetComboItem(cboParent2Table), False)
    Do While Not rsCols.EOF
    
      With cboParent2Keyfield
        .AddItem rsCols!ColumnName
        .ItemData(.NewIndex) = rsCols!ColumnID
      End With
      
      rsCols.MoveNext
    Loop
    rsCols.Close
  End If

End Sub

Private Sub cboTargetCategory_Click()
  Changed = True
End Sub

Private Sub cboTargetColumn_Click()
  Changed = True
End Sub

Private Sub cboTargetKeyField_Click()
  Changed = True
End Sub

Private Sub cboTargetTable_Click()

  Dim sSQL As String
  Dim rsCols As ADODB.Recordset
  Dim rsParents As ADODB.Recordset
  Dim rsTables As ADODB.Recordset
  Dim iParents As Integer
  
  Screen.MousePointer = vbHourglass
  
  cboParent1Table.Clear
  cboParent1Keyfield.Clear
  cboParent2Table.Clear
  cboParent2Keyfield.Clear
     
  With cboTargetCategory
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboTargetType
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
'  With cboTargetGUID
'    .Clear
'    .AddItem "<None>"
'    .ItemData(.NewIndex) = 0
'  End With
  
  With cboTargetColumn
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboTargetKeyField
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  With cboTargetTitle
    .Clear
    .AddItem "<None>"
    .ItemData(.NewIndex) = 0
  End With
  
  'Get all the columns for the selected table
  
  
  
  Set rsCols = datGeneral.GetColumnNames(GetComboItem(cboTargetTable), False)
  Do While Not rsCols.EOF
  
    With cboTargetColumn
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetKeyField
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetCategory
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetType
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    With cboTargetTitle
      .AddItem rsCols!ColumnName
      .ItemData(.NewIndex) = rsCols!ColumnID
    End With
    
    rsCols.MoveNext
  Loop
  rsCols.Close


  ' Get parent tables
  sSQL = "SELECT [ParentID] FROM dbo.[ASRSysRelations] WHERE [ChildID] = " & GetComboItem(cboTargetTable) & ";"
  Set rsParents = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    
  iParents = rsParents.RecordCount

  Do While Not rsParents.EOF
    sSQL = "SELECT TableName, TableID FROM ASRSysTables " & _
           "WHERE TableID = " & rsParents!ParentID & " " & _
           ""
    Set rsTables = mdatData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
    Do While Not rsTables.EOF
      cboParent1Table.AddItem rsTables!TableName
      cboParent1Table.ItemData(cboParent1Table.NewIndex) = rsTables!TableID
      
      cboParent2Table.AddItem rsTables!TableName
      cboParent2Table.ItemData(cboParent1Table.NewIndex) = rsTables!TableID
      
      rsTables.MoveNext
    Loop
    rsParents.MoveNext
    rsTables.Close
  Loop
  rsParents.Close

  
  EnableControl lblParent1Table, (iParents > 0)
  EnableControl cboParent1Table, (iParents > 0)
  EnableControl lblParent1Keyfield, (iParents > 0)
  EnableControl cboParent1Keyfield, (iParents > 0)
  
  EnableControl lblParent2Table, (iParents > 1)
  EnableControl cboParent2Table, (iParents > 1)
  EnableControl lblParent2Keyfield, (iParents > 1)
  EnableControl cboParent2Keyfield, (iParents > 1)

  Changed = True

  Set rsParents = Nothing
  Set rsTables = Nothing
  Set rsCols = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub cboTargetTitle_Click()
  Changed = True
End Sub

Private Sub cboTargetType_Click()
  Changed = True
End Sub

Private Sub cboTypes_Change()
  Changed = True
End Sub

Private Sub cboTypes_Click()
  Changed = True
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()

  Dim fOK As Boolean
  
  If Not ValidateDefinition Then
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
    
  If Not SaveDefinition Then
    Screen.MousePointer = vbDefault
    Exit Sub
  End If
  
  Screen.MousePointer = vbDefault
  
  Me.Hide

End Sub


Private Function InsertDocumentMap(pstrSQL As String) As Long

  ' Insert definition into the name table and return the ID.

  On Error GoTo ErrorTrap

  Dim sSQL As String
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  Dim fSavedOK As Boolean
  
  fSavedOK = True
  
  Set cmADO = New ADODB.Command
  
  With cmADO
    .CommandText = "sp_ASRInsertNewUtility"
    .CommandType = adCmdStoredProc
    .CommandTimeout = 0
  
    Set .ActiveConnection = gADOCon
              
    Set pmADO = .CreateParameter("newID", adInteger, adParamOutput)
    .Parameters.Append pmADO
            
    Set pmADO = .CreateParameter("insertString", adLongVarChar, adParamInput, -1)
    .Parameters.Append pmADO
    pmADO.Value = pstrSQL
              
    Set pmADO = .CreateParameter("tablename", adVarChar, adParamInput, 255)
    .Parameters.Append pmADO
    pmADO.Value = SQLTableDef
              
    Set pmADO = .CreateParameter("idcolumnname", adVarChar, adParamInput, 30)
    .Parameters.Append pmADO
    pmADO.Value = "DocumentMapID"
              
    Set pmADO = Nothing
            
    cmADO.Execute
              
    If Not fSavedOK Then
      COAMsgBox "The new record could not be created." & vbCrLf & vbCrLf & _
        Err.Description, vbOKOnly + vbExclamation, app.ProductName
        InsertDocumentMap = 0
        Set cmADO = Nothing
        Exit Function
    End If
    
    InsertDocumentMap = IIf(IsNull(.Parameters(0).Value), 0, .Parameters(0).Value)
          
  End With
  
  Set cmADO = Nothing

  Exit Function
  
ErrorTrap:
  
  fSavedOK = False
  Resume Next
  
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

  IsV1ModuleSetupValid True

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Private Sub optReadOnly_Click()
  Changed = True
End Sub

Private Sub optReadWrite_Click()
  Changed = True
End Sub

Private Sub txtDesc_Change()
  Changed = True
End Sub

Private Sub txtName_Change()
  Changed = True
End Sub

Private Sub PopulateCategoriesCombo()

  Dim rsCategories As ADODB.Recordset
  Dim sTableName As String
  Dim sColumnName As String
  
  If Not IsV1ModuleSetupValid(False) Then Exit Sub
  
  sTableName = mclsGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYTABLE)))
  sColumnName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_CATEGORYCOLUMN)))

  cboCategories.Clear
  cboCategories.AddItem "<None>"
  cboCategories.ItemData(cboCategories.NewIndex) = 0

  Set rsCategories = mclsGeneral.GetRecords("SELECT DISTINCT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "]")
  With rsCategories
    Do While Not .EOF
    
      cboCategories.AddItem .Fields(1).Value
      cboCategories.ItemData(cboCategories.NewIndex) = .Fields(0).Value
      
      .MoveNext
    Loop
    
    .Close
  End With

  Set rsCategories = Nothing

  PopulateTypesCombo

End Sub

Private Sub PopulateTypesCombo()

  Dim rsTypes As ADODB.Recordset
  Dim sTableName As String
  Dim sColumnName As String
  Dim sSolumnCategoryName As String
  
  If Not IsV1ModuleSetupValid(False) Then Exit Sub
  
  sTableName = mclsGeneral.GetTableName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPETABLE)))
  sColumnName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECOLUMN)))
  sSolumnCategoryName = mclsGeneral.GetColumnName(Val(GetModuleParameter(MODULEKEY_DOCMANAGEMENT, PARAMETERKEY_DOCMAN_TYPECATEGORYCOLUMN)))

  cboTypes.Clear
  cboTypes.AddItem "<None>"
  cboTypes.ItemData(cboTypes.NewIndex) = 0

  Set rsTypes = mclsGeneral.GetRecords("SELECT [id], [" & sColumnName & "] FROM dbo.[" & sTableName & "] " & _
                      "WHERE [" & sSolumnCategoryName & "] = '" & Replace(cboCategories.Text, "'", "''") & "'")
  With rsTypes
  
    Do While Not .EOF
      cboTypes.AddItem (.Fields(1).Value)
      cboTypes.ItemData(cboTypes.NewIndex) = .Fields(0).Value
      
      .MoveNext
    Loop
    
    .Close
  End With

  Set rsTypes = Nothing

End Sub
Private Function CheckUniqueName(sName As String, lngCurrentID As Long) As Boolean

  Dim sSQL As String
  Dim rsTemp As Recordset
  
  sSQL = "SELECT * FROM ASRSysDocumentManagementTypes " & _
         " WHERE UPPER(Name) = '" & UCase(Replace(sName, "'", "''")) & "'" & _
         " AND DocumentMapID <> " & CStr(lngCurrentID)
  
  Set rsTemp = datGeneral.GetRecords(sSQL)
  
  If rsTemp.BOF And rsTemp.EOF Then
    CheckUniqueName = True
  Else
    CheckUniqueName = False
  End If
  
  Set rsTemp = Nothing

End Function


'
'Private mlngDocumentMapID As Long
'Private mblnCancelled As Boolean
'Private mblnReadOnly As Boolean
'Private mblnFromCopy As Boolean
'Private mblnDefinitionCreator As Boolean
'Private mbChanged As Boolean
'Private mlngLabelDefinitionID As Long
'Private mlngDocumentDefinitionID As Long
''Private datData As DataMgr.clsDataAccess
'Private mdatData As DataMgr.clsDataAccess
'Private mclsGeneral As DataMgr.clsGeneral

Private Sub txtUserName_Change()
  Changed = True
End Sub
