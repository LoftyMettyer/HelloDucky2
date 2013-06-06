VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAuditFilter2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Audit Log Filter"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   8007
   Icon            =   "frmAuditFilter2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraFilterItems 
      Caption         =   "Filter criteria :"
      Height          =   2550
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   6300
      Begin SSDataWidgets_B.SSDBGrid ssDBGridFilterCriteria 
         Height          =   1455
         Left            =   200
         TabIndex        =   6
         Top             =   300
         Width           =   5895
         ScrollBars      =   2
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         ColumnHeaders   =   0   'False
         GroupHeadLines  =   0
         HeadLines       =   0
         Col.Count       =   7
         AllowUpdate     =   0   'False
         MultiLine       =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowColumnMoving=   0
         AllowGroupSwapping=   0   'False
         AllowColumnSwapping=   0
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   3
         SelectByCell    =   -1  'True
         BalloonHelp     =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         Columns.Count   =   7
         Columns(0).Width=   3519
         Columns(0).Caption=   "FilterColumn"
         Columns(0).Name =   "DisplayColumnName"
         Columns(0).AllowSizing=   0   'False
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   2910
         Columns(1).Caption=   "FilterOperator"
         Columns(1).Name =   "DisplayOperator"
         Columns(1).AllowSizing=   0   'False
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   3519
         Columns(2).Caption=   "FilterText"
         Columns(2).Name =   "DisplayValue"
         Columns(2).AllowSizing=   0   'False
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "FilterColumnID"
         Columns(3).Name =   "FieldName"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "FilterOperatorID"
         Columns(4).Name =   "OperatorID"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "Datatype"
         Columns(5).Name =   "Datatype"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   3200
         Columns(6).Visible=   0   'False
         Columns(6).Caption=   "Special"
         Columns(6).Name =   "Special"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   10407
         _ExtentY        =   2558
         _StockProps     =   79
      End
      Begin VB.TextBox txtNoCriteria 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1450
         Left            =   200
         TabIndex        =   9
         Text            =   "<Add criteria from below to this list>"
         Top             =   300
         Width           =   5900
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   400
         Left            =   3525
         TabIndex        =   7
         Top             =   1950
         Width           =   1200
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Remove &All"
         Height          =   400
         Left            =   4890
         TabIndex        =   8
         Top             =   1950
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   3795
      TabIndex        =   4
      Top             =   4620
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   5145
      TabIndex        =   5
      Top             =   4620
      Width           =   1200
   End
   Begin VB.Frame fraNewFilterCriteria 
      Caption         =   "Define more criteria :"
      Height          =   1710
      Left            =   60
      TabIndex        =   11
      Top             =   2745
      Width           =   6300
      Begin VB.CommandButton cmdAddToList 
         Caption         =   "A&dd to List"
         Height          =   400
         Left            =   4890
         TabIndex        =   3
         Top             =   1080
         Width           =   1200
      End
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   555
         Width           =   1995
      End
      Begin VB.ComboBox cboOperators 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   550
         Width           =   1500
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   4100
         TabIndex        =   2
         Top             =   550
         Width           =   2000
      End
      Begin VB.ComboBox cboValues 
         Height          =   315
         ItemData        =   "frmAuditFilter2.frx":000C
         Left            =   4100
         List            =   "frmAuditFilter2.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   550
         Width           =   2000
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Field :"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lbCondition 
         BackStyle       =   0  'Transparent
         Caption         =   "Operator :"
         Height          =   195
         Left            =   2400
         TabIndex        =   14
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   4095
         TabIndex        =   13
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAuditFilter2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private miAuditType As audType
Private mblnCancelled As Boolean

Public Property Get FilterArray() As Variant
  
  ' Return the array of filter criteria.
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim avFilterArray() As Variant
    
  ' NB. the filter array has 6 columns :
  ' Column 1 - the column display name.
  ' Column 2 - the operator display.
  ' Column 3 - the value.
  ' Column 4 - the column fieldname.
  ' Column 5 - the operator ID.
  ' Column 6 - the datatype.
  
  ' Refresh the related columns array.
  With ssDBGridFilterCriteria
    ReDim avFilterArray(6, .Rows)
    
    .MoveFirst
    
    For iLoop = 1 To .Rows
      avFilterArray(1, iLoop) = .Columns("DisplayColumnName").Text
      avFilterArray(2, iLoop) = .Columns("DisplayOperator").Text
      avFilterArray(3, iLoop) = .Columns("DisplayValue").Text
      avFilterArray(4, iLoop) = .Columns("FieldName").Text
      avFilterArray(5, iLoop) = .Columns("OperatorID").Text
      avFilterArray(6, iLoop) = .Columns("DataType").Text
      
      .MoveNext
    Next iLoop
  End With
    
  FilterArray = avFilterArray
  
End Property

Private Sub DeCodeArray(myarray() As Variant)

  Dim iCount As Integer
  
  For iCount = 1 To UBound(myarray(), 2)
    With ssDBGridFilterCriteria
      .AddItem myarray(1, iCount) & vbTab & _
               myarray(2, iCount) & vbTab & _
               myarray(3, iCount) & vbTab & _
               myarray(4, iCount) & vbTab & _
               myarray(5, iCount) & vbTab & _
               myarray(6, iCount)
    End With
  Next iCount

End Sub

Public Function Initialise(iAuditType As audType, myarray() As Variant) As Boolean

  ' Set the module level variable of the audit type
  miAuditType = iAuditType
  
  ' Set the filter array to be the one passed in
  If UBound(myarray(), 2) > 0 Then
    DeCodeArray myarray
  End If
  
  Select Case miAuditType
  
    Case audRecords
    
      With cboColumns
        .AddItem "User"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Date / Time"
        .ItemData(.NewIndex) = sqlDate
        .AddItem "Table"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Column"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Old Value"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "New Value"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Record Description"
        .ItemData(.NewIndex) = sqlVarchar
      End With
      
    Case audPermissions
    
      With cboColumns
        .AddItem "User"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Date / Time"
        .ItemData(.NewIndex) = sqlDate
        .AddItem "User Group"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "View / Table"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Column"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Action"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Permission"
        .ItemData(.NewIndex) = sqlVarchar
      End With
    
    Case audGroups
  
      With cboColumns
        .AddItem "User"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Date / Time"
        .ItemData(.NewIndex) = sqlDate
        .AddItem "User Group"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "User Login"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Action"
        .ItemData(.NewIndex) = sqlVarchar
      End With
  
    Case audAccess

      With cboColumns
        .AddItem "Date / Time"
        .ItemData(.NewIndex) = sqlDate
        .AddItem "User Group"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "User"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Computer Name"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Module"
        .ItemData(.NewIndex) = sqlVarchar
        .AddItem "Action"
        .ItemData(.NewIndex) = sqlVarchar
      End With

  End Select
  
  RefreshControls
  
End Function

Private Function GetFieldName(strColumnDisplayName As String) As String

  Select Case miAuditType
  
    Case audRecords
    
      If strColumnDisplayName = "User" Then
        GetFieldName = "Username"
      ElseIf strColumnDisplayName = "Date / Time" Then
        GetFieldName = "DateTimeStamp"
      ElseIf strColumnDisplayName = "Table" Then
        GetFieldName = "TableName"
      ElseIf strColumnDisplayName = "Column" Then
        GetFieldName = "ColumnName"
      ElseIf strColumnDisplayName = "Old Value" Then
        GetFieldName = "OldValue"
      ElseIf strColumnDisplayName = "New Value" Then
        GetFieldName = "NewValue"
      ElseIf strColumnDisplayName = "Record Description" Then
        GetFieldName = "RecordDesc"
      Else
        Debug.Assert False
      End If
      
    Case audPermissions
    
      If strColumnDisplayName = "User" Then
        GetFieldName = "Username"
      ElseIf strColumnDisplayName = "Date / Time" Then
        GetFieldName = "DateTimeStamp"
      ElseIf strColumnDisplayName = "User Group" Then
        GetFieldName = "GroupName"
      ElseIf strColumnDisplayName = "Table" Then
        GetFieldName = "ViewTableName"
      ElseIf strColumnDisplayName = "Column" Then
        GetFieldName = "ColumnName"
      ElseIf strColumnDisplayName = "Action" Then
        GetFieldName = "Action"
      ElseIf strColumnDisplayName = "Permission" Then
        GetFieldName = "Permission"
      ElseIf strColumnDisplayName = "View / Table" Then
       GetFieldName = "viewTableName"
      Else
        Debug.Assert False
      End If
    
    Case audGroups
  
      If strColumnDisplayName = "User" Then
        GetFieldName = "Username"
      ElseIf strColumnDisplayName = "Date / Time" Then
        GetFieldName = "DateTimeStamp"
      ElseIf strColumnDisplayName = "User Group" Then
        GetFieldName = "GroupName"
      ElseIf strColumnDisplayName = "User Login" Then
        GetFieldName = "UserLogin"
      ElseIf strColumnDisplayName = "Action" Then
        GetFieldName = "Action"
      Else
        Debug.Assert False
      End If
  
    Case audAccess

      If strColumnDisplayName = "Date / Time" Then
        GetFieldName = "DateTimeStamp"
      ElseIf strColumnDisplayName = "User Group" Then
        GetFieldName = "UserGroup"
      ElseIf strColumnDisplayName = "User" Then
        GetFieldName = "Username"
      ElseIf strColumnDisplayName = "Computer Name" Then
        GetFieldName = "ComputerName"
      ElseIf strColumnDisplayName = "Module" Then
        GetFieldName = "HRProModule"
      ElseIf strColumnDisplayName = "Action" Then
        GetFieldName = "Action"
      Else
        Debug.Assert False
      End If

  End Select


End Function

Private Function GetDataType(strColumn As String) As Integer

  If strColumn = "Date / Time" Then
    GetDataType = sqlDate
  Else
    GetDataType = sqlVarchar
  End If
  
End Function


Private Sub RefreshOperators()
  ' Populate the Filter Operator combo.
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iDataType As Integer
  
  ' Get the Filter Column data type.
  iDataType = GetDataType(cboColumns.Text)
  
  With cboOperators
    
    ' Clear the combo.
    .Clear
    
    If iDataType = sqlDate Then
      .AddItem OperatorDescription(giFILTEROP_ON)
      .ItemData(.NewIndex) = giFILTEROP_ON
      .AddItem OperatorDescription(giFILTEROP_NOTON)
      .ItemData(.NewIndex) = giFILTEROP_NOTON
      .AddItem OperatorDescription(giFILTEROP_AFTER)
      .ItemData(.NewIndex) = giFILTEROP_AFTER
      .AddItem OperatorDescription(giFILTEROP_BEFORE)
      .ItemData(.NewIndex) = giFILTEROP_BEFORE
      'TM20011102 Fault 2155
      .AddItem OperatorDescription(giFILTEROP_ONORAFTER)
      .ItemData(.NewIndex) = giFILTEROP_ONORAFTER
      .AddItem OperatorDescription(giFILTEROP_ONORBEFORE)
      .ItemData(.NewIndex) = giFILTEROP_ONORBEFORE
    Else
      .AddItem OperatorDescription(giFILTEROP_IS)
      .ItemData(.NewIndex) = giFILTEROP_IS
      .AddItem OperatorDescription(giFILTEROP_ISNOT)
      .ItemData(.NewIndex) = giFILTEROP_ISNOT
      .AddItem OperatorDescription(giFILTEROP_CONTAINS)
      .ItemData(.NewIndex) = giFILTEROP_CONTAINS
'      .AddItem OperatorDescription(giFILTEROP_DOESNOTCONTAIN)
'      .ItemData(.NewIndex) = giFILTEROP_DOESNOTCONTAIN
    End If
          
    .Enabled = .ListCount > 0
    txtValue.Enabled = .Enabled
    cboValues.Enabled = .Enabled
    
    If .ListCount > 0 Then
      .ListIndex = 0
    End If
  
  End With
  
  
  'MH20010118 Fault 1633
  'When the user entered a number which was too large caused
  'run time error. This is a simple way to stop it
  If iDataType = sqlNumeric Or iDataType = sqlInteger Then
    txtValue.MaxLength = 9
  Else
    txtValue.MaxLength = 0
  End If
  
  
  txtValue.Visible = (iDataType <> sqlBoolean)
  cboValues.Visible = (iDataType = sqlBoolean)
  cboValues.ListIndex = 0

End Sub

Private Sub cboColumns_Click()

  RefreshOperators

End Sub

Private Sub cmdCancel_Click()

  mblnCancelled = True
  Me.Hide
  
End Sub

Private Sub cmdOK_Click()

  mblnCancelled = False
  Me.Hide
  
End Sub

Private Sub Form_Activate()
'NHRD11072003 Fault 4158
Dim iRowIndex As Integer
With ssDBGridFilterCriteria
  If .Rows > 0 Then
    If iRowIndex = 0 Then
      .MoveFirst
    ElseIf iRowIndex >= .Rows Then
      .MoveLast
    End If
    .SelBookmarks.Add .Bookmark
  End If
End With
  
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
  
  If UnloadMode = vbFormControlMenu Then
    cmdCancel_Click
    Exit Sub
  End If

End Sub



Public Property Get Cancelled() As Boolean

  Cancelled = mblnCancelled
  
End Property

Private Function OperatorDescription(piOperatorCode As Integer) As String
  ' Return the textual description og the given operator.
  Dim sDesc As String
  
  Select Case piOperatorCode
    Case giFILTEROP_EQUALS
      sDesc = "is equal to"
'      sDesc = "equals"
    Case giFILTEROP_NOTEQUALTO
      sDesc = "is NOT equal to"
'      sDesc = "not equal to"
    Case giFILTEROP_ISATMOST
      sDesc = "is at most"
    Case giFILTEROP_ISATLEAST
      sDesc = "is at least"
    Case giFILTEROP_ISMORETHAN
      sDesc = "is more than"
    Case giFILTEROP_ISLESSTHAN
      sDesc = "is less than"
    
    
    
    Case giFILTEROP_ON
      'sDesc = "on"
      sDesc = "is equal to"
    Case giFILTEROP_NOTON
      'sDesc = "not on"
      sDesc = "is NOT equal to"
    Case giFILTEROP_AFTER
      sDesc = "after"
      'sDesc = "greater than"
    Case giFILTEROP_BEFORE
      sDesc = "before"
      'sDesc = "less than"
    Case giFILTEROP_ONORAFTER
      'sDesc = "on or after"
      sDesc = "is equal to or after"
    Case giFILTEROP_ONORBEFORE
      'sDesc = "on or before"
      sDesc = "is equal to or before"
    
    
    
    Case giFILTEROP_CONTAINS
      sDesc = "contains"
    Case giFILTEROP_IS
'      sDesc = "is"
      sDesc = "is equal to"
    Case giFILTEROP_DOESNOTCONTAIN
      sDesc = "does not contain"
    Case giFILTEROP_ISNOT
'      sDesc = "is not"
      sDesc = "is NOT equal to"
    Case Else
      sDesc = ""
  End Select
  
  OperatorDescription = sDesc
  
End Function


Private Sub cmdRemove_Click()
  ' Remove the selected row from the grid.
  Dim iRowIndex As Integer
  
  With ssDBGridFilterCriteria
    iRowIndex = .AddItemRowIndex(.Bookmark)
    
    ' RH 20/07/00 - Bug Fix 633. Wierd grid thing - doesnt delete the last row.
    If .Rows = 1 And iRowIndex = 0 Then
      .RemoveAll
    Else
      '.RemoveItem iRowIndex
      .DeleteSelected
    End If

    If .Rows > 0 Then
      If iRowIndex = 0 Then
        .MoveFirst
      ElseIf iRowIndex >= .Rows Then
        .MoveLast
      End If

      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  RefreshControls

End Sub

Private Sub cmdClearAll_Click()
  ' Clearr all criteria from the grid.
  ssDBGridFilterCriteria.RemoveAll
  RefreshControls
  
End Sub

Private Sub cmdAddToList_Click()
  
  ' Add the defined filter criteria to the grid.
  Dim fValidParameters As Boolean
  Dim iDataType As Integer
  
  ' Check that all parameters are valid.
  ' Check that a valid column has been selected.
  fValidParameters = (cboColumns.ListIndex <> -1)
  
  If Not fValidParameters Then
    MsgBox "Invalid column selected.", vbExclamation, App.ProductName
    cboColumns.SetFocus
  End If
  
  ' Check that a valid operator has been selected.
  If fValidParameters Then
    fValidParameters = (cboOperators.ListIndex <> -1)
    
    If Not fValidParameters Then
      MsgBox "Invalid operator selected.", vbExclamation, App.ProductName
      cboOperators.SetFocus
    End If
  End If
    
  ' Check that the value entered is valid.
  
  If fValidParameters Then
    
    iDataType = cboColumns.ItemData(cboColumns.ListIndex) 'GetDataType(cboColumns.Text)
    
    Select Case iDataType
      
      Case sqlDate ' Date columns.
        ' Ensure that the value entered is a date.
        fValidParameters = (Len(txtValue.Text) = 0) Or _
          (IsDate(txtValue.Text))
        If Not fValidParameters Then
          txtValue.ForeColor = vbRed
          MsgBox "You have entered an invalid date.", vbExclamation, App.ProductName
          txtValue.ForeColor = vbWindowText
          txtValue.Text = ""
          txtValue.SetFocus
        End If
      
      Case sqlVarchar
        ' No validation required.
          
    End Select
  End If
    
  ' Add the definition to the grid.
  ' Grid column order is :
  ' Display column name, operator, value, fieldname, operator ID.
  If fValidParameters Then
    With ssDBGridFilterCriteria
      
      .AddItem cboColumns.List(cboColumns.ListIndex) & vbTab & _
               cboOperators.List(cboOperators.ListIndex) & vbTab & _
               txtValue.Text & vbTab & _
               GetFieldName(cboColumns.List(cboColumns.ListIndex)) & vbTab & _
               cboOperators.ItemData(cboOperators.ListIndex) & vbTab & _
               iDataType
      .MoveLast
      .SelBookmarks.RemoveAll
      .SelBookmarks.Add .Bookmark
      RefreshControls
      
      cboColumns.ListIndex = 0
      cboOperators.ListIndex = 0
      txtValue.Text = ""
      cboValues.ListIndex = 0
      cboColumns.SetFocus
    End With
  End If
  
End Sub

Private Sub RefreshControls()
  ' Refresh the screen controls.
  ssDBGridFilterCriteria.Visible = (ssDBGridFilterCriteria.Rows > 0)
  txtNoCriteria.Visible = (ssDBGridFilterCriteria.Rows = 0)
  cmdRemove.Enabled = (ssDBGridFilterCriteria.Rows > 0)
  cmdClearAll.Enabled = (ssDBGridFilterCriteria.Rows > 0)
  
  'TM20020116 Fault 3355 - select the first item in the list.
  If cboColumns.ListCount > 0 Then
    cboColumns.ListIndex = 0
  End If
  
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub ssDBGridFilterCriteria_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
'NHRD11072003 Fault 4158
DispPromptMsg = False
With ssDBGridFilterCriteria
  'here I am
  If .Rows > 0 Then
    .MoveLast
  End If
End With
End Sub

Private Sub txtValue_GotFocus()

  cmdAddToList.Default = True
  
End Sub

Private Sub txtValue_LostFocus()

  cmdAddToList.Default = False
  cmdOk.Default = False
  
End Sub

