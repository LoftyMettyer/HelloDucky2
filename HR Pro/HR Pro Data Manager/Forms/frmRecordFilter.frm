VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{604A59D5-2409-101D-97D5-46626B63EF2D}#1.0#0"; "TDBNumbr.ocx"
Begin VB.Form frmRecordFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Record Filter"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1055
   Icon            =   "frmRecordFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraFilterItems 
      Caption         =   "Filter criteria :"
      Height          =   2500
      Left            =   150
      TabIndex        =   11
      Top             =   100
      Width           =   7850
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Remove &All"
         Height          =   400
         Left            =   6400
         TabIndex        =   9
         Top             =   1900
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   400
         Left            =   5000
         TabIndex        =   8
         Top             =   1900
         Width           =   1200
      End
      Begin SSDataWidgets_B.SSDBGrid ssDBGridFilterCriteria 
         Height          =   1455
         Left            =   200
         TabIndex        =   10
         Top             =   300
         Width           =   7400
         _Version        =   196617
         DataMode        =   2
         RecordSelectors =   0   'False
         GroupHeaders    =   0   'False
         ColumnHeaders   =   0   'False
         GroupHeadLines  =   0
         HeadLines       =   0
         Col.Count       =   6
         DividerType     =   2
         AllowUpdate     =   0   'False
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
         SelectTypeRow   =   1
         BalloonHelp     =   0   'False
         MaxSelectedRows =   1
         ForeColorEven   =   0
         BackColorEven   =   -2147483643
         BackColorOdd    =   -2147483643
         RowHeight       =   423
         ExtraHeight     =   185
         Columns.Count   =   6
         Columns(0).Width=   4233
         Columns(0).Caption=   "FilterColumn"
         Columns(0).Name =   "FilterColumn"
         Columns(0).AllowSizing=   0   'False
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   4419
         Columns(1).Caption=   "FilterOperator"
         Columns(1).Name =   "FilterOperator"
         Columns(1).AllowSizing=   0   'False
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   4339
         Columns(2).Caption=   "FilterText"
         Columns(2).Name =   "FilterText"
         Columns(2).AllowSizing=   0   'False
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "FilterColumnID"
         Columns(3).Name =   "FilterColumnID"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "FilterOperatorID"
         Columns(4).Name =   "FilterOperatorID"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "FilterColumnDataType"
         Columns(5).Name =   "FilterColumnDataType"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         TabNavigation   =   1
         _ExtentX        =   13053
         _ExtentY        =   2566
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
      Begin VB.TextBox txtNoCriteria 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1450
         Left            =   200
         TabIndex        =   16
         Text            =   "<Add criteria from below to this list>"
         Top             =   300
         Width           =   7400
      End
   End
   Begin VB.Frame fraNewFilterCriteria 
      Caption         =   "Define more criteria :"
      Height          =   1600
      Left            =   150
      TabIndex        =   12
      Top             =   2700
      Width           =   7850
      Begin VB.ComboBox cboColumns 
         Height          =   315
         Left            =   200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   550
         Width           =   2200
      End
      Begin VB.ComboBox cboOperators 
         Height          =   315
         Left            =   2500
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   550
         Width           =   2650
      End
      Begin VB.TextBox txtStringValue 
         Height          =   315
         Left            =   5250
         TabIndex        =   2
         Top             =   720
         Width           =   2400
      End
      Begin VB.ComboBox cboValues 
         Height          =   315
         ItemData        =   "frmRecordFilter.frx":000C
         Left            =   5250
         List            =   "frmRecordFilter.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   550
         Width           =   2400
      End
      Begin VB.CommandButton cmdAddToList 
         Caption         =   "A&dd to List"
         Default         =   -1  'True
         Height          =   400
         Left            =   6400
         TabIndex        =   5
         Top             =   1000
         Width           =   1200
      End
      Begin TDBNumberCtrl.TDBNumber tdbNumberValue 
         Height          =   315
         Left            =   5250
         TabIndex        =   4
         Top             =   885
         Visible         =   0   'False
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   999999999
         MinValue        =   -999999999
         Value           =   0
         SelStart        =   0
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   ""
         Format          =   "###############"
         DisplayFormat   =   "###############"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "&Caption"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmRecordFilter.frx":0027
         MousePointer    =   0
      End
      Begin VB.Label lblValue 
         BackStyle       =   0  'Transparent
         Caption         =   "Value :"
         Height          =   195
         Left            =   5295
         TabIndex        =   15
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbCondition 
         BackStyle       =   0  'Transparent
         Caption         =   "Operator :"
         Height          =   195
         Left            =   2500
         TabIndex        =   14
         Top             =   300
         Width           =   1050
      End
      Begin VB.Label lblColumn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Field :"
         Height          =   195
         Left            =   195
         TabIndex        =   13
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   6800
      TabIndex        =   7
      Top             =   4450
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   400
      Left            =   5400
      TabIndex        =   6
      Top             =   4450
      Width           =   1200
   End
End
Attribute VB_Name = "frmRecordFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfCancelled As Boolean
Private mvar_rsRecords As ADODB.Recordset
Private mvar_objTableView As CTablePrivilege
Private mvar_objColumnPrivileges As CColumnPrivileges

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
  
End Property

Public Property Get FilterArray() As Variant
  ' Return the array of filter criteria.
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim avFilterArray() As Variant
    
  ' NB. the filter array has 3 columns :
  ' Column 1 - the filter column ID.
  ' Column 2 - the filter operator ID.
  ' Column 3 - the filter value.
  
  ' Refresh the related columns array.
  With ssDBGridFilterCriteria
    ReDim avFilterArray(3, .Rows)
    
    .MoveFirst
    
    For iLoop = 1 To .Rows
      avFilterArray(1, iLoop) = .Columns("FilterColumnID").Text
      avFilterArray(2, iLoop) = .Columns("FilterOperatorID").Text
      
      Select Case .Columns("FilterColumnDataType").Text
        Case sqlNumeric
          avFilterArray(3, iLoop) = Replace(.Columns("FilterText").Text, ",", "")
       
        Case Else
          avFilterArray(3, iLoop) = .Columns("FilterText").Text
      
      End Select
      
      .MoveNext
    Next iLoop
  End With
    
  FilterArray = avFilterArray
  
End Property


Private Sub cboColumns_Click()
  ' Update the filter condition combo with the appropriate conditions for the current column data type.
  cboOperators_Refresh
  
End Sub

Private Sub cmdAddToList_Click()
  ' Add the defined filter criteria to the grid.
  Dim fValidParameters As Boolean
  Dim iDataType As Integer
  Dim strSearchValue As String
  
  ' Check that all parameters are valid.
  ' Check that a valid column has been selected.
  fValidParameters = (cboColumns.ItemData(cboColumns.ListIndex) > 0)
  
  If Not fValidParameters Then
    COAMsgBox "Invalid column selected.", vbExclamation, App.ProductName
    cboColumns.SetFocus
  End If
  
  ' Check that a valid operator has been selected.
  If fValidParameters Then
    fValidParameters = (cboOperators.ItemData(cboOperators.ListIndex) > 0)
    
    If Not fValidParameters Then
      COAMsgBox "Invalid operator selected.", vbExclamation, App.ProductName
      cboOperators.SetFocus
    End If
  End If
    
  ' Check that the value entered is valid.
  If fValidParameters Then
    ' Get the Filter Column data type.
    fValidParameters = mvar_objColumnPrivileges.IsValid(cboColumns.List(cboColumns.ListIndex))
    
    If Not fValidParameters Then
      COAMsgBox "Invalid column selected.", vbExclamation, App.ProductName
      cboColumns.SetFocus
    End If
  End If
  
  If fValidParameters Then
    iDataType = mvar_objColumnPrivileges.Item(cboColumns.List(cboColumns.ListIndex)).DataType
    
    Select Case iDataType
      Case sqlOle  ' Not required as OLEs are not permitted in the Quick Filter Column selection.
      Case sqlBoolean ' Logic columns. Values are tied to the combo, so must be valid.
        
      Case sqlNumeric, sqlInteger ' Numeric and Integer columns.
        ' Ensure that the value entered is numeric.
        If Len(tdbNumberValue.Text) = 0 Then
          tdbNumberValue.Text = "0"
        End If
        fValidParameters = IsNumeric(Replace(tdbNumberValue.Text, ",", ""))
        If Not fValidParameters Then
          COAMsgBox "Invalid numeric value entered.", vbExclamation, App.ProductName
          tdbNumberValue.SetFocus
        End If

      Case sqlDate ' Date columns.
        ' Ensure that the value entered is a date.
        
        'AE20071123 Fault #7892
        'fValidParameters = (IsDate(txtStringValue.Text)
        fValidParameters = (IsDate(txtStringValue.Text) And (ValidateDate(txtStringValue.Text) <> vbNullString))
        
        'JPD 20031022 Fault 7350
        ' Check for dates that SQL can't handle (ie. earlier than 1/1/1753)
        If fValidParameters Then
          fValidParameters = (CDate(txtStringValue.Text) >= #1/1/1753#)
        End If
        ' Check for times coming in (eg. 1.1.1)
        If fValidParameters Then
          fValidParameters = ConvertSQLDateToLocale(Format(CDate(txtStringValue.Text), "mm/dd/yyyy")) = CDate(txtStringValue.Text)
        End If
        
        If fValidParameters Then
          txtStringValue.Text = IIf(Len(txtStringValue.Text) = 0, "", CDate(txtStringValue.Text))
        Else
          If Len(txtStringValue.Text) = 0 Then
            txtStringValue.Text = ""
            fValidParameters = True
          Else
            txtStringValue.ForeColor = vbRed
            COAMsgBox "You have entered an invalid date.", vbExclamation, App.ProductName
            txtStringValue.ForeColor = vbWindowText
            txtStringValue.Text = ""
            txtStringValue.SetFocus
          End If
        End If
      
      Case sqlVarChar, sqlVarBinary, sqlLongVarChar  ' Character and Photo columns (photo columns are really character columns).
        ' No validation required.
          
    End Select
  End If
    
  ' Add the definition to the grid.
  ' Grid column order is the filter column name, filter operator, filter value,
  ' filter colmn ID, filter operator ID.
  If fValidParameters Then
    
    Select Case iDataType
      Case sqlNumeric, sqlInteger
        strSearchValue = tdbNumberValue.Text
      Case sqlBoolean
        strSearchValue = cboValues.List(cboValues.ListIndex)
      Case Else
        strSearchValue = txtStringValue.Text
    End Select
    
    With ssDBGridFilterCriteria
      .AddItem cboColumns.List(cboColumns.ListIndex) & vbTab & _
        cboOperators.List(cboOperators.ListIndex) & vbTab & _
        strSearchValue & vbTab & _
        cboColumns.ItemData(cboColumns.ListIndex) & vbTab & _
        cboOperators.ItemData(cboOperators.ListIndex) & vbTab & iDataType
      .MoveLast
      .SelBookmarks.Add .Bookmark
      RefreshControls
      
      cboColumns.ListIndex = 0
      cboOperators.ListIndex = 0
      txtStringValue.Text = ""
      cboValues.ListIndex = 0
      cboColumns.SetFocus
    End With
  End If
  
  DoGridSize
  
End Sub

Private Sub DoGridSize()

If ssDBGridFilterCriteria.Rows > ssDBGridFilterCriteria.VisibleRows Then

  ' Show The Scrollbars
  ssDBGridFilterCriteria.ScrollBars = ssScrollBarsVertical
  ssDBGridFilterCriteria.Columns(2).Width = 2230

Else

  ' Hide The Scrollbars
  ssDBGridFilterCriteria.ScrollBars = ssScrollBarsNone
  ssDBGridFilterCriteria.Columns(2).Width = 2390

End If


End Sub

Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide
  
End Sub


Private Sub cmdClearAll_Click()
  ' Clearr all criteria from the grid.
  ssDBGridFilterCriteria.RemoveAll
  RefreshControls
  DoGridSize
  
End Sub

Private Sub cmdOK_Click()
  mfCancelled = False
  Me.Hide

End Sub


Private Sub cboColumns_Populate()
  ' Populate the filter columns combo.
  Dim iDataType As Integer
  Dim iColumnType As Integer
  Dim lngColumnID As Long
'  Dim fldFilter As Field
  Dim fldFilter As ADODB.Field
  Dim objTableView As CTablePrivilege
  Dim sTableViewName As String

  With cboColumns
    ' Clear the filter columns combo.
    .Clear
    
    ' Add an item to the combo for each column in the current recordset.
    ' NB. Do not add systems, link, Photo or OLE columns, or columns from parent tables.
    For Each fldFilter In mvar_rsRecords.Fields
      
      
      'MH20010119 Fault 1651
      'This fixes a fault which we only managed to recreate on SQL2000
      'but I'm not convinced that it is SQL2000 only so I have left it
      'in for SQL 7.0 too (Don't think it will do any harm).

      'If (Left(fldFilter.Name, 1) <> "?") And _
        (UCase(Trim(fldFilter.Properties("BASETABLENAME"))) = UCase(Trim(mvar_objTableView.TableName))) Then
      If Not IsNull(fldFilter.Properties("BASETABLENAME")) Then
        
        sTableViewName = fldFilter.Properties("BASETABLENAME")
        Set objTableView = gcoTablePrivileges.FindRealSource(sTableViewName)
        If Not objTableView Is Nothing Then
          sTableViewName = objTableView.TableName
        End If

        If (Left(fldFilter.Name, 1) <> "?") And _
          (UCase(Trim(sTableViewName)) = UCase(Trim(mvar_objTableView.TableName))) Then
          
          
          ' Get the column's type and data type.
          If mvar_objColumnPrivileges.IsValid(fldFilter.Name) Then
            iDataType = mvar_objColumnPrivileges.Item(fldFilter.Name).DataType
            iColumnType = mvar_objColumnPrivileges.Item(fldFilter.Name).ColumnType
            lngColumnID = mvar_objColumnPrivileges.Item(fldFilter.Name).ColumnID
                  
            If (iColumnType <> colSystem) And _
              (iColumnType <> colLink) And _
              (iDataType <> sqlVarBinary) And _
              (iDataType <> sqlOle) Then
  
              ' Add the column to the combo list.
              .AddItem fldFilter.Name
              .ItemData(.NewIndex) = lngColumnID
            End If
          End If
        End If
    
      End If  'MH20010119
    
    Next
    
    .Enabled = (.ListCount > 0)
    If .ListCount > 0 Then
      .ListIndex = 0
    End If
    
    cboOperators_Refresh
    txtStringValue.Enabled = .Enabled
    cboValues.Enabled = .Enabled
    cmdAddToList.Enabled = .Enabled
  End With

End Sub

Public Sub Initialise(ByVal prsNewValue As ADODB.Recordset, pobjTableView As CTablePrivilege, pavFilterArray As Variant)
  ' initialise the record edit filter form.
  Set mvar_rsRecords = prsNewValue
  Set mvar_objTableView = pobjTableView
  
  Set mvar_objColumnPrivileges = GetColumnPrivileges(IIf(mvar_objTableView.ViewID > 0, mvar_objTableView.ViewName, mvar_objTableView.TableName))
    
  cboColumns_Populate
  ssDBGridFilterCriteria_Populate pavFilterArray
  
  RefreshControls
  
  tdbNumberValue.Top = cboValues.Top
  tdbNumberValue.Left = cboValues.Left
  
  txtStringValue.Top = cboValues.Top
  txtStringValue.Left = cboValues.Left
  
End Sub


Private Sub cboOperators_Refresh()
  ' Populate the Filter Operator combo.
  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim iDataType As SQLDataType
  Dim lngColumnID As Long
  
  ' Get the Filter Column data type.
  fOK = mvar_objColumnPrivileges.IsValid(cboColumns.List(cboColumns.ListIndex))
  If fOK Then
    iDataType = mvar_objColumnPrivileges.Item(cboColumns.List(cboColumns.ListIndex)).DataType
    lngColumnID = mvar_objColumnPrivileges.Item(cboColumns.List(cboColumns.ListIndex)).ColumnID
  End If
  
  With cboOperators
    ' Clear the combo.
    .Clear
    
    If fOK Then
      Select Case iDataType
        Case sqlOle  ' Not required as OLEs are not permitted in the Quick Filter Column selection.
        Case sqlBoolean ' Logic columns.
          .AddItem OperatorDescription(giFILTEROP_EQUALS)
          .ItemData(.NewIndex) = giFILTEROP_EQUALS
          
        Case sqlNumeric, sqlInteger ' Numeric and Integer columns.
          .AddItem OperatorDescription(giFILTEROP_EQUALS)
          .ItemData(.NewIndex) = giFILTEROP_EQUALS
          .AddItem OperatorDescription(giFILTEROP_NOTEQUALTO)
          .ItemData(.NewIndex) = giFILTEROP_NOTEQUALTO
          .AddItem OperatorDescription(giFILTEROP_ISMORETHAN)
          .ItemData(.NewIndex) = giFILTEROP_ISMORETHAN
          .AddItem OperatorDescription(giFILTEROP_ISATLEAST)
          .ItemData(.NewIndex) = giFILTEROP_ISATLEAST
          .AddItem OperatorDescription(giFILTEROP_ISLESSTHAN)
          .ItemData(.NewIndex) = giFILTEROP_ISLESSTHAN
          .AddItem OperatorDescription(giFILTEROP_ISATMOST)
          .ItemData(.NewIndex) = giFILTEROP_ISATMOST
          
        Case sqlDate ' Date columns.
          .AddItem OperatorDescription(giFILTEROP_ON)
          .ItemData(.NewIndex) = giFILTEROP_ON
          .AddItem OperatorDescription(giFILTEROP_NOTON)
          .ItemData(.NewIndex) = giFILTEROP_NOTON
          .AddItem OperatorDescription(giFILTEROP_AFTER)
          .ItemData(.NewIndex) = giFILTEROP_AFTER
          .AddItem OperatorDescription(giFILTEROP_BEFORE)
          .ItemData(.NewIndex) = giFILTEROP_BEFORE
          .AddItem OperatorDescription(giFILTEROP_ONORAFTER)
          .ItemData(.NewIndex) = giFILTEROP_ONORAFTER
          .AddItem OperatorDescription(giFILTEROP_ONORBEFORE)
          .ItemData(.NewIndex) = giFILTEROP_ONORBEFORE
        
        Case sqlVarChar, sqlLongVarChar, sqlVarBinary  ' Character and Photo columns (photo columns are really character columns).
          .AddItem OperatorDescription(giFILTEROP_IS)
          .ItemData(.NewIndex) = giFILTEROP_IS
          .AddItem OperatorDescription(giFILTEROP_ISNOT)
          .ItemData(.NewIndex) = giFILTEROP_ISNOT
          .AddItem OperatorDescription(giFILTEROP_CONTAINS)
          .ItemData(.NewIndex) = giFILTEROP_CONTAINS
          .AddItem OperatorDescription(giFILTEROP_DOESNOTCONTAIN)
          .ItemData(.NewIndex) = giFILTEROP_DOESNOTCONTAIN
      End Select
          
    End If
    
    .Enabled = (.ListCount > 0)
    txtStringValue.Enabled = .Enabled
    cboValues.Enabled = .Enabled
    
    If .ListCount > 0 Then
      .ListIndex = 0
    End If
  
  End With
  
  
  'MH20010118 Fault 1633
  'When the user entered a number which was too large caused
  'run time error. This is a simple way to stop it
  If iDataType = sqlNumeric Or iDataType = sqlInteger Then
    txtStringValue.MaxLength = 9
  Else
    txtStringValue.MaxLength = 0
  End If
  
  ' Refresh the data entry section
  RefreshValueEntryBox lngColumnID

End Sub

Private Sub cmdRemove_Click()
  ' Remove the selected row from the grid.
  Dim iRowIndex As Integer
  
  With ssDBGridFilterCriteria
    iRowIndex = .AddItemRowIndex(.Bookmark)
    
    ' RH 20/07/00 - Bug Fix 633. Wierd grid thing - doesnt delete the last row.
    If .Rows = 1 And iRowIndex = 0 Then
      .RemoveAll
    Else
      .RemoveItem iRowIndex
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
  DoGridSize

End Sub

Private Sub Form_Load()
Const GRIDROWHEIGHT = 239

ssDBGridFilterCriteria.RowHeight = GRIDROWHEIGHT

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    cmdCancel_Click
    Exit Sub
  End If

End Sub




Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub


Private Sub ssDBGridFilterCriteria_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
  ' Select the row.
  ssDBGridFilterCriteria.SelBookmarks.Add ssDBGridFilterCriteria.Bookmark

End Sub

Private Sub txtStringValue_GotFocus()
  ' Select all of the text in the textbox.
  UI.txtSelText
End Sub



Private Sub RefreshControls()
  ' Refresh the screen controls.
  ssDBGridFilterCriteria.Visible = (ssDBGridFilterCriteria.Rows > 0)
  txtNoCriteria.Visible = (ssDBGridFilterCriteria.Rows = 0)
  cmdRemove.Enabled = (ssDBGridFilterCriteria.Rows > 0)
  cmdClearAll.Enabled = (ssDBGridFilterCriteria.Rows > 0)
  
End Sub

Private Sub ssDBGridFilterCriteria_Populate(pavFilterArray As Variant)
  ' Populate the grid with the given filter criteria.
  Dim iLoop As Integer
  Dim objColumn As CColumnPrivilege
  
  With ssDBGridFilterCriteria
    .RemoveAll
    
    For iLoop = 1 To UBound(pavFilterArray, 2)
      ' Find the Filter Column name.
      For Each objColumn In mvar_objColumnPrivileges
        If objColumn.ColumnID = pavFilterArray(1, iLoop) Then
          .AddItem objColumn.ColumnName & vbTab & _
            OperatorDescription(CStr(pavFilterArray(2, iLoop))) & vbTab & _
            pavFilterArray(3, iLoop) & vbTab & _
            pavFilterArray(1, iLoop) & vbTab & _
            pavFilterArray(2, iLoop)
          Exit For
        End If
      Next objColumn
      Set objColumn = Nothing
    Next iLoop
    
    If .Rows > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With

End Sub

Private Function OperatorDescription(piOperatorCode As Integer) As String
  ' Return the textual description og the given operator.
  Dim sDesc As String
  
  Select Case piOperatorCode
    Case giFILTEROP_EQUALS
      'sDesc = "equals"
      sDesc = "is equal to"
    Case giFILTEROP_NOTEQUALTO
      'sDesc = "not equal to"
      sDesc = "is NOT equal to"
    Case giFILTEROP_ISATMOST
      'sDesc = "is at most"
      sDesc = "is less than or equal to"
    Case giFILTEROP_ISATLEAST
      'sDesc = "is at least"
      sDesc = "is greater than or equal to"
    Case giFILTEROP_ISMORETHAN
      'sDesc = "is more than"
      sDesc = "is greater than"
    Case giFILTEROP_ISLESSTHAN
      'sDesc = "is less than"
      sDesc = "is less than"
    Case giFILTEROP_ON
      'sDesc = "on"
      sDesc = "is equal to"
    Case giFILTEROP_NOTON
      'sDesc = "not on"
      sDesc = "is NOT equal to"
    Case giFILTEROP_AFTER
      sDesc = "after"
    Case giFILTEROP_BEFORE
      sDesc = "before"
    Case giFILTEROP_ONORAFTER
      'sDesc = "on or after"
      sDesc = "is equal to or after"
    Case giFILTEROP_ONORBEFORE
      'sDesc = "on or before"
      sDesc = "is equal to or before"
    Case giFILTEROP_CONTAINS
      sDesc = "contains"
    Case giFILTEROP_IS
      'sDesc = "is"
      sDesc = "is equal to"
    Case giFILTEROP_DOESNOTCONTAIN
      sDesc = "does not contain"
    Case giFILTEROP_ISNOT
      'sDesc = "is not"
      sDesc = "is NOT equal to"
    Case Else
      sDesc = ""
  End Select
  
  OperatorDescription = sDesc
  
End Function

Private Sub txtStringValue_LostFocus()

  Dim iDataType As Integer
  Dim fValidParameters As Boolean
  
  iDataType = mvar_objColumnPrivileges.Item(cboColumns.List(cboColumns.ListIndex)).DataType
    
  If iDataType = sqlDate Then
  
    ' Ensure that the value entered is a date.
    'AE20071123 Fault #7892
    'fValidParameters = (Len(txtStringValue.Text) = 0) Or _
    '  (IsDate(txtStringValue.Text))
    fValidParameters = ((Len(txtStringValue.Text) = 0) Or _
      (IsDate(txtStringValue.Text)) And (ValidateDate(txtStringValue.Text) <> vbNullString))
    
    fValidParameters = fValidParameters
    
    If Not fValidParameters Then
      txtStringValue.ForeColor = vbRed
      COAMsgBox "You have entered an invalid date.", vbExclamation, App.ProductName
      txtStringValue.ForeColor = vbWindowText
      txtStringValue.Text = ""
      txtStringValue.SetFocus
    End If
      
  End If
    
End Sub

' Display the correct data entry box
Private Sub RefreshValueEntryBox(tlngColumnID As Long)

  Dim iDataType As SQLDataType
  Dim strFormat As String
  Dim iDecimals As Integer
  Dim iSize As Integer
  Dim iCount As Integer
  
  iDataType = datGeneral.GetColumnDataType(tlngColumnID)

  Select Case iDataType
  
    Case sqlNumeric, sqlInteger
    
      strFormat = IIf(datGeneral.DoesColumnUseSeparators(tlngColumnID), "#0", "0")
      iDecimals = datGeneral.GetDecimalsSize(tlngColumnID)
      iSize = datGeneral.GetDataSize(tlngColumnID)
      
      strFormat = "0"
      For iCount = 2 To (iSize - iDecimals)
        If datGeneral.DoesColumnUseSeparators(tlngColumnID) = True Then
          strFormat = IIf(iCount Mod 3 = 0 And (iCount <> (iSize - iDecimals)), ",#", "#") & strFormat
        Else
          strFormat = "#" & strFormat
        End If
      Next iCount
  
      If iDecimals > 0 Then
        strFormat = strFormat & "."
        For iCount = 1 To iDecimals
          strFormat = strFormat & "0"
        Next iCount
      End If
      
      tdbNumberValue.DisplayFormat = strFormat
      tdbNumberValue.Format = strFormat
      tdbNumberValue.Text = 0
      tdbNumberValue.Visible = True
      cboValues.Visible = False
      txtStringValue.Visible = False
    
    Case sqlBoolean
      cboValues.ListIndex = 0
      cboValues.Visible = True
      tdbNumberValue.Visible = False
      txtStringValue.Visible = False
      
    Case Else
      txtStringValue.Text = ""
      txtStringValue.Visible = True
      cboValues.Visible = False
      tdbNumberValue.Visible = False
  
  End Select

End Sub

