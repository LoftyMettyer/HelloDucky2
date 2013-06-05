VERSION 5.00
Begin VB.Form frmDataTransferColumn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Transfer Columns"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1026
   Icon            =   "frmDataTransferColumn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   400
      Left            =   2100
      TabIndex        =   8
      Top             =   4600
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   400
      Left            =   3355
      TabIndex        =   9
      Top             =   4600
      Width           =   1200
   End
   Begin VB.Frame fraTo 
      Caption         =   "Destination :"
      Height          =   1260
      Left            =   150
      TabIndex        =   11
      Top             =   3235
      Width           =   4405
      Begin VB.ComboBox cboToColumn 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   700
         Width           =   2865
      End
      Begin VB.ComboBox cboToTable 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   300
         Width           =   2865
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   3
         Left            =   200
         TabIndex        =   15
         Top             =   760
         Width           =   630
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   2
         Left            =   200
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraFrom 
      Caption         =   "Source :"
      Height          =   3075
      Left            =   150
      TabIndex        =   10
      Top             =   100
      Width           =   4405
      Begin VB.OptionButton optText 
         Caption         =   "&Free Text"
         Height          =   315
         Left            =   200
         TabIndex        =   4
         Top             =   1635
         Width           =   1230
      End
      Begin VB.TextBox txtOther 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   5
         Top             =   2040
         Width           =   2865
      End
      Begin VB.OptionButton optSystemDate 
         Caption         =   "&System Date"
         Height          =   315
         Left            =   200
         TabIndex        =   3
         Top             =   2640
         Width           =   1530
      End
      Begin VB.OptionButton optTable 
         Caption         =   "Column &Value"
         Height          =   315
         Left            =   200
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.ComboBox cboFromColumn 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1100
         Width           =   2865
      End
      Begin VB.ComboBox cboFromTable 
         Height          =   315
         Left            =   1380
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   2865
      End
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text :"
         Height          =   195
         Left            =   500
         TabIndex        =   16
         Top             =   2100
         Width           =   435
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Column :"
         Height          =   195
         Index           =   1
         Left            =   495
         TabIndex        =   13
         Top             =   1155
         Width           =   765
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Table :"
         Height          =   195
         Index           =   0
         Left            =   500
         TabIndex        =   12
         Top             =   760
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmDataTransferColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mbCancelled As Boolean
Private datData As HRProDataMgr.clsDataAccess
Private mintDataType As Integer
Private mlngSize As Long
Private mintDecimals As Integer
Private miOLEType As HRProDataMgr.OLEType
Private mstrColumnsAlreadySelected As String
Private mfrmParent As Form
Private mlngFromTableID As Long
Private mstrFromTable As String
Private mlngToTableID As Long
Private mstrToTable As String


Public Property Get ParentForm() As Form
  Set ParentForm = mfrmParent
End Property

Public Property Let ParentForm(ByVal frmNewValue As Form)
  Set mfrmParent = frmNewValue
End Property


Public Sub Initialise(bNew As Boolean, lFromTableID As Long, lToTableID As Long, Optional sFromTable As String, _
        Optional sFromColumn As String, Optional sToTable As String, Optional sToColumn As String, _
        Optional sOtherText As String)
            
  Dim sSQL As String
  Dim rsChild As Recordset
  Dim sParent As String

  Set datData = New HRProDataMgr.clsDataAccess
  Call CheckWhichColumnsAreAlreadyUsed(bNew)
  
  
  sParent = ""
  Set rsChild = GetTableDetails(lFromTableID, sParent)
  mlngFromTableID = lFromTableID
  mstrFromTable = sParent


  With cboFromTable
    .Clear
    .AddItem sParent
    .ItemData(.NewIndex) = lFromTableID
    Do While Not rsChild.EOF
      .AddItem rsChild!TableName
      .ItemData(.NewIndex) = rsChild!TableID
      rsChild.MoveNext
    Loop
    If .ListCount > 0 Then
      If bNew Then
        SetComboText cboFromTable, sParent
      Else
        If Len(sFromTable) > 0 Then
          SetComboText cboFromTable, sFromTable
        Else
          .ListIndex = 0
        End If
      End If
    End If
  End With


  sParent = ""
  Set rsChild = GetTableDetails(lToTableID, sParent)
  mlngToTableID = lToTableID
  mstrToTable = sParent

  With cboToTable
    .Clear
    .AddItem sParent
    .ItemData(.NewIndex) = lToTableID
    Do While Not rsChild.EOF
      .AddItem rsChild!TableName
      .ItemData(.NewIndex) = rsChild!TableID
      rsChild.MoveNext
    Loop
    If .ListCount > 0 Then
      If bNew Then
        SetComboText cboToTable, sParent
      Else
        SetComboText cboToTable, sToTable
      End If
    End If
  End With
    
  If Not bNew Then
    'If Len(sOtherText) = 0 Then
    If Len(sFromColumn) > 0 Then
      'cboFromColumn.Text = sFromColumn
      SetComboText cboFromColumn, sFromColumn
    Else
      If sOtherText = "System Date" Then
        optSystemDate.Value = True
      Else
        optText.Value = True
        txtOther = sOtherText
      End If
    End If
    'cboToColumn.Text = sToColumn
    SetComboText cboToColumn, sToColumn
  End If
    
  rsChild.Close
  Set rsChild = Nothing
    
  Screen.MousePointer = vbDefault
            
End Sub

Private Sub cboFromColumn_Click()
  
  Dim rsTemp As Recordset
  Dim sSQL As String
  
  sSQL = "SELECT dataType, size, decimals, OLEType" & _
    " FROM ASRSysColumns" & _
    " WHERE columnID = " & CStr(cboFromColumn.ItemData(cboFromColumn.ListIndex))
  Set rsTemp = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  mintDataType = Val(rsTemp.Fields("DataType").Value)
  miOLEType = rsTemp.Fields("OLEType").Value
  
  If (mintDataType = sqlNumeric) Or _
    (mintDataType = sqlVarChar) Or _
    (mintDataType = sqlLongVarChar) Then
    mlngSize = rsTemp!Size
    mintDecimals = rsTemp!Decimals
  End If
  
  If cboToTable.ListCount > 0 Then
    ComboClick cboToTable, cboToColumn, True, cboFromColumn.Text
  End If
  
  Set rsTemp = Nothing

End Sub

Private Sub cboFromTable_Click()
  CheckToTables True
  ComboClick cboFromTable, cboFromColumn, False
  ComboClick cboToTable, cboToColumn, True, cboFromColumn.Text
End Sub

Private Sub ComboClick(cboTable As ComboBox, cboColumn As ComboBox, blnDestinationColumn As Boolean, _
                       Optional strDefault As String = vbNullString)

  
  'Columns will only appear in the destination column combo box
  'if they meet ALL of the following conditions:
  '
  ' 1. The column is on the table selected in the destination table combo.
  ' 2. The column matches the data type selected
  '    (also must be a date column if system date is selected
  ' 3. The destination column is the same size or bigger than the source column.
  '    (if free text then check the length of the free text)
  ' 4. The destination column has the same or more decimals than the source column.
  ' 5. The column is not already selected in the data transfer definition
  '    (unless you are currently editting a column then that column will be in
  '     the combo)


  Dim sSQL As String
  Dim rsColumns As Recordset
  Dim lngDefaultItem As Long
  Dim strSearch As String
  Dim lngMaxLength As Long


  cboColumn.Clear
  If cboTable.ListIndex <> -1 Then

    sSQL = "SELECT ColumnName, ColumnID, Size FROM ASRSysColumns " & _
           "WHERE TableID = " & cboTable.ItemData(cboTable.ListIndex) & _
           " AND ColumnType <> 3 AND ColumnType <> 4 "
         
    'If datatype is not zero then we are checking
    'the destination columns available
    If blnDestinationColumn Then
    
      If mstrColumnsAlreadySelected <> vbNullString Then
        sSQL = sSQL & " AND ColumnID NOT IN (" & _
                      mstrColumnsAlreadySelected & ")"
      End If


      'MH20071016 Fault 12506 (Moved code up outside IF block)
      'If numeric only show columns which are longer
      '(or the same size) as the selected source column
      If (mintDataType = sqlNumeric) Or _
        (mintDataType = sqlVarChar) Or _
        (mintDataType = sqlLongVarChar) Or _
        (mintDataType = 0) Then
        sSQL = sSQL & " AND Size >= " & CStr(mlngSize) & _
                      " AND Decimals >= " & CStr(mintDecimals)
      End If


      If mintDataType <> 0 Then
    
        sSQL = sSQL & " AND ReadOnly = 0" & _
                      " AND DataType = " & CStr(mintDataType)
    
        ' Compatible OLE types
        If mintDataType = sqlVarBinary Or mintDataType = sqlOle Then
          sSQL = sSQL & " AND OLEType = " & Str(miOLEType)
        End If
      
      Else
        ' Block embedded/linked OLE and photo types
        sSQL = sSQL & " AND OLEType < 2"
      End If
      
    End If
  
    Set rsColumns = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  
  
    lngDefaultItem = 0
    lngMaxLength = 1
    strSearch = LCase$(Replace(strDefault, "_", ""))
    With cboColumn
    
      Do While Not rsColumns.EOF
        .AddItem rsColumns!ColumnName
        .ItemData(.NewIndex) = rsColumns!ColumnID
      
        If lngDefaultItem = 0 Then
          If LCase$(Replace(rsColumns!ColumnName, "_", "")) = strSearch Then
            lngDefaultItem = .ItemData(.NewIndex)
          End If
        End If

        If rsColumns!Size > lngMaxLength Then
          lngMaxLength = rsColumns!Size
        End If

        rsColumns.MoveNext
      Loop

      If lngMaxLength > 255 Then lngMaxLength = 255
      txtOther.MaxLength = lngMaxLength

    End With
    
    rsColumns.Close
    Set rsColumns = Nothing
    
  End If
  

  With cboColumn
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    .Enabled = IIf(.ListCount > 0, True, False)
    If lngDefaultItem > 0 Then
      SetComboItem cboColumn, lngDefaultItem
    Else
      If .ListCount > 0 Then
        .ListIndex = 0
      End If
    End If
  End With

End Sub

Private Sub cboToTable_Click()
  ComboClick cboToTable, cboToColumn, True, cboFromColumn.Text
End Sub

Private Sub cmdCancel_Click()
  Cancelled = True
  Unload Me
End Sub

Public Property Get Cancelled() As Boolean
  Cancelled = mbCancelled
End Property

Public Property Let Cancelled(ByVal bCancel As Boolean)
  mbCancelled = bCancel
End Property

Private Sub cmdOK_Click()
  
  If cboToColumn.ListIndex < 0 Then
    MsgBox "No destination column selected", vbExclamation
    Exit Sub
  End If
  
  If ValidColumnTransfer = False Then
    Exit Sub
  End If
  
  Cancelled = False
  Me.Hide
End Sub

Private Sub CheckWhichColumnsAreAlreadyUsed(blnNew As Boolean)

  Dim lngRow As Long
  Dim pvarbookmark As Variant
  Dim lngColumnID As Long
  Dim lngColumnCurrentlyEditting As Long

  mstrColumnsAlreadySelected = vbNullString
  
  With mfrmParent.grdColumns
    
    If blnNew = False Then
      'Store the ID of the row which you are currently
      'editting and do not include this from the list
      'of columns already selected
      lngColumnCurrentlyEditting = Val(mfrmParent.grdColumns.Columns(8).Text)
    Else
      lngColumnCurrentlyEditting = 0
    End If
    
    'MH20001109 Fault 1331
    '.Row = 0
    .MoveFirst
    For lngRow = 0 To .Rows - 1
      pvarbookmark = .GetBookmark(lngRow)
      lngColumnID = Val(.Columns(8).CellText(pvarbookmark))
      
      'If columnID is not column currently editting
      'and columnID is greater than zero then don't
      'allow the user to select the column again
      If lngColumnID > 0 And lngColumnID <> lngColumnCurrentlyEditting Then
        mstrColumnsAlreadySelected = mstrColumnsAlreadySelected & _
          IIf(mstrColumnsAlreadySelected <> "", ", ", "") & CStr(lngColumnID)
      End If
    Next
  End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 192 Then
    KeyCode = 0
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode <> vbFormCode Then
    Cancelled = True
  End If
End Sub

Private Sub Form_Resize()
  'JPD 20030908 Fault 5756
  DisplayApplication

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set datData = Nothing
End Sub

Private Sub optSystemDate_Click()

  CheckToTables False
  mintDataType = sqlDate
  ComboClick cboToTable, cboToColumn, True

  cboFromTable.Enabled = False
  cboFromTable.BackColor = vbButtonFace
  cboFromColumn.Enabled = False
  cboFromColumn.BackColor = vbButtonFace
  lblTitle(0).Enabled = False
  lblTitle(1).Enabled = False
  txtOther.Enabled = False
  txtOther.Text = ""
  txtOther.BackColor = vbButtonFace
  lblOther.Enabled = False

End Sub

Private Sub optTable_Click()

  Call cboFromColumn_Click
  
  txtOther.Enabled = False
  txtOther.Text = ""
  txtOther.BackColor = vbButtonFace
  lblOther.Enabled = False
  cboFromTable.Enabled = True
  cboFromTable.BackColor = vbWindowBackground
  cboFromColumn.Enabled = True
  cboFromColumn.BackColor = vbWindowBackground
  lblTitle(0).Enabled = True
  lblTitle(1).Enabled = True

End Sub

Private Sub optText_Click()

  CheckToTables False
  mintDataType = 0   'Show all !
  mlngSize = Len(txtOther)
  mintDecimals = 0
  ComboClick cboToTable, cboToColumn, True

  cboFromTable.Enabled = False
  cboFromTable.BackColor = vbButtonFace
  cboFromColumn.Enabled = False
  cboFromColumn.BackColor = vbButtonFace
  lblTitle(0).Enabled = False
  lblTitle(1).Enabled = False
  txtOther.Enabled = True
  lblOther.Enabled = True
  txtOther.BackColor = vbWindowBackground
  If Me.Visible Then
    txtOther.SetFocus
  End If

End Sub


Private Function GetTableDetails(lTableID As Long, sTableName As String) As ADODB.Recordset

  Dim sSQL As String
  Dim rsTable As Recordset
    
  sSQL = "Select TableName From ASRSysTables Where TableID = " & lTableID
  Set rsTable = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  sTableName = rsTable(0)
  rsTable.Close
    
  'Get all the child tables related to the parent (or child)
  sSQL = "SELECT ASRSysTables.TableName, ASRSysTables.TableID FROM ASRSysTables INNER JOIN " & _
         "ASRSysRelations ON ASRSysTables.TableID = ASRSysRelations.ChildID Where " & _
         "ASRSysRelations.ParentID = " & lTableID
  Set rsTable = datData.OpenRecordset(sSQL, adOpenForwardOnly, adLockReadOnly)
  Set GetTableDetails = rsTable
    
End Function


Private Function ValidColumnTransfer() As Boolean

  Dim iFromColumnSize As Integer
  Dim iToColumnSize As Integer
  Dim iDataType As SQLDataType

  ValidColumnTransfer = True
  
  If optTable.Value = True Then
    iDataType = datGeneral.GetDataType(cboToTable.ItemData(cboToTable.ListIndex), cboToColumn.ItemData(cboToColumn.ListIndex))
    If iDataType = sqlOle Or iDataType = sqlVarBinary Then
      iFromColumnSize = datGeneral.GetOLEMaxSize(cboFromColumn.ItemData(cboFromColumn.ListIndex))
      iToColumnSize = datGeneral.GetOLEMaxSize(cboToColumn.ItemData(cboToColumn.ListIndex))
      
      If iToColumnSize >= 0 And (iFromColumnSize > iToColumnSize Or iFromColumnSize = -1) Then
        ValidColumnTransfer = IIf(MsgBox("The source column is potentially larger than the destination column, which may result in data being read only." _
            & vbCrLf & "Are you sure you wish to continue?", vbQuestion & vbYesNo, Me.Caption) = vbYes, True, False)
      End If
      
    End If
  End If
  
  
  If optText.Value = True Then
    Select Case datGeneral.GetDataType(cboToTable.ItemData(cboToTable.ListIndex), cboToColumn.ItemData(cboToColumn.ListIndex))
    Case sqlDate
      If Not IsDate(txtOther) Then
        ValidColumnTransfer = False
      End If
    Case sqlNumeric, sqlInteger
      If Val(txtOther) = 0 And Trim$(txtOther) <> "0" Then
        ValidColumnTransfer = False
      End If
    Case sqlBoolean
      If txtOther <> Chr(34) & "0" & Chr(34) Then
        If txtOther <> Chr(34) & "1" & Chr(34) Then
          ValidColumnTransfer = False
        End If
      End If
    End Select
    
    If Not ValidColumnTransfer Then
      MsgBox "The free text does not match the destination column data type", vbExclamation
    End If
    
  End If

End Function

Private Sub txtOther_Change()
 
  cboToColumn.Visible = False
  'mintDataType = sqlVarChar
  mlngSize = Len(txtOther.Text)
  'mintDecimals = 0
  ComboClick cboToTable, cboToColumn, True, cboToColumn.Text
  cboToColumn.Visible = True

End Sub


Private Sub CheckToTables(blnCheckFromTable As Boolean)

  Dim intCount As Integer
  Dim intFound As Integer
  
  With cboToTable
    
    'Determine if the main destination table is in the combo box
    intFound = -1
    For intCount = 0 To .ListCount - 1
      If .ItemData(intCount) = mlngToTableID Then
        intFound = intCount
        Exit For
      End If
    Next

    If cboFromTable.ItemData(cboFromTable.ListIndex) <> mlngFromTableID And blnCheckFromTable Then
      'If the user has selected a child of the source then
      'ensure that the main destination table is not in the destination combo
      If intFound <> -1 Then
        .RemoveItem intFound
      End If
    Else
      'If the main source table has been selected then
      'ensure that the main destination table is in the combo
      If intFound = -1 Then
        .AddItem mstrToTable
        .ItemData(.NewIndex) = mlngToTableID
      End If
    End If
        
    If .ListCount > 0 And .ListIndex = -1 Then
      .ListIndex = 0
    End If
    .BackColor = IIf(.ListCount > 0, vbWindowBackground, vbButtonFace)
    .Enabled = IIf(.ListCount > 0, True, False)

  End With

End Sub

