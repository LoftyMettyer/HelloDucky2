VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmPicklistFind 
   Caption         =   "Add Items to Picklist"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1051
   Icon            =   "frmPicklistFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   100
      TabIndex        =   6
      Top             =   3100
      Width           =   2600
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1395
         TabIndex        =   4
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Frame fraOrders 
      Caption         =   "Options :"
      Height          =   1200
      Left            =   100
      TabIndex        =   5
      Top             =   100
      Width           =   6000
      Begin VB.ComboBox cmbView 
         Height          =   315
         Left            =   1035
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   4815
      End
      Begin VB.ComboBox cmbOrders 
         Height          =   315
         Left            =   1035
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   700
         Width           =   4815
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "View :"
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblOrder 
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   255
         Left            =   195
         TabIndex        =   7
         Top             =   765
         Width           =   780
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssOleDBGridFindColumns 
      Height          =   1410
      Left            =   100
      TabIndex        =   2
      Top             =   1500
      Width           =   6000
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
      GroupHeadLines  =   0
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
      _ExtentX        =   10583
      _ExtentY        =   2487
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
Attribute VB_Name = "frmPicklistFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Table variables.
Private mobjTable As CTablePrivilege

' Record display option variables.
Private mlngOrderID As Long
Private mlngViewID As Long

' Recordset variables.
Private mrsFindRecords As ADODB.Recordset
Private mlngRecordCount As Long

' Form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer

Private mstrExistingIDs As String

Private mavFindColumns() As Variant        ' Find columns details

Private Const dblFINDFORM_MINWIDTH = 5000
Private Const dblFINDFORM_MINHEIGHT = 5000

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled
End Property

Private Sub cmbOrders_Click()
  Dim fOK As Boolean
  
  ' Do nothing if the form is not visible.
  fOK = Me.Visible

  If fOK Then
    Screen.MousePointer = vbHourglass
  End If

  ' Do nothing if there are no records.
  If fOK Then
    fOK = (mlngRecordCount > 0)
  End If
    
  ' Do nothing if there are no orders defined.
  If fOK Then
    fOK = (cmbOrders.ItemData(cmbOrders.ListIndex) > 0)
  End If

  If fOK Then
    ' Set the order ID variable.
    mlngOrderID = cmbOrders.ItemData(cmbOrders.ListIndex)
    
    ' Refresh the recordset for the new order.
    GetRecords
  
    ' RH 06/09/00 - If not cancelled (because user cant see any records on the selected order)
    If mfCancelled <> True Then
      With ssOleDBGridFindColumns
        cmdSelect.Enabled = (.Rows > 0)
      
        If .Rows > 0 Then
          .MoveFirst
          .SelBookmarks.Add (.Bookmark)
          .SetFocus
        End If
      End With
    End If
  End If
  
  Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmbView_Click()
  Dim fOK As Boolean
  
  ' Do nothing if the form is not visible.
  fOK = Me.Visible
  
  If fOK Then
    Screen.MousePointer = vbHourglass
  End If

  If fOK Then
    ' Set the view ID variable.
    If mobjTable.AllowSelect And _
      (cmbView.ListIndex = 0) Then
      mlngViewID = 0
    Else
      mlngViewID = cmbView.ItemData(cmbView.ListIndex)
    End If
    
    GetRecords
    
    With ssOleDBGridFindColumns
      cmdSelect.Enabled = (.Rows > 0)
    
      If .Rows > 0 Then
        .MoveFirst
        .SelBookmarks.Add (.Bookmark)
        .SetFocus
      End If
    End With
  End If

  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide

End Sub

Private Sub cmdSelect_Click()
  ssOleDBGridFindColumns_DblClick

End Sub

Private Sub Form_Load()
  Hook Me.hWnd, dblFINDFORM_MINWIDTH, dblFINDFORM_MINHEIGHT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
    Cancel = True
    Me.Hide
  End If
    
End Sub

Private Sub Form_Resize()
  Dim lCount As Long
  Dim lWidth As Long
  Dim iLastColumnIndex As Integer
  Dim iMaxPosition As Integer
   
  Const dblCOORD_XGAP = 200
  Const dblCOORD_YGAP = 200
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  ' Ensure the form does not get narrower than the defined minimum for a Find window.
'  If Me.Width < dblFINDFORM_MINWIDTH Then
'    Me.Width = dblFINDFORM_MINWIDTH
'  End If
'
'  ' Ensure the form does not get wider than the screen.
'  If Me.Width > Screen.Width Then
'    Me.Width = Screen.Width
'  End If

  ' Initialise the form height.
  If Not mfSizing Then
    mfSizing = True
    Me.Height = Screen.Height / 3
  End If
            
  ' Ensure the form does not get shorter than the defined minimum for a Find window.
  If Me.Height < dblFINDFORM_MINHEIGHT Then
    mfSizing = True
    Me.Height = dblFINDFORM_MINHEIGHT
  End If
            
  ' Ensure the form does not get taller than the screen.
  If Me.Height > Screen.Height Then
    Me.Height = Screen.Height
  End If
        
  ' Size the Order frame and the controls therein.
  fraOrders.Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
  cmbOrders.Width = fraOrders.Width - (dblCOORD_XGAP * 6)
  cmbView.Width = fraOrders.Width - (dblCOORD_XGAP * 6)

  ' Size the Find grid.
  With ssOleDBGridFindColumns
    .Width = fraOrders.Width
    .Height = Me.ScaleHeight - .Top - fraButtons.Height - (2 * dblCOORD_YGAP)
  End With
        
  ' Size the frames with the command buttons in.
  With fraButtons
    .Top = ssOleDBGridFindColumns.Top + ssOleDBGridFindColumns.Height + dblCOORD_YGAP
    .Left = Me.ScaleWidth - fraButtons.Width - dblCOORD_XGAP
  End With

  If ((ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.FirstRow + 1) < (ssOleDBGridFindColumns.VisibleRows)) And _
    (ssOleDBGridFindColumns.FirstRow > 1) Then

    ssOleDBGridFindColumns.FirstRow = IIf(ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.VisibleRows + 1 < 1, _
      1, _
      ssOleDBGridFindColumns.Rows - ssOleDBGridFindColumns.VisibleRows + 1)
  End If

  ' Stretch the last find column to fit the grid.
  iLastColumnIndex = -1
  iMaxPosition = -1
  With ssOleDBGridFindColumns
    For lCount = 0 To (.Cols - 1)
      If .Columns(lCount).Visible Then
        lWidth = lWidth + .Columns(lCount).Width
        If .Columns(lCount).Position > iMaxPosition Then
          iMaxPosition = .Columns(lCount).Position
          iLastColumnIndex = lCount
        End If
      End If
    Next lCount
    
    If (lWidth < .Width) And _
      (iLastColumnIndex >= 0) Then
      .Columns(iLastColumnIndex).Width = .Columns(iLastColumnIndex).Width + (.Width - lWidth)
    End If
  End With
    
  ' Get rid of the icon off the form
  RemoveIcon Me
    
End Sub





Private Sub GetRecords()
  ' Read the required information about the selected table.
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim fNoSelect As Boolean
  Dim iNextIndex As Integer
  Dim lngFirstFindColumnID As Long
  Dim lngFirstSortColumnID As Long
  Dim sSQL As String
  Dim sSource As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sOrderString As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim objTableView As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
  
  fNoSelect = False
  
  sOrderString = ""
  sJoinCode = ""
  sColumnList = ""

  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)
  
  mfFirstColumnsMatch = False
  lngFirstFindColumnID = 0
  lngFirstSortColumnID = 0
  mfFirstColumnAscending = True
  miFirstColumnDataType = 0
  
  ReDim mavFindColumns(3, 0)
  
  ' Get the default order items from the database.
  Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)
  
  If rsInfo.EOF And rsInfo.BOF Then
    MsgBox "No order defined for this " & IIf(mlngViewID > 0, "view.", "table.") & _
      vbCrLf & "Unable to display records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      ' Get the column privileges collection for the given table.
      If rsInfo!TableID = mobjTable.TableID Then
        If mlngViewID = 0 Then
          sSource = mobjTable.TableName
        Else
          sSource = GetViewName(mlngViewID)
        End If
      Else
        sSource = rsInfo!TableName
      End If
      Set objColumnPrivileges = GetColumnPrivileges(sSource)
      sRealSource = gcoTablePrivileges.Item(sSource).RealSource
    
      fColumnOK = objColumnPrivileges.IsValid(rsInfo!ColumnName)
      
      If fColumnOK Then
        fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect
      End If
      Set objColumnPrivileges = Nothing
    
      If fColumnOK Then
        ' The column can be read from the base table/view, or directly from a parent table.
        If rsInfo!Type = "F" Then
          ' Add the column to the column list.
          sColumnList = sColumnList & _
            IIf(Len(sColumnList) > 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName)
          
          mavFindColumns(0, UBound(mavFindColumns, 2)) = rsInfo!ColumnID
          mavFindColumns(1, UBound(mavFindColumns, 2)) = datGeneral.GetDataSize(rsInfo!ColumnID)
          mavFindColumns(2, UBound(mavFindColumns, 2)) = datGeneral.GetDecimalsSize(rsInfo!ColumnID)
          mavFindColumns(3, UBound(mavFindColumns, 2)) = datGeneral.DoesColumnUseSeparators(rsInfo!ColumnID)
          ReDim Preserve mavFindColumns(3, UBound(mavFindColumns, 2) + 1)
          
          ' Remember the first Find column.
          If lngFirstFindColumnID = 0 Then
            lngFirstFindColumnID = rsInfo!ColumnID
          End If
        Else
          ' Add the column to the order string.
          sOrderString = sOrderString & _
            IIf(Len(sOrderString) > 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName) & _
            IIf(rsInfo!Ascending, "", " DESC")
          
          ' Remember the first Order column.
          If lngFirstSortColumnID = 0 Then
            lngFirstSortColumnID = rsInfo!ColumnID
            mfFirstColumnAscending = rsInfo!Ascending
            miFirstColumnDataType = rsInfo!DataType
          End If
        End If
    
        ' If the column comes from a parent table, then add the table to the Join code.
        If rsInfo!TableID <> mobjTable.TableID Then
          ' Check if the table has already been added to the join code.
          fFound = False
          For iNextIndex = 1 To UBound(alngTableViews, 2)
            If alngTableViews(1, iNextIndex) = 0 And _
              alngTableViews(2, iNextIndex) = rsInfo!TableID Then
              fFound = True
              Exit For
            End If
          Next iNextIndex
          
          If Not fFound Then
            ' The table has not yet been added to the join code, so add it to the array and the join code.
            iNextIndex = UBound(alngTableViews, 2) + 1
            ReDim Preserve alngTableViews(2, iNextIndex)
            alngTableViews(1, iNextIndex) = 0
            alngTableViews(2, iNextIndex) = rsInfo!TableID
            
            sJoinCode = sJoinCode & _
              " LEFT OUTER JOIN " & sRealSource & _
              " ON " & CurrentTableViewName & ".ID_" & Trim(Str(rsInfo!TableID)) & _
              " = " & sRealSource & ".ID"
          End If
        End If
      Else
        ' The column cannot be read from the base table/view, or directly from a parent table.
        ' If it is a column from a prent table, then try to read it from the views on the parent table.
        If rsInfo!TableID <> mobjTable.TableID Then
          ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
          ReDim asViews(0)
          For Each objTableView In gcoTablePrivileges.Collection
            If (Not objTableView.IsTable) And _
              (objTableView.TableID = rsInfo!TableID) And _
              (objTableView.AllowSelect) Then
              
              sSource = objTableView.ViewName
              sRealSource = gcoTablePrivileges.Item(sSource).RealSource

              ' Get the column permission for the view.
              Set objColumnPrivileges = GetColumnPrivileges(sSource)

              If objColumnPrivileges.IsValid(rsInfo!ColumnName) Then
                If objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect Then
                  ' Add the view info to an array to be put into the column list or order code below.
                  iNextIndex = UBound(asViews) + 1
                  ReDim Preserve asViews(iNextIndex)
                  asViews(iNextIndex) = objTableView.ViewName
                  
                  ' Add the view to the Join code.
                  ' Check if the view has already been added to the join code.
                  fFound = False
                  For iNextIndex = 1 To UBound(alngTableViews, 2)
                    If alngTableViews(1, iNextIndex) = 1 And _
                      alngTableViews(2, iNextIndex) = objTableView.ViewID Then
                      fFound = True
                      Exit For
                    End If
                  Next iNextIndex
          
                  If Not fFound Then
                    ' The view has not yet been added to the join code, so add it to the array and the join code.
                    iNextIndex = UBound(alngTableViews, 2) + 1
                    ReDim Preserve alngTableViews(2, iNextIndex)
                    alngTableViews(1, iNextIndex) = 1
                    alngTableViews(2, iNextIndex) = objTableView.ViewID
          
                    sJoinCode = sJoinCode & _
                      " LEFT OUTER JOIN " & sRealSource & _
                      " ON " & CurrentTableViewName & ".ID_" & Trim(Str(objTableView.TableID)) & _
                      " = " & sRealSource & ".ID"
                  End If
                End If
              End If
              Set objColumnPrivileges = Nothing

            End If
          Next objTableView
          Set objTableView = Nothing
    
          ' The current user does have permission to 'read' the column through a/some view(s) on the
          ' table.
          If UBound(asViews) = 0 Then
            fNoSelect = True
          Else
            ' Add the column to the column list.
            sColumnCode = ""
            For iNextIndex = 1 To UBound(asViews)
              If iNextIndex = 1 Then
                sColumnCode = "CASE "
              End If
                
              sColumnCode = sColumnCode & _
                " WHEN NOT " & asViews(iNextIndex) & "." & rsInfo!ColumnName & " IS NULL THEN " & asViews(iNextIndex) & "." & rsInfo!ColumnName
            Next iNextIndex
              
            If Len(sColumnCode) > 0 Then
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END AS " & _
                IIf(rsInfo!Type = "F", "", "'?") & _
                rsInfo!ColumnName & _
                IIf(rsInfo!Type = "F", "", "'")
                
              sColumnList = sColumnList & _
                IIf(Len(sColumnList) > 0, ", ", "") & _
                sColumnCode

              If rsInfo!Type = "F" Then
                ' Remember the first Find column.
                If lngFirstFindColumnID = 0 Then
                  lngFirstFindColumnID = rsInfo!ColumnID
                End If
              
                mavFindColumns(0, UBound(mavFindColumns, 2)) = rsInfo!ColumnID
                mavFindColumns(1, UBound(mavFindColumns, 2)) = datGeneral.GetDataSize(rsInfo!ColumnID)
                mavFindColumns(2, UBound(mavFindColumns, 2)) = datGeneral.GetDecimalsSize(rsInfo!ColumnID)
                mavFindColumns(3, UBound(mavFindColumns, 2)) = datGeneral.DoesColumnUseSeparators(rsInfo!ColumnID)
                ReDim Preserve mavFindColumns(3, UBound(mavFindColumns, 2) + 1)
              Else
                ' Add the column to the order string.
                sOrderString = sOrderString & _
                  IIf(Len(sOrderString) > 0, ", ", "") & _
                  "'?" & Trim(rsInfo!ColumnName) & "'" & _
                  IIf(rsInfo!Ascending, "", " DESC")

                ' Remember the first Order column.
                If lngFirstSortColumnID = 0 Then
                  lngFirstSortColumnID = rsInfo!ColumnID
                  mfFirstColumnAscending = rsInfo!Ascending
                  miFirstColumnDataType = rsInfo!DataType
                End If
              End If
            End If
          End If
        End If
      End If
      
      rsInfo.MoveNext
    Loop
    
    ' Inform the user if they do not have permission to see the data.
    If fNoSelect Then
      MsgBox "You do not have 'read' permission on all of the columns in the selected order." & _
        vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
    End If
    
    mfFirstColumnsMatch = (lngFirstFindColumnID = lngFirstSortColumnID)
  
    ' Create the string for creating the items that will appear in the listbox.
    If Len(sColumnList) > 0 Then
      
      ' RH 06/09 Only show records that are not already in the picklist
      
      'sSQL = "SELECT " & sColumnList & ", " & CurrentTableViewName & ".id" & _
        " FROM " & CurrentTableViewName & _
        " " & sJoinCode & _
        IIf(Len(sOrderString) > 0, " ORDER BY " & sOrderString, "")
    
      sSQL = "SELECT " & sColumnList & ", " & CurrentTableViewName & ".id" & _
        " FROM " & CurrentTableViewName & _
        " " & sJoinCode & _
        IIf(Len(mstrExistingIDs) > 0, " WHERE " & CurrentTableViewName & ".ID NOT IN (" & mstrExistingIDs & ") ", "") & _
        IIf(Len(sOrderString) > 0, " ORDER BY " & sOrderString, "")
    
      ' Get the required recordset.
      
      ' RH 06/09/00 - Change recordset from keyset to static readonly
      'Set mrsFindRecords = datGeneral.GetMainRecordset(sSQL)
      Set mrsFindRecords = datGeneral.GetReadOnlyRecords(sSQL)
  
     
      ' Get the recordset's record count.
      ' RH 06/09/00 - Cant use RecordCount function anymore because we are excluding the
      '               records that are already in the picklist and this function doesnt
      '               take this into account. lets use the .recordcount property instead.
      mlngRecordCount = RecordCount

      ' Configure the grid.
      ConfigureGrid
    Else
      MsgBox "You do not have permission to read any of the columns in the selected order for this " & IIf(mlngViewID > 0, "view.", "table.") & _
          vbCrLf & "Unable to display records.", vbExclamation, "Security"
      'TM20061031 - Fault 11657
      'There may be no records in the default view, but there may be in other views.
      ' So don't quit out here.
      'mfCancelled = True
      'Me.Hide
    End If
  End If
  
  rsInfo.Close
  Set rsInfo = Nothing

End Sub


Public Function Initialise(plngTableID As Long, Optional pstrExistingIDs As String) As Boolean
  ' Initialise the Picklist find form.

  ' RH 06/09/00 - Set module level variable
  mstrExistingIDs = pstrExistingIDs
  
  ' Get the table object.
  Set mobjTable = gcoTablePrivileges.FindTableID(plngTableID)
  
  ' Get the table's default order ID.
  mlngOrderID = mobjTable.DefaultOrderID
  
  ' Populate the View combo.
  ConfigureViewCombo
  If cmbView.ListCount = 0 Then
    MsgBox "You do not have 'read' permission on this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    'mfCancelled = True
    'Me.Hide
    Initialise = False
    Exit Function
  End If
  
  If mobjTable.AllowSelect Then
    mlngViewID = 0
  Else
    mlngViewID = cmbView.ItemData(cmbView.ListIndex)
  End If
  
  ' Populate the Orders combo.
  ConfigureOrdersCombo
  If cmbOrders.ListCount = 0 Then
    MsgBox "There are no orders defined for this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    'mfCancelled = True
    'Me.Hide
    Initialise = False
    Screen.MousePointer = vbDefault
    Exit Function
  End If
  
  ' Get the find records.
  Set mrsFindRecords = New Recordset
  Screen.MousePointer = vbHourglass
  GetRecords
    
  'JPD 20030507 There may be no records in the default view, but there may be in other views.
  ' So don't quit out here.
'  If mrsFindRecords.BOF And mrsFindRecords.EOF Then
'    MsgBox "There are no records which do not already exist in this picklist.", vbExclamation + vbOKOnly, "Picklists"
'    Initialise = False
'    Screen.MousePointer = vbDefault
'    Exit Function
'  End If
  
  If ssOleDBGridFindColumns.Rows > 0 Then
    ssOleDBGridFindColumns.MoveFirst
    ssOleDBGridFindColumns.SelBookmarks.Add (ssOleDBGridFindColumns.Bookmark)
  End If
  
  cmdSelect.Enabled = (ssOleDBGridFindColumns.Rows <> 0)
  
  Screen.MousePointer = vbDefault
  Initialise = True
  
End Function
Private Sub ConfigureViewCombo()
  ' Populate the 'views' combo with the views available for the table.
  Dim objTableView As CTablePrivilege
  
'  ' Add the table to the combo if the user has permission to read it.
'  If mobjTable.AllowSelect Then
'    cmbView.AddItem RemoveUnderScores(mobjTable.TableName)
'    cmbView.ItemData(cmbView.NewIndex) = mobjTable.TableID
'  End If
  
  ' Add the table's views to the combo if the user has permission to read them.
  For Each objTableView In gcoTablePrivileges.Collection
    If (Not objTableView.IsTable) And _
      (objTableView.TableID = mobjTable.TableID) Then
      
      If objTableView.AllowSelect Then
        cmbView.AddItem "'" & RemoveUnderScores(Trim(objTableView.ViewName)) & "' view"
        cmbView.ItemData(cmbView.NewIndex) = objTableView.ViewID
      End If
    End If
  Next objTableView
  Set objTableView = Nothing
  
    'NHRD20060213 Fault 10516 Moved code from above and forced this to be the first
    ' item in the dropdown with the Optional index parameter in Additem
    ' Add the table to the combo if the user has permission to read it.
    If mobjTable.AllowSelect Then
      cmbView.AddItem RemoveUnderScores(mobjTable.TableName), 0
      cmbView.ItemData(cmbView.NewIndex) = mobjTable.TableID
    End If
  
  With cmbView
    ' Select the first view.
    If .ListCount > 0 Then
      .ListIndex = 0
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

End Sub

Private Sub ConfigureGrid()
  ' Configure the grid to display the required columns.
  Dim iLoop As Integer
  Dim lngWidth As Long

  UI.LockWindow Me.hWnd
  
  lngWidth = 0
  mfFormattingGrid = True
  
  With ssOleDBGridFindColumns
        
    .Redraw = False
    .Columns.RemoveAll

    For iLoop = 0 To (mrsFindRecords.Fields.Count - 1)
      .Columns.Add iLoop
      .Columns(iLoop).Name = mrsFindRecords.Fields(iLoop).Name
      .Columns(iLoop).Visible = (UCase(mrsFindRecords.Fields(iLoop).Name) <> "ID") And _
        (Left(mrsFindRecords.Fields(iLoop).Name, 1) <> "?")
      .Columns(iLoop).Caption = RemoveUnderScores(mrsFindRecords.Fields(iLoop).Name)
      .Columns(iLoop).Alignment = ssCaptionAlignmentLeft
      .Columns(iLoop).CaptionAlignment = ssColCapAlignUseColumnAlignment
    
      ' If the find column is a logic column then set the grid column style to be 'checkbox'.
      If mrsFindRecords.Fields.Item(iLoop).Type = adBoolean Then
        .Columns(iLoop).Style = ssStyleCheckBox
      End If

      ' Total the size of the grid columns
      If .Columns(iLoop).Visible Then
        lngWidth = lngWidth + .Columns(iLoop).Width
      End If
    Next iLoop

    mfFormattingGrid = False
    .Rebind
    .Rows = mlngRecordCount
    .Redraw = True
  End With

  ' Adjust the size of the window to fit the grid.
  lngWidth = lngWidth + (fraOrders.Left * 2) + _
    (((UI.GetSystemMetrics(SM_CXFRAME) * 2) + _
    UI.GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX)

  If ssOleDBGridFindColumns.Rows > ssOleDBGridFindColumns.VisibleRows Then
    lngWidth = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20
  End If

  Me.Width = lngWidth + 120

  UI.UnlockWindow
  
End Sub


Private Function RecordCount() As Long
  ' Return the number of records in the recordset.
  Dim rsTemp As Recordset

  Set rsTemp = datGeneral.GetRecords("SELECT COUNT(id) FROM " & CurrentTableViewName & _
  IIf(Len(mstrExistingIDs) > 0, " WHERE " & CurrentTableViewName & ".ID NOT IN (" & mstrExistingIDs & ") ", ""))

  If (rsTemp.EOF And rsTemp.BOF) Then
    RecordCount = 0
  Else
    RecordCount = rsTemp(0)
  End If
  rsTemp.Close
  Set rsTemp = Nothing
  
End Function

Private Function CurrentTableViewName() As String
  ' Return the name of the table, or view.
  CurrentTableViewName = IIf(mlngViewID = 0, mobjTable.RealSource, GetViewName(mlngViewID))
  
End Function

Private Function GetViewName(plngViewID As Long) As String
  ' Return the name of the given view.
  Dim sName As String
  Dim objTableView As CTablePrivilege

  Set objTableView = gcoTablePrivileges.FindViewID(plngViewID)
  If Not objTableView Is Nothing Then
    sName = objTableView.ViewName
  Else
    sName = ""
  End If
  
  Set objTableView = Nothing
  
  GetViewName = sName
  
End Function


Private Sub ConfigureOrdersCombo()
  ' Initialise the form to be called from a primary screen.
  Dim fOrderFound As Boolean
  Dim iIndex As Integer
  Dim rsOrder As Recordset

  If mlngViewID > 0 Then
    Set rsOrder = datGeneral.GetViewOrders(mlngViewID, mobjTable.TableID)
  Else
    Set rsOrder = datGeneral.GetOrders(mobjTable.TableID)
  End If
  
  ' Populate the Orders combo.
  With cmbOrders
    .Clear
  
    ' Add the orders to the combo.
    Do While Not rsOrder.EOF
      .AddItem RemoveUnderScores(Trim(rsOrder!Name))
      .ItemData(.NewIndex) = rsOrder!OrderID
      rsOrder.MoveNext
    Loop
    rsOrder.Close
    Set rsOrder = Nothing
        
    If .ListCount > 0 Then
      ' Select the last used order if possible.
      fOrderFound = False
      For iIndex = 0 To (.ListCount - 1)
        If (.ItemData(iIndex) = mlngOrderID) Then
          fOrderFound = True
          .ListIndex = iIndex
          Exit For
        End If
      Next iIndex
      
      If Not fOrderFound Then
        .ListIndex = 0
      End If
      
      .Enabled = True
    Else
      ' No orders.
      .AddItem "<No Order>"
      .ListIndex = 0
      MsgBox "No orders defined for this table.", vbInformation, Me.Caption
      .Enabled = False
    End If
  End With
  
End Sub










Private Sub Form_Unload(Cancel As Integer)
  'Tidy things up before unloading
  If mrsFindRecords.State <> adStateClosed Then mrsFindRecords.Close
  Set mrsFindRecords = Nothing

  Unhook Me.hWnd
End Sub


Private Sub ssOleDBGridFindColumns_DblClick()
  If ssOleDBGridFindColumns.SelBookmarks.Count > 0 Then
    mfCancelled = False
    Me.Hide
  End If

End Sub


Public Function SelectedRecordIDs() As Variant
  ' Return the ID of the selected reocrd in the grid.
  Dim alngSelectedIDs() As Long
  Dim lngIndex As Long
  
  ReDim alngSelectedIDs(0)
  
  For lngIndex = 0 To (ssOleDBGridFindColumns.SelBookmarks.Count - 1)
    ReDim Preserve alngSelectedIDs(lngIndex + 1)
    
    'JDM - 31/10/01 - Fault 3055 - Didn't select bottom record - Why? The control is pants...
    If ssOleDBGridFindColumns.SelBookmarks(lngIndex) >= ssOleDBGridFindColumns.Rows Then
      ssOleDBGridFindColumns.MoveFirst
      ssOleDBGridFindColumns.MoveRecords ssOleDBGridFindColumns.Rows
    Else
    
      'JDM - 30/10/01 - Fault 3054 - Problem with setting the bookmark - It's occassionally a few records out.
      '                              Setting the firstrow *seems* to fix it.
      ssOleDBGridFindColumns.FirstRow = ssOleDBGridFindColumns.SelBookmarks(lngIndex)
      ssOleDBGridFindColumns.Bookmark = ssOleDBGridFindColumns.SelBookmarks(lngIndex)
    
    End If
    
    alngSelectedIDs(lngIndex + 1) = ssOleDBGridFindColumns.Columns((ssOleDBGridFindColumns.Cols - 1)).Value
  Next lngIndex
  
  SelectedRecordIDs = alngSelectedIDs
  
End Function


Private Sub ssOleDBGridFindColumns_KeyPress(KeyAscii As Integer)
  Dim lngThistime As Long
  Static sFind As String
  Static lngLastTime As Long
  
  Select Case KeyAscii
    Case vbKeyReturn
      ssOleDBGridFindColumns_DblClick
    
    ' Otherwise find the record
    Case Else
      ' Only search for alphanumeric characters.
      If (KeyAscii >= 32) And (KeyAscii <= 255) Then
        lngThistime = GetTickCount
        If lngLastTime + 1500 < lngThistime Then
          sFind = Chr(KeyAscii)
        Else
          sFind = sFind & Chr(KeyAscii)
        End If
        lngLastTime = lngThistime
        LocateRecord sFind
      End If
  End Select

End Sub



Private Sub LocateRecord(psSearchString As String)
  Dim fFound As Boolean
  Dim fUseBinarySearch As Boolean
  Dim iIndex As Long
  Dim iComparisonResult As Integer
  Dim lngLoop As Long
  Dim lngUpper As Long
  Dim lngLower As Long
  Dim lngJump As Long
  Dim lngFirstFindColumn As Long
  Dim lngFirstOrderColumn As Long
  Dim varFoundBookmark As Variant
  Dim varOriginalBookmark As Variant

  If ssOleDBGridFindColumns.Rows = 0 Then
    Exit Sub
  End If
  
  Screen.MousePointer = vbHourglass
  
  fUseBinarySearch = mfFirstColumnsMatch

  If fUseBinarySearch Then
    If (miFirstColumnDataType <> sqlVarChar) And _
     (miFirstColumnDataType <> sqlVarBinary) And _
     (miFirstColumnDataType <> sqlNumeric) And _
     (miFirstColumnDataType <> sqlInteger) Then

      fUseBinarySearch = False
    End If
  End If

  ' Search the grid for the required string.
  fFound = False

  lngLower = 1
  lngUpper = mlngRecordCount

  With ssOleDBGridFindColumns
    .Redraw = False

    varOriginalBookmark = .Bookmark

    If fUseBinarySearch Then
      ' Binary search the grid for the required string.
      Do
        Select Case miFirstColumnDataType
          Case sqlVarChar, sqlVarBinary
            ' JPD String comparison changed from using VB's strComp function to
            ' using our own DictionaryCompareStrings function. VB's strComp
            ' function does not use the same order as that used when SQL orders
            ' by a character column. The DictionaryCompareStrings does.
            'iComparisonResult = StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare)
            iComparisonResult = datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString))
            
          Case sqlNumeric, sqlInteger
            If Val(ssOleDBGridFindColumns.Columns(0).Text) = Val(psSearchString) Then
              iComparisonResult = 0
            ElseIf Val(ssOleDBGridFindColumns.Columns(0).Text) < Val(psSearchString) Then
              iComparisonResult = -1
            Else
              iComparisonResult = 1
            End If
        End Select

        If Not mfFirstColumnAscending Then
          iComparisonResult = iComparisonResult * -1
        End If

        Select Case iComparisonResult
          Case 0    ' String found.
            fFound = True
            varFoundBookmark = .Bookmark
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do

          Case -1   ' Current record is before the required record.
            lngLower = .Bookmark + 1
            lngJump = ((lngUpper - .Bookmark) \ 2)
            If lngLower > lngUpper Then Exit Do

          Case 1    ' Current record is after the required record.
            lngUpper = .Bookmark - 1
            lngJump = -((.Bookmark - lngLower) \ 2) - 1
            If lngLower > lngUpper Then Exit Do
        End Select

        If lngLower = lngUpper Then
          lngJump = lngUpper - .Bookmark
        End If

        ' Move to the middle record of the recmaining records to search.
        .MoveRecords lngJump
      Loop

      If fFound Then
        .Bookmark = varFoundBookmark
      Else
        .MoveRecords varOriginalBookmark - .Bookmark
      End If
    Else
      ' Sequential search the grid for the required string.
      .MoveFirst
      For lngLoop = lngLower To lngUpper
        ' JPD String comparison changed from using VB's strComp function to
        ' using our own DictionaryCompareStrings function. VB's strComp
        ' function does not use the same order as that used when SQL orders
        ' by a character column. The DictionaryCompareStrings does.
        'If StrComp(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare) = 0 Then
        If datGeneral.DictionaryCompareStrings(UCase(Left(ssOleDBGridFindColumns.Columns(0).Text, Len(psSearchString))), UCase(psSearchString)) = 0 Then
          Exit For
        End If

        If lngLoop < lngUpper Then
          .MoveNext
        Else
          .Bookmark = varOriginalBookmark
        End If
      Next lngLoop
    End If

    .SelBookmarks.RemoveAll
    .SelBookmarks.Add .Bookmark

    .Redraw = True
  End With
  
  Screen.MousePointer = vbDefault

End Sub





Private Sub ssOleDBGridFindColumns_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsFindRecords.MoveLast
    Else
      mrsFindRecords.MoveFirst
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
  End If

  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrsFindRecords.Move NumberOfRowsToMove
  NewLocation = mrsFindRecords.Bookmark

End Sub


Private Sub ssOleDBGridFindColumns_UnboundReadData(ByVal RowBuf As SSDataWidgets_B_OLEDB.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim strFormat As String
  
  iRowsRead = 0
  
  ' Do nothing if we a re just formatting the grid,
  ' ot if there a re no records to display.
  If (mfFormattingGrid) Or (mlngRecordCount = 0) Then Exit Sub
  
  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsFindRecords.EOF Then
        mrsFindRecords.MoveLast
      End If
    Else
      If Not mrsFindRecords.BOF Then
        mrsFindRecords.MoveFirst
      End If
    End If
  Else
    mrsFindRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsFindRecords.BOF Or mrsFindRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsFindRecords.Fields.Count - 1)
          Select Case mrsFindRecords.Fields(iFieldIndex).Type
            Case adDBTimeStamp
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), DateFormat)
            
            Case adNumeric
              ' Are thousand separators used
              strFormat = "0"
              If mavFindColumns(3, iFieldIndex) Then
                strFormat = "#,0"
              End If
              If mavFindColumns(2, iFieldIndex) > 0 Then
                strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              End If
              
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsFindRecords(iFieldIndex), strFormat)
            
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsFindRecords(iFieldIndex)
          
          End Select
          
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
  
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsFindRecords.Bookmark
  
    End Select
    
    If ReadPriorRows Then
      mrsFindRecords.MovePrevious
    Else
      mrsFindRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub



