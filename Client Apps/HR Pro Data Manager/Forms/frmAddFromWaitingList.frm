VERSION 5.00
Object = "{0F987290-56EE-11D0-9C43-00A0C90F29FC}#1.0#0"; "ActBar.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmAddFromWaitingList 
   Caption         =   "Add Delegate From Waiting List"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1006
   Icon            =   "frmAddFromWaitingList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraOrder 
      Caption         =   "Options :"
      Height          =   1200
      Left            =   100
      TabIndex        =   4
      Top             =   100
      Width           =   6000
      Begin VB.ComboBox cmbOrders 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   700
         Width           =   4900
      End
      Begin VB.ComboBox cmbView 
         Height          =   315
         Left            =   900
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   4900
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "View :"
         Height          =   270
         Left            =   200
         TabIndex        =   8
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblOrder 
         BackStyle       =   0  'Transparent
         Caption         =   "Order :"
         Height          =   255
         Left            =   200
         TabIndex        =   7
         Top             =   760
         Width           =   600
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   3495
      TabIndex        =   3
      Top             =   3850
      Width           =   2610
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   400
         Left            =   60
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1410
         TabIndex        =   2
         Top             =   0
         Width           =   1200
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdDelegates 
      Height          =   2010
      Left            =   100
      TabIndex        =   0
      Top             =   1500
      Width           =   6000
      _Version        =   196617
      DataMode        =   1
      RecordSelectors =   0   'False
      GroupHeaders    =   0   'False
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
      SelectTypeCol   =   0
      SelectTypeRow   =   3
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   10583
      _ExtentY        =   3545
      _StockProps     =   79
      ForeColor       =   0
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
   Begin ActiveBarLibraryCtl.ActiveBar ActiveBar1 
      Left            =   90
      Tag             =   "BAND_FIND"
      Top             =   3885
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
      Bands           =   "frmAddFromWaitingList.frx":000C
   End
End
Attribute VB_Name = "frmAddFromWaitingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjDelegateTable As CTablePrivilege

' Course record variables.
Private mlngCourseID As Long
Private msCourseTitle As String

' Delegate recordset variables.
Private mrsDelegateRecords As New ADODB.Recordset
Private mlngRecordCount As Long
Private mlngSelectedRecordID As Long
Private mlngSelectedWLID As Long

' Delegate recordset location variables.
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer

' Form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean

' Print button variable
Private mblnPrintCancelled As Boolean

Private mavFindColumns() As Variant        ' Find columns details

'NPG20080107 Fault 12866
' Link table variables.
Private mlngLookupColumnID As Long
Private msLookupColumnName As String
' Record display option variables.
Private mlngOrderID As Long
Private mlngViewID As Long

'NPG20080317 Fault 12983
Private mintViewSelectedIndex As Integer
Private mintOrderSelectedIndex As Integer

Private Const dblFORM_MINWIDTH = 5000
Private Const dblFORM_MINHEIGHT = 5000

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property


'To deactivate the right-click menu totally (preventing customisation)
'NB : The Allow Customisation = False property seems not to work!
Private Sub ActiveBar1_PreCustomizeMenu(ByVal Cancel As ActiveBarLibraryCtl.ReturnBool)
  Cancel = True
End Sub


Private Function GetDelegateRecords() As Boolean
  ' Construct a recordset of the delegates that have the given course title
  ' on their Waiting List.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNoSelect As Boolean
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  Dim lngFirstFindColumnID As Long
  Dim lngFirstSortColumnID As Long
  Dim sSQL As String
  Dim sRecordCount As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sOrderString As String
  Dim sWhereCode As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim rsTemp As Recordset
  Dim objTableView As CTablePrivilege
  Dim objDelegateTable As CTablePrivilege
  Dim objWaitingListTable As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
  Dim lngColumnID As Long
  'NPG20080109 Fault 12866
  Dim sSource As String
  

  Screen.MousePointer = vbHourglass
  
  fNoSelect = False
  
  sOrderString = ""
  sJoinCode = ""
  sColumnList = ""
  sWhereCode = ""

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

  ' Get the Delegate table object.
  Set objDelegateTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)

  Set objWaitingListTable = gcoTablePrivileges.Item(gsWaitListTableName)

  If Len(gsWaitListOverrideColumnName) > 0 Then
    ' Get the column privileges collection for the given table.
    sRealSource = objWaitingListTable.RealSource
  
    Set objColumnPrivileges = GetColumnPrivileges(gsWaitListTableName)
    fColumnOK = objColumnPrivileges.Item(gsWaitListOverrideColumnName).AllowSelect
    lngColumnID = objColumnPrivileges.Item(gsWaitListOverrideColumnName).ColumnID
    Set objColumnPrivileges = Nothing
  
    If fColumnOK Then
      ' The column CAN be read from the WL table.
      ' Add the column to the column list.
      sColumnList = sColumnList & _
        IIf(Len(sColumnList) > 0, ", ", "") & _
        sRealSource & "." & Trim(gsWaitListOverrideColumnName)
  
      mavFindColumns(0, UBound(mavFindColumns, 2)) = lngColumnID
      mavFindColumns(1, UBound(mavFindColumns, 2)) = datGeneral.GetDataSize(lngColumnID)
      mavFindColumns(2, UBound(mavFindColumns, 2)) = datGeneral.GetDecimalsSize(lngColumnID)
      mavFindColumns(3, UBound(mavFindColumns, 2)) = datGeneral.DoesColumnUseSeparators(lngColumnID)
      ReDim Preserve mavFindColumns(3, UBound(mavFindColumns, 2) + 1)
  
      ' Remember the first Find column.
      If lngFirstFindColumnID = 0 Then
        lngFirstFindColumnID = lngColumnID
      End If
  
      ' Add the column to the order string.
      sOrderString = sOrderString & _
        IIf(Len(sOrderString) > 0, ", ", "") & _
        sRealSource & "." & Trim(gsWaitListOverrideColumnName)
  
      ' Remember the first Order column.
      If lngFirstSortColumnID = 0 Then
        lngFirstSortColumnID = lngColumnID
        mfFirstColumnAscending = True
        miFirstColumnDataType = sqlDate
      End If
    End If
  End If


  ' Get the default order items from the database.
  
  'NPG20080109 Fault 12866
  ' This needs to exclude any columns that are in the order, but not in the view...
  Set rsInfo = datGeneral.GetOrderDefinition(mlngOrderID)

  fOK = Not (rsInfo.EOF And rsInfo.BOF)
  If Not fOK Then
    COAMsgBox "No default order defined for the delegate table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      ' Get the column privileges collection for the given table.
      '      sRealSource = gcoTablePrivileges.Item(rsInfo!TableName).RealSource
      '      Set objColumnPrivileges = GetColumnPrivileges(rsInfo!TableName)

      
      'NPG20080220 Fault 12911
      If rsInfo!TableID = mobjDelegateTable.TableID Then
        If mlngViewID = 0 Then
          sSource = mobjDelegateTable.TableName
        Else
          sSource = GetViewName(mlngViewID)
        End If
      Else
        sSource = rsInfo!TableName
      End If
        
      sRealSource = gcoTablePrivileges.Item(sSource).RealSource
      Set objColumnPrivileges = GetColumnPrivileges(sSource)
      
      
      'NPG20080109 Fault 12866
      ' is this column in the view? If not, skip it...
      
      fColumnOK = objColumnPrivileges.IsValid(rsInfo!ColumnName)
      
      If fColumnOK Then
        fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect
      End If
      Set objColumnPrivileges = Nothing

      If fColumnOK Then
        ' The column CAN be read from the Delegate table, or directly from a parent table.
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
        If rsInfo!TableID <> glngEmployeeTableID Then
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
          End If
        End If
      ElseIf rsInfo!TableID <> mobjDelegateTable.TableID Then
        ' The column CANNOT be read from the Delegate table, or directly from a parent table.
        ' Try to read it from the views on the table.
        
        ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
        ReDim asViews(0)
        For Each objTableView In gcoTablePrivileges.Collection
          If (Not objTableView.IsTable) And _
            (objTableView.TableID = rsInfo!TableID) And _
            (objTableView.AllowSelect) Then
              
            sSource = objTableView.ViewName
            sRealSource = gcoTablePrivileges.Item(objTableView.ViewName).RealSource

            ' Get the column permission for the view.
            Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)

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

'NPG20080109 Fault 12866
ContinueLoop:
      rsInfo.MoveNext
    Loop

    ' Inform the user if they do not have permission to see the data.
    If fNoSelect Then
      COAMsgBox "You do not have 'read' permission on all of the columns in the selected order." & _
        vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
    End If
    
    mfFirstColumnsMatch = (lngFirstFindColumnID = lngFirstSortColumnID)
  
'NPG20080317 Fault 12893
  fOK = False
  If Len(gsWaitListOverrideColumnName) > 0 Then
      fOK = IIf(InStr(1, sColumnList, ",", vbTextCompare) > 0, True, False)
  Else
      fOK = IIf(Len(sColumnList) > 0, True, False)
  End If

  If fOK Then
    ' Use the Delegate table as the base if it can be read.
    If (objDelegateTable.AllowSelect) Or _
      (objDelegateTable.TableType = tabTopLevel) Then
              
'NPG20080109 Fault 12866
'        sSQL = "SELECT " & sColumnList & ", " & _
'          objDelegateTable.RealSource & ".id," & _
'          objWaitingListTable.RealSource & ".id AS '??WaitingListRecordID'" & _
'          " FROM " & objDelegateTable.RealSource
      
'        sSQL = "SELECT " & sColumnList & ", " & _
'          sRealSource & ".id," & _
'          objWaitingListTable.RealSource & ".id AS '??WaitingListRecordID'" & _
'          " FROM " & sRealSource
        sSQL = "SELECT " & sColumnList & ", " & _
          CurrentTableViewName & ".id, " & _
          objWaitingListTable.RealSource & ".id AS '??WaitingListRecordID'" & _
          " FROM " & CurrentTableViewName
          
'        sRecordCount = "SELECT COUNT(" & sRealSource & ".ID)" & _
'          " FROM " & sRealSource
        sRecordCount = "SELECT COUNT(" & CurrentTableViewName & ".ID)" & _
          " FROM " & CurrentTableViewName
        
        ' Join any other tables and views that are used.
        For iNextIndex = 1 To UBound(alngTableViews, 2)
          If alngTableViews(1, iNextIndex) = 0 Then
            Set objTableView = gcoTablePrivileges.FindTableID(alngTableViews(2, iNextIndex))
          Else
            Set objTableView = gcoTablePrivileges.FindViewID(alngTableViews(2, iNextIndex))
          End If
          
          If objTableView.TableID = glngEmployeeTableID Then
            ' Join a view of the Delegate table.
'            sSQL = sSQL & _
'              " LEFT OUTER JOIN " & objTableView.RealSource & _
'              " ON " & sRealSource & ".ID = " & objTableView.RealSource & ".ID"
'            sRecordCount = sRecordCount & _
'              " LEFT OUTER JOIN " & objTableView.RealSource & _
'              " ON " & sRealSource & ".ID = " & objTableView.RealSource & ".ID"
'            If Not objDelegateTable.AllowSelect Then
'              sWhereCode = sWhereCode & _
'                IIf(Len(sWhereCode) > 0, " OR (", "(") & _
'                sRealSource & ".ID IN (SELECT ID FROM " & objTableView.RealSource & "))"
'            End If
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & CurrentTableViewName & ".ID = " & objTableView.RealSource & ".ID"
            sRecordCount = sRecordCount & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & CurrentTableViewName & ".ID = " & objTableView.RealSource & ".ID"
            If Not mobjDelegateTable.AllowSelect Then
              sWhereCode = sWhereCode & _
                IIf(Len(sWhereCode) > 0, " OR (", "(") & _
                CurrentTableViewName & ".ID IN (SELECT ID FROM " & objTableView.RealSource & "))"
            End If
          Else
            ' Join a parent table/view.
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & sRealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
              " = " & objTableView.RealSource & ".ID"
          End If
          Set objTableView = Nothing
        Next iNextIndex

        sSQL = sSQL & _
          " INNER JOIN " & objWaitingListTable.RealSource & _
          " ON (" & sRealSource & ".id = " & objWaitingListTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & _
          " AND " & objWaitingListTable.RealSource & "." & gsWaitListCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "')"
      
        sRecordCount = sRecordCount & _
          " INNER JOIN " & objWaitingListTable.RealSource & _
          " ON (" & sRealSource & ".id = " & objWaitingListTable.RealSource & ".id_" & Trim(Str(glngEmployeeTableID)) & _
          " AND " & objWaitingListTable.RealSource & "." & gsWaitListCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "')"
      
        sSQL = sSQL & _
          IIf(Len(sWhereCode) > 0, " WHERE " & sWhereCode, "")
        sRecordCount = sRecordCount & _
          IIf(Len(sWhereCode) > 0, " WHERE " & sWhereCode, "")
          
        ' Tag on the 'order by' code.
        sSQL = sSQL & _
          IIf(Len(sOrderString) > 0, " ORDER BY " & sOrderString, "")
      
        ' Get the required recordset.
        Set mrsDelegateRecords = datGeneral.GetPersistentRecords(sSQL, adOpenStatic, adLockReadOnly)
          
        ' Get the recordset's record count.
        Set rsTemp = datGeneral.GetRecords(sRecordCount)
        If (rsTemp.EOF And rsTemp.BOF) Then
          mlngRecordCount = 0
        Else
          mlngRecordCount = rsTemp(0)
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    
        ' Check we have delegate records.
        fOK = (mlngRecordCount > 0)
        If Not fOK Then
          COAMsgBox "No delegate records found.", vbExclamation, Me.Caption
          ConfigureGrid
          ' NPG20090120 Fault 13343
          ' Allow the method to continue so user can select a different view
          fOK = True
        End If
        
        If fOK Then
          ' Configure the grid.
          ConfigureGrid
        End If
      Else
        ' Unable to read from the delegate table.
        COAMsgBox "You do not have permission to read the Delegate table." & _
          vbCrLf & "Unable to display records.", vbExclamation, "Security"
        fOK = False
      End If
    Else
      COAMsgBox "You do not have permission to read any of the columns in the Delegate table's default order." & _
        vbCrLf & "Unable to display records.", vbExclamation, "Security"
      fOK = False
    End If
  End If

  'NPG20080317 Fault 12984
  If fOK = False Then
    cmbView.ListIndex = mintViewSelectedIndex
    cmbView.ListIndex = mintOrderSelectedIndex
  Else
    mintViewSelectedIndex = cmbView.ListIndex
    mintOrderSelectedIndex = cmbView.ListIndex
  End If


  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  GetDelegateRecords = fOK
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error reading Delegate records.", vbExclamation, Me.Caption
  fOK = False
  Resume TidyUpAndExit
  
End Function


Public Sub ResizeFindColumns()

  Dim dblCurrentSize As Double
  Dim dblNewSize As Double
  Dim iCount As Integer
  Dim dblResizeFactor As Double
  Dim bNeedScrollBars As Boolean

  With grdDelegates

    .Redraw = False
    bNeedScrollBars = IIf(.Rows > .VisibleRows, True, False)

    ' Calculate the existing size of the find grid
    dblCurrentSize = 0
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then
        dblCurrentSize = dblCurrentSize + .Columns(iCount).Width
      End If
    Next iCount

    ' Calculate size of resized grid
    dblNewSize = .Width
    If bNeedScrollBars Then
      dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX)
    End If
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXFRAME) * 2)
    dblNewSize = dblNewSize - (UI.GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX)

    ' Silly bug with resizing
    If dblCurrentSize < 1 Or .Cols < 1 Then Exit Sub

    ' Calculate the ratio that the grid needs to be resized to
    dblResizeFactor = Round(dblNewSize / dblCurrentSize, 2)

    ' Scroll through adjusting each column according to the resize factor
    For iCount = 0 To (.Cols - 1)
      If .Columns(iCount).Visible Then

        'Make the last column nice & snug
        If iCount = (.Cols - 2) Then
          .Columns(iCount).Width = dblNewSize
        Else
          .Columns(iCount).Width = (.Columns(iCount).Width * dblResizeFactor)
          dblNewSize = dblNewSize - .Columns(iCount).Width - 5
        End If

        SaveUserSetting "FindOrder" + LTrim(Str(mlngOrderID)), .Columns(iCount).Name, .Columns(iCount).Width

      End If
    Next iCount

    .Redraw = True

  End With

End Sub

Private Function GetViewName(plngViewID As Long) As String
  'NPG20080109 Fault 12866
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

Private Function CurrentTableViewName() As String
  ' Return the name of the table, or view.
  CurrentTableViewName = IIf(mlngViewID = 0, mobjDelegateTable.RealSource, GetViewName(mlngViewID))
  
End Function



Public Function Initialise(plngCourseID As Long, _
  pobjCourseTableView As CTablePrivilege) As Boolean
  ' Initialise the form.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsCourses As ADODB.Recordset
  Dim objColumns As CColumnPrivileges
  
  Set mobjDelegateTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)
  
  mlngCourseID = plngCourseID
  
  fOK = ValidateParameters
  
  If fOK Then
    ' Check that the course title can be read from the selected course record.
    If pobjCourseTableView.IsTable Then
      Set objColumns = GetColumnPrivileges(pobjCourseTableView.TableName)
    Else
      Set objColumns = GetColumnPrivileges(pobjCourseTableView.ViewName)
    End If
    
    fOK = objColumns.IsValid(gsCourseTitleColumnName)
    If Not fOK Then
      COAMsgBox "The '" & gsCourseTitleColumnName & "' column is not in your current view.", vbOKOnly + vbInformation, App.ProductName
    End If
  End If
  
  If fOK Then
    fOK = objColumns.Item(gsCourseTitleColumnName).AllowSelect
    If Not fOK Then
      COAMsgBox "You do not have 'read' permission on the '" & gsCourseTitleColumnName & "' column.", vbOKOnly + vbInformation, App.ProductName
    End If
  End If
  
  Set objColumns = Nothing
    
  If fOK Then
    ' Get the given course title.
    sSQL = "SELECT " & gsCourseTitleColumnName & _
      " FROM " & pobjCourseTableView.RealSource & _
      " WHERE id = " & Trim(Str(plngCourseID))
      
    Set rsCourses = datGeneral.GetRecords(sSQL)
    With rsCourses
      fOK = Not (.EOF And .BOF)
        
      If fOK Then
        ' Read the course details into member variables.
        msCourseTitle = IIf(IsNull(.Fields(gsCourseTitleColumnName)), "", .Fields(gsCourseTitleColumnName))
      End If
        
      .Close
    End With
    Set rsCourses = Nothing
   
   
  mlngOrderID = glngEmployeeOrderID

   
'NPG20080107 Fault 12866
  ' Populate the View combo.
  ConfigureViewCombo '1 'plngViewID
  If cmbView.ListCount = 0 Then
    COAMsgBox "You do not have 'read' permission on this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
    Initialise = False
    Exit Function
  End If
   
   
  ' Populate the Orders combo.
  ConfigureOrdersCombo
  Screen.MousePointer = vbHourglass
  If cmbOrders.ListCount = 0 Then
    COAMsgBox "There are no orders defined for this table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
    mfCancelled = True
    Me.Hide
    Initialise = False
    Exit Function
  End If
   
   
    ' Get the required course records.
    fOK = GetDelegateRecords
  End If
  
  Initialise = fOK
  
End Function

Private Sub cmbOrders_Click()
  'NPG20080109 Fault 12866
  
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
    
    GetDelegateRecords
  
    With grdDelegates
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

'NPG20080107 Fault 12866
Private Sub cmbView_Click()
  Dim fOK As Boolean

  ' Do nothing if the form is not visible.
  fOK = Me.Visible

  If fOK Then
    Screen.MousePointer = vbHourglass
  End If


  ' NPG20080411 Fault 12913
  ' Do nothing if there are no records.
  ' NPG20081201 Fault 13342 & 13343 - remove this block of code.
  '  If fOK Then
  '    fOK = (mlngRecordCount > 0)
  '  End If

  If fOK Then
    mlngViewID = cmbView.ItemData(cmbView.ListIndex)
    
    GetDelegateRecords

    With grdDelegates
      cmdSelect.Enabled = (.Rows > 0)

      If .Rows > 0 Then
        .MoveFirst
        .SelBookmarks.Add (.Bookmark)
'        .SetFocus
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
  grdDelegates_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
  Case vbKeyF1
    If ShowAirHelp(Me.HelpContextID) Then
      KeyCode = 0
    End If
End Select

End Sub

Private Sub Form_Load()

  ' Get rid of the icon off the form
  RemoveIcon Me
  
  Hook Me.hWnd, dblFORM_MINWIDTH, dblFORM_MINHEIGHT
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    mfCancelled = True
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

  If Me.WindowState = vbNormal Then

    'NPG20080220 Fault 12910
    fraOrder.Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
    cmbOrders.Width = fraOrder.Width - (dblCOORD_XGAP * 6)
    cmbView.Width = fraOrder.Width - (dblCOORD_XGAP * 6)
  
    ' Size the grid.
    With grdDelegates
      'NPG20080220 Fault 12910
      '.Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
      .Width = fraOrder.Width
      .Height = Me.ScaleHeight - .Top - fraButtons.Height - (2 * dblCOORD_YGAP)
    End With
  
    fraButtons.Top = grdDelegates.Top + grdDelegates.Height + dblCOORD_YGAP
    fraButtons.Left = Me.ScaleWidth - fraButtons.Width - dblCOORD_XGAP
  
    DoEvents
  
    If ((grdDelegates.Rows - grdDelegates.FirstRow + 1) < (grdDelegates.VisibleRows)) And _
      (grdDelegates.FirstRow > 1) Then
  
      grdDelegates.FirstRow = IIf(grdDelegates.Rows - grdDelegates.VisibleRows + 1 < 1, _
        1, grdDelegates.Rows - grdDelegates.VisibleRows + 1)
    End If
  
    ResizeFindColumns

  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Tidy things up before unloading
  If Not mrsDelegateRecords Is Nothing Then
    If mrsDelegateRecords.State = adStateOpen Then
      mrsDelegateRecords.Close
    End If
    Set mrsDelegateRecords = Nothing
  End If
  
  Unhook Me.hWnd

End Sub



Private Sub ConfigureGrid()
  ' Populate the grid.
  Dim iLoop As Integer
  Dim lngWidth As Long
  
  UI.LockWindow Me.hWnd
  
  lngWidth = 0
  mfFormattingGrid = True
 
  'Check if an override has been set up
  
  With grdDelegates
    .Redraw = False
    .Columns.RemoveAll
    
    For iLoop = 0 To (mrsDelegateRecords.Fields.Count - 1)
      .Columns.Add iLoop
      .Columns(iLoop).Name = mrsDelegateRecords.Fields(iLoop).Name
      .Columns(iLoop).Visible = (UCase(mrsDelegateRecords.Fields(iLoop).Name) <> "ID") And _
        (Left(mrsDelegateRecords.Fields(iLoop).Name, 1) <> "?")
      .Columns(iLoop).Caption = RemoveUnderScores(mrsDelegateRecords.Fields(iLoop).Name)
      .Columns(iLoop).Alignment = ssCaptionAlignmentLeft
      .Columns(iLoop).CaptionAlignment = ssColCapAlignUseColumnAlignment
    
      ' If the find column is a logic column then set the grid column style to be 'checkbox'.
      If mrsDelegateRecords.Fields.Item(iLoop).Type = adBoolean Then
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
    
    ' Select the top row.
    If mlngRecordCount > 0 Then
      .MoveFirst
      .SelBookmarks.Add .Bookmark
    End If
  End With
  
  
  
  ' Adjust the size of the window to fit the grid.
  lngWidth = lngWidth + _
    (((UI.GetSystemMetrics(SM_CXFRAME) * 2) + _
    UI.GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX)

  If grdDelegates.Rows > grdDelegates.VisibleRows Then
    lngWidth = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20
  End If

  Me.Width = lngWidth + 120

  UI.UnlockWindow
  
End Sub

Private Sub grdDelegates_DblClick()
  Dim fOK As Boolean
  Dim nTotalSelRows As Variant
  Dim intCount As Integer
  Dim arrayBookmarks() As Variant
  
'  If grdDelegates.SelBookmarks.Count > 0 Then
'    ' Get the ID of the selected record.
'    mlngSelectedRecordID = grdDelegates.Columns((grdDelegates.Cols - 2)).Value
'    mlngSelectedWLID = grdDelegates.Columns((grdDelegates.Cols - 1)).Value
'
'    ' Create booking record.
'    fOK = CreateBooking
'
'    If fOK Then
'      mfCancelled = False
'      Me.Hide
'    End If
'  End If

  'NHRD15012007 Suggestion from fault log 4867
  'Build a string of all the currently selected event entries
  If grdDelegates.SelBookmarks.Count > 0 Then
    'Workout how many records have been selected
    nTotalSelRows = grdDelegates.SelBookmarks.Count
    'Redimension the arrays to the count of the bookmarks
    ReDim arrayBookmarks(nTotalSelRows)
  
    For intCount = 1 To nTotalSelRows
      arrayBookmarks(intCount) = grdDelegates.SelBookmarks.Item(intCount - 1)
    Next intCount
  
    For intCount = 1 To nTotalSelRows
      grdDelegates.Bookmark = arrayBookmarks(intCount)
      ' Get the ID of the selected record.
      mlngSelectedRecordID = grdDelegates.Columns((grdDelegates.Cols - 2)).Value
      mlngSelectedWLID = grdDelegates.Columns((grdDelegates.Cols - 1)).Value
      
      ' Create booking record.
      fOK = CreateBooking
      
    Next intCount
    
    If fOK Then
      mfCancelled = False
      Me.Hide
    End If
  End If
  
End Sub

Private Sub grdDelegates_KeyPress(KeyAscii As Integer)
  Dim lngThistime As Long
  Static sFind As String
  Static lngLastTime As Long
  
  Select Case KeyAscii
    Case vbKeyReturn
      grdDelegates_DblClick
    
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
  
  If grdDelegates.Rows = 0 Then
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
  
  With grdDelegates
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
            'iComparisonResult = StrComp(UCase(Left(.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare)
            iComparisonResult = datGeneral.DictionaryCompareStrings(UCase(Left(.Columns(0).Text, Len(psSearchString))), UCase(psSearchString))
            
          Case sqlNumeric, sqlInteger
            If Val(.Columns(0).Text) = Val(psSearchString) Then
              iComparisonResult = 0
            ElseIf Val(.Columns(0).Text) < Val(psSearchString) Then
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
        'If StrComp(UCase(Left(.Columns(0).Text, Len(psSearchString))), UCase(psSearchString), vbTextCompare) = 0 Then
        If datGeneral.DictionaryCompareStrings(UCase(Left(.Columns(0).Text, Len(psSearchString))), UCase(psSearchString)) = 0 Then
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



Private Sub ConfigureOrdersCombo()
  'NPG20080109 Fault 12866

  ' Initialise the form to be called from a primary screen.
  Dim fOrderFound As Boolean
  Dim iIndex As Integer
  Dim rsOrder As Recordset
  
  Set rsOrder = datGeneral.GetOrders(mobjDelegateTable.TableID)
  
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
      .AddItem "<No order>"
      .ItemData(.NewIndex) = 0
      .ListIndex = 0
      COAMsgBox "No orders defined for this " & IIf(mlngViewID > 0, "table.", "view."), vbInformation, Me.Caption
      .Enabled = False
    End If
  End With
  
End Sub


Private Sub ConfigureViewCombo()
  ' Populate the 'views' combo with the views available for the link table.
  Dim objTableView As CTablePrivilege
  Dim iListIndex As Integer
  Dim iLoop As Integer
  
  ' Add the table's views to the combo if the user has permission to read them.
  For Each objTableView In gcoTablePrivileges.Collection
    If (Not objTableView.IsTable) And _
      (objTableView.TableID = mobjDelegateTable.TableID) Then
      
      If objTableView.AllowSelect Then
        With cmbView
            .AddItem "'" & RemoveUnderScores(Trim(objTableView.ViewName)) & "' view"
            .ItemData(cmbView.NewIndex) = objTableView.ViewID
            
'JPD 20051213 Combo is now sorted so we need to determine the default view's index 'after' the combo has been fully populated.
        End With
      End If
    End If
  Next objTableView
  Set objTableView = Nothing
  
  'NHRD20060213 Fault 10516 Moved code from above and forced this to be the first
  ' item in the dropdown with the Optional index parameter in Additem
  ' Add the table to the combo if the user has permission to read it.
  If mobjDelegateTable.AllowSelect Then
    cmbView.AddItem RemoveUnderScores(mobjDelegateTable.TableName), 0
    cmbView.ItemData(cmbView.NewIndex) = 0
  End If
  
  With cmbView
    ' Select the first view.
    If .ListCount > 0 Then
      For iLoop = 0 To cmbView.ListCount - 1
        ' Establish the listindex for the option we want selected by default.
        If .ItemData(iLoop) = glngDefaultBulkBookingViewID Then
          iListIndex = iLoop
          Exit For
        End If
      Next iLoop
      
      .ListIndex = iListIndex
      .Enabled = True
    Else
      .Enabled = False
    End If
  End With

  If cmbView.ListCount > 0 Then
    mlngViewID = cmbView.ItemData(cmbView.ListIndex)
  End If

End Sub

'NPG20080220 Fault 12911
'Private Sub ConfigureViewCombo(lngViewID As Long)
'  ' Populate the 'views' combo with the views available for the link table.
'  Dim fOK As Boolean
'  Dim objTableView As CTablePrivilege
'  Dim objLookupColumns As CColumnPrivileges
'  Dim iViewStart As Integer
'  Dim iLoop As Integer
'  Dim fAdded As Boolean
'  Dim sCaption As String
'
'
'  ' Add the table's views to the combo if the user has permission to read them.
'  For Each objTableView In gcoTablePrivileges.Collection
'    If (Not objTableView.IsTable) And _
'            (objTableView.TableID = glngPersonnelTableID) Then
'      fOK = False
'      If objTableView.AllowSelect Then
'        If mlngLookupColumnID > 0 Then
'          Set objLookupColumns = GetColumnPrivileges(objTableView.ViewName)
'          If objLookupColumns.IsValid(msLookupColumnName) Then
'            If objLookupColumns(msLookupColumnName).AllowSelect Then
'              fOK = True
'            End If
'          End If
'          Set objLookupColumns = Nothing
'        Else
'          fOK = True
'        End If
'      End If
'
'      If fOK Then
'        'JPD 20040416 Fault 8499
'        fAdded = False
'        sCaption = "'" & RemoveUnderScores(Trim(objTableView.ViewName)) & "' view"
'        For iLoop = iViewStart To (cmbView.ListCount - 1)
'          If UCase(sCaption) < UCase(cmbView.List(iLoop)) Then
'            cmbView.AddItem sCaption, iLoop
'            cmbView.ItemData(cmbView.NewIndex) = objTableView.ViewID
'
'            fAdded = True
'
'            'JPD 20040524 Fault 8685
'            Exit For
'          End If
'        Next iLoop
'
'        If Not fAdded Then
'          cmbView.AddItem sCaption
'          cmbView.ItemData(cmbView.NewIndex) = objTableView.ViewID
'        End If
'      End If
'    End If
'  Next objTableView
'  Set objTableView = Nothing
'
'  With cmbView
'    ' Select the first view.
'    If .ListCount > 0 Then
'      '.ListIndex = 0
'      .Enabled = True
'      SetComboItem cmbView, lngViewID
'      If cmbView.ListIndex = -1 Then
'        cmbView.ListIndex = 0
'      End If
'    Else
'      .Enabled = False
'    End If
'  End With
'
'End Sub



Private Sub grdDelegates_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsDelegateRecords.MoveLast
    Else
      mrsDelegateRecords.MoveFirst
    End If
  Else
    mrsDelegateRecords.Bookmark = StartLocation
  End If
  
  'JPD 20040803 Fault 9013
  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrsDelegateRecords.Move NumberOfRowsToMove
  NewLocation = mrsDelegateRecords.Bookmark

End Sub


Private Sub grdDelegates_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
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
      If Not mrsDelegateRecords.EOF Then
        mrsDelegateRecords.MoveLast
      End If
    Else
      If Not mrsDelegateRecords.BOF Then
        mrsDelegateRecords.MoveFirst
      End If
    End If
  Else
    mrsDelegateRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsDelegateRecords.MovePrevious
    Else
      mrsDelegateRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsDelegateRecords.BOF Or mrsDelegateRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsDelegateRecords.Fields.Count - 1)
          Select Case mrsDelegateRecords.Fields(iFieldIndex).Type
            Case adDBTimeStamp
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsDelegateRecords(iFieldIndex), DateFormat)
            
            Case adNumeric
              ' Are thousand separators used
              strFormat = "0"
              If mavFindColumns(3, iFieldIndex) Then
                strFormat = "#,0"
              End If
              If mavFindColumns(2, iFieldIndex) > 0 Then
                strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              End If
              
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsDelegateRecords(iFieldIndex), strFormat)
            
            Case adBoolean
                  'sVALUES = sVALUES & IIf(rsParent.Fields(iFields).Value, 1, 0) & ","
                  RowBuf.Value(iRowIndex, iFieldIndex) = IIf(mrsDelegateRecords(iFieldIndex).Value, "True", "False")
                  
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsDelegateRecords(iFieldIndex)
          
          End Select
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsDelegateRecords.Bookmark
  
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsDelegateRecords.Bookmark
  
    End Select
    
    If ReadPriorRows Then
      mrsDelegateRecords.MovePrevious
    Else
      mrsDelegateRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub



Private Function ValidateParameters() As Boolean
  ' Validate the Training Booking module parameters
  ' used by the 'AddFromWaitingList' function.
  Dim iLoop As Integer
  Dim fValid As Boolean
  Dim objColumn As CColumnPrivilege
  Dim objColumns As CColumnPrivileges
  Dim objTable As CTablePrivilege
  Dim alngRelatedColumns() As Long
  
  ' Check that the Training Booking module is enabled.
  fValid = gfTrainingBookingEnabled

  ' Check that the user has the required permissions on the Training Bookings table.
  If fValid Then
    Set objTable = gcoTablePrivileges.Item(gsTrainBookTableName)
    fValid = objTable.AllowInsert
    If Not fValid Then
      COAMsgBox "You do not have 'new' permission on the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
    End If
    
    If fValid Then
      Set objColumns = GetColumnPrivileges(gsTrainBookTableName)
      
'''      fValid = objColumns.Item(gsTrainBookCourseTitleName).AllowUpdate
'''      If Not fValid Then
'''        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCourseTitleName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
'''      End If
      
      If fValid Then
        fValid = objColumns.Item(gsTrainBookStatusColumnName).AllowUpdate
        If Not fValid Then
          COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
        End If
      End If
      
      ' Check the user has permission to edit all of the related Training Booking columns.
      If fValid Then
        alngRelatedColumns = RelatedColumns
        
        For iLoop = 1 To UBound(alngRelatedColumns, 2)
          Set objColumn = objColumns.FindColumnID(alngRelatedColumns(1, iLoop))
          fValid = Not objColumn Is Nothing
          
          If Not fValid Then
            COAMsgBox "Unable to find all related columns in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
            Exit For
          Else
            fValid = objColumn.AllowUpdate
            If Not fValid Then
              COAMsgBox "You do not have 'edit' permission on the '" & objColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
              Exit For
            End If
          End If
          Set objColumn = Nothing
        Next iLoop
      End If
    
      Set objColumns = Nothing
    End If
    
    Set objTable = Nothing
  End If

  ' Check that the user has the required permissions on the Waiting List table.
  If fValid Then
    Set objTable = gcoTablePrivileges.Item(gsWaitListTableName)
    fValid = objTable.AllowDelete
    If Not fValid Then
      COAMsgBox "You do not have 'delete' permission on the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
    End If
    
    If fValid Then
      Set objColumns = GetColumnPrivileges(gsWaitListTableName)
      
      fValid = objColumns.Item(gsWaitListCourseTitleColumnName).AllowSelect
      If Not fValid Then
        COAMsgBox "You do not have 'read' permission on the '" & gsWaitListCourseTitleColumnName & "' column in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
      End If
    
      ' Check the user has permission to edit all of the related Training Booking columns.
      If fValid Then
        For iLoop = 1 To UBound(alngRelatedColumns, 2)
          Set objColumn = objColumns.FindColumnID(alngRelatedColumns(2, iLoop))
          fValid = Not objColumn Is Nothing
          
          If Not fValid Then
            COAMsgBox "Unable to find all related columns in the '" & gsWaitListTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
            Exit For
          Else
            fValid = objColumn.AllowSelect
            If Not fValid Then
              COAMsgBox "You do not have 'read' permission on the '" & objColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly + vbInformation, App.ProductName
              Exit For
            End If
          End If
          Set objColumn = Nothing
        Next iLoop
      End If
      
      Set objColumns = Nothing
    End If
    
    Set objTable = Nothing
  End If

  ValidateParameters = fValid
  
End Function

Private Function CreateBooking() As Boolean
  ' Create the booking record.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fBooked As Boolean
  Dim fFound As Boolean
  Dim fInTransaction As Boolean
  Dim iLoop As Integer
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sErrorMsg As String
  Dim sColumnName As String
  Dim sColumnList As String
  Dim sValueList As String
  Dim frmPrompt As frmTrainingBookingPrompt
  Dim objDelegateTable As CTablePrivilege
  Dim objTrainingBookingTable As CTablePrivilege
  Dim objWaitingListTable As CTablePrivilege
  Dim objTBColumn As CColumnPrivilege
  Dim objWLColumn As CColumnPrivilege
  Dim objTBColumns As CColumnPrivileges
  Dim objWLColumns As CColumnPrivileges
  Dim alngRelatedColumns() As Long
  Dim asAddedColumns() As String
  Dim sRecordDescription As String
  ' Get the Delegate table object to determine the order
  Set objDelegateTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)
  
  fOK = True
  fInTransaction = False
  
  ' Prompt the user whether they want to Book the course,
  ' or just Provisionally Book it.
  ' JPD 1/3/01 Only prompt if they have 'provisional' as a valid booking status.
  If gfTrainBookStatus_P Then
    Set frmPrompt = New frmTrainingBookingPrompt
    With frmPrompt
      'NHRD15012007 Suggestion from fault log 4867 - Added caption change to see which booking is being actioned
      'NHRD07022007 10384 - modified the caption to always show surname first name using newly created generic EvaluateRecordDescription
      '   NB - Private EvaluateRecordDescription's were used in frmRecEdit4, frmBulkbooking and frmRecordProfilePreview
      '   I have now put a Public generic version in modTrainingBookingSpecifics
      sRecordDescription = EvaluateRecordDescription(mlngSelectedRecordID, objDelegateTable.RecordDescriptionID)
      .Caption = sRecordDescription
      .Show vbModal
  
      fOK = Not .Cancelled
      fBooked = .Booked
      Unload frmPrompt
    End With
    Set frmPrompt = Nothing
    Set objDelegateTable = Nothing
  Else
    fBooked = True
  End If
  
  If fOK Then
    fOK = (mlngSelectedRecordID > 0)
  End If

  ' Check that we are not over-booking a course.
  If fOK Then
    ' Only check that the selected course is not fully booked if the new booking is inlcuded
    ' in the number booked.
    If gfCourseIncludeProvisionals Or fBooked Then
      fOK = TrainingBooking_CheckOverbooking(mlngCourseID, 0)
    End If
  End If

  ' Check that current employee has (or will have) satisfied the pre-requisite criteria.
  If fOK Then
    fOK = TrainingBooking_CheckPreRequisites(mlngCourseID, mlngSelectedRecordID)
  End If
    
  ' Check that the current employee is not unavailable for the selected course.
  If fOK Then
    fOK = TrainingBooking_CheckAvailability(mlngCourseID, mlngSelectedRecordID)
  End If
     
  ' Check that we are not over-lapping another booking.
  If fOK Then
    fOK = TrainingBooking_CheckOverlappedBooking(mlngCourseID, mlngSelectedRecordID, 0)
  End If

  If fOK Then
    Set objTrainingBookingTable = gcoTablePrivileges.Item(gsTrainBookTableName)
    Set objTBColumns = GetColumnPrivileges(gsTrainBookTableName)
    
    Set objWaitingListTable = gcoTablePrivileges.Item(gsWaitListTableName)
    Set objWLColumns = GetColumnPrivileges(gsWaitListTableName)
    
    ' Create the booking record.
    sColumnList = _
      "id_" & Trim(Str(glngEmployeeTableID)) & ", " & _
      "id_" & Trim(Str(glngCourseTableID)) & ", " & _
      gsTrainBookStatusColumnName
      
    sValueList = _
      Trim(Str(mlngSelectedRecordID)) & ", " & _
      Trim(Str(mlngCourseID)) & ", " & _
      IIf(fBooked, "'B'", "'P'")
      
    ' Initialise the array of columns already added to the 'INSERT' street.
    ReDim asAddedColumns(1)
    asAddedColumns(1) = UCase(Trim(gsTrainBookStatusColumnName))
      
    ' Add the related columns to the 'insert' string.
    alngRelatedColumns = RelatedColumns
  
    For iLoop = 1 To UBound(alngRelatedColumns, 2)
      ' Get the column
      Set objTBColumn = objTBColumns.FindColumnID(alngRelatedColumns(1, iLoop))
      Set objWLColumn = objWLColumns.FindColumnID(alngRelatedColumns(2, iLoop))
      
      ' Check that the Training Booking column has not already been added to the 'insert' string.
      fFound = False
      For iNextIndex = 1 To UBound(asAddedColumns)
        If UCase(Trim(objTBColumn.ColumnName)) = asAddedColumns(iNextIndex) Then
          fFound = True
          Exit For
        End If
      Next iNextIndex
    
      If Not fFound Then
        ' The current TB column is not in the 'insert' string so add it now,
        ' and add it to the array of added columns.
        sColumnList = sColumnList & _
          ", " & objTBColumn.ColumnName
      
        iNextIndex = UBound(asAddedColumns) + 1
        ReDim Preserve asAddedColumns(iNextIndex)
        asAddedColumns(iNextIndex) = UCase(Trim(objTBColumn.ColumnName))
        
        sValueList = sValueList & _
          ", " & objWLColumn.ColumnName
      End If
      
      Set objTBColumn = Nothing
      Set objWLColumn = Nothing
    Next iLoop
  End If
  
  If fOK Then
    gADOCon.BeginTrans
    fInTransaction = True
    
    sSQL = "INSERT INTO " & objTrainingBookingTable.RealSource & _
      " (" & sColumnList & ")" & _
      " SELECT " & sValueList & _
      " FROM " & objWaitingListTable.RealSource & _
      " WHERE id = " & Trim(Str(mlngSelectedWLID))
    
    Screen.MousePointer = vbHourglass
    
    sErrorMsg = ""
    fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)

    If Not fOK Then
      Screen.MousePointer = vbDefault
      COAMsgBox "Unable to create booking record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbInformation, App.ProductName
      Screen.MousePointer = vbHourglass
      
      gADOCon.RollbackTrans
      fInTransaction = False
    End If
  End If
  
  If fOK Then
    ' Delete the record in the Waiting List table.
    sSQL = "DELETE FROM " & objWaitingListTable.RealSource & _
      " WHERE id = " & Trim(Str(mlngSelectedWLID))

    sErrorMsg = ""
    fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)

    If Not fOK Then
      Screen.MousePointer = vbDefault
      COAMsgBox "Unable to delete waiting list record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbInformation, App.ProductName
      Screen.MousePointer = vbHourglass
        
      gADOCon.RollbackTrans
      fInTransaction = False
    End If
  End If

TidyUpAndExit:
  If fInTransaction Then
    If fOK Then
      gADOCon.CommitTrans
    Else
      gADOCon.RollbackTrans
    End If
    fInTransaction = False
  End If

  Set objTrainingBookingTable = Nothing
  Set objWaitingListTable = Nothing
  Set objTBColumns = Nothing
  Set objWLColumns = Nothing
  
  Screen.MousePointer = vbDefault
  
  CreateBooking = fOK
  Exit Function

ErrorTrap:
  fOK = False
  COAMsgBox Err.Description, vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit

End Function

Private Sub ActiveBar1_Click(ByVal Tool As ActiveBarLibraryCtl.Tool)
  ' Perform the given toolbar function.
  Select Case Tool.Name
    Case "ID_Print"
     PrintGrid
  End Select
End Sub

Public Sub PrintGrid()
Dim pstrError As String

  On Error GoTo ErrorTrap
  gobjErrorStack.PushStack "clsPrintGrid.PrintGrid()"
  
  'NHRD16072004 Fault 8740
  If Printers.Count = 0 Then
    pstrError = "Unable to print as no printers are installed."
    COAMsgBox pstrError, vbExclamation + vbOKOnly, "HR Pro"
    GoTo TidyUpAndExit
  End If
  
  With grdDelegates
    If .Rows < 1 Then
      COAMsgBox "There is no data in the current view to print", vbInformation + vbOKOnly, App.Title
      Exit Sub
    End If
    .Redraw = False
    Screen.MousePointer = vbHourglass

    .HeadFont.Underline = True

    .PageHeaderFont.Name = "Verdana"
    .PageHeaderFont.Size = 8
    .PageHeaderFont.Bold = True
    .PageHeaderFont.Underline = True

    .PageFooterFont.Name = "Verdana"
    .PageFooterFont.Size = 8
    .PageFooterFont.Bold = False
    .PageFooterFont.Underline = False
    
    ' Force to print all rows
    .PrintData ssPrintAllRows + ssPrintFieldOrder, False, gbPrinterPrompt
    .HeadFont.Underline = False

    .Redraw = True
    Screen.MousePointer = vbDefault

  End With
  
  'TM20011219 Fault 3154 - Show print confirm message if required.
  'NB. Would have used the clsPrintGrid class but the grid on this form is of a
  'different type.
  ' Display a printing complete prompt
  Dim strMBText As String
  Dim msb As frmMessageBox
  Dim iShowMeAgain As Integer
  
  If gbPrinterConfirm And Not (mblnPrintCancelled) Then

'TM20020924 Fault 4356 - Ideally would have used the clsPrintGrid class but this requires a different grid type.
' For the short term have removed the printer device name from the Print Confirm message box.
'    strMBText = "Printing complete." _
'      & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
    strMBText = "Printing complete."
  
    iShowMeAgain = IIf(gbPrinterConfirm, 1, 0)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      'TM20020930 Fault 4462 - the checkbox should not be checked.
      iShowMeAgain = 0
      msb.MessageBox strMBText, vbInformation, App.ProductName, iShowMeAgain, "Don't show me this confirmation again."
      gbPrinterConfirm = IIf(iShowMeAgain = 0, True, False)
      SavePCSetting "Printer", "Confirm", gbPrinterConfirm
      Set msb = Nothing
    End If
  
  End If
  
TidyUpAndExit:
  mblnPrintCancelled = False
  gobjErrorStack.PopStack
  Exit Sub
ErrorTrap:
  gobjErrorStack.HandleError

End Sub

Private Sub grdDelegates_PrintError(ByVal PrintError As Long, Response As Integer)
 'If the user cancelled the print standard dialog box then set print cancelled flag.
  mblnPrintCancelled = False
  If PrintError = 30457 Then 'User cancelled print
    mblnPrintCancelled = True
  End If
  'Set to 0 to prevent a default error message from being displayed
  Response = 0

End Sub
Private Sub grdDelegates_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
    'Define a page header that includes Table and View Name
     ssPrintInfo.PageHeader = "Delegate Waiting List for " + msCourseTitle
    
    'Define a page footer that specifies when the grid was printed.
    'vbTAb will centre the text, two vbTab's will right justify
    'More info in Data Widgets 3.0 Help.
     ssPrintInfo.PageFooter = "Printed on <date> at <Time> by " + gsUserName + "         Page <page number> "

    'Specify that we want each row's height to expand so that all data is displayed,
    'but up to a maximum of 10 lines.
    ssPrintInfo.RowAutoSize = True      'So rows are expanded in height as necessary
    ssPrintInfo.MaxLinesPerRow = 10     'but up to a maximum of 10 lines.
    ssPrintInfo.Portrait = False        'Force Landscape to avoid rows split on two pages
    'Print column and group headers at the top of each page.
    '(Use ssTopOfReport if you want the headers to appear on the first page.)
    ssPrintInfo.PrintHeaders = ssTopOfPage
End Sub


