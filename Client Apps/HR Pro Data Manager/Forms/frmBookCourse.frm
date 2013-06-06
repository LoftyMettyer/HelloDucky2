VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmBookCourse 
   Caption         =   "Book Course"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1010
   Icon            =   "frmBookCourse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   3090
      TabIndex        =   3
      Top             =   3930
      Width           =   2565
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1350
         TabIndex        =   2
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   400
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdCourses 
      Height          =   3600
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5490
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
      SelectTypeRow   =   1
      SelectByCell    =   -1  'True
      BalloonHelp     =   0   'False
      MaxSelectedRows =   1
      ForeColorEven   =   0
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      TabNavigation   =   1
      _ExtentX        =   9684
      _ExtentY        =   6350
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
End
Attribute VB_Name = "frmBookCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Employee record variables.
Private mlngEmployeeID As Long

' Waiting List record variables.
Private mlngWaitListID As Long

' Course recordset variables.
Private mrsCourseRecords As New ADODB.Recordset
Private mlngRecordCount As Long

' Course record variables.
Private msCourseTitle As String
Private mlngSelectedRecordID As Long

' Course recordset location variables.
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer

' Form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean

Private mavFindColumns() As Variant        ' Find columns details

Private Const dblFORM_MINWIDTH = 4000
Private Const dblFORM_MINHEIGHT = 4000

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property

Private Function GetCourseRecords() As Boolean
  ' Construct a recordset of the courses that match the required title, and have a start date
  ' on or after the system date
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNoSelect As Boolean
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim fSelectFromCourseTableOK As Boolean
  Dim iNextIndex As Integer
  Dim lngFirstFindColumnID As Long
  Dim lngFirstSortColumnID As Long
  Dim sSQL As String
  Dim sTodaysDate As String
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
  Dim objCourseTable As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String

  Screen.MousePointer = vbHourglass
  
  fNoSelect = False

  sOrderString = ""
  sJoinCode = ""
  sColumnList = ""
  sWhereCode = ""
  sTodaysDate = "'" & Replace(Format(Date, "mm/dd/yyyy"), UI.GetSystemDateSeparator, "/") & "'"
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

  mfFirstColumnsMatch = False
  lngFirstFindColumnID = 0
  lngFirstSortColumnID = 0
  mfFirstColumnAscending = True
  miFirstColumnDataType = 0
  fSelectFromCourseTableOK = False
  
  ReDim mavFindColumns(3, 0)
  
  ' Get the Course table object.
  Set objCourseTable = gcoTablePrivileges.FindTableID(glngCourseTableID)
  
  ' Get the default order items from the database.
  Set rsInfo = datGeneral.GetOrderDefinition(glngCourseOrderID)

  fOK = Not (rsInfo.EOF And rsInfo.BOF)
  If Not fOK Then
    COAMsgBox "No order defined for the course table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      sRealSource = gcoTablePrivileges.Item(rsInfo!TableName).RealSource
      
      ' Get the column privileges collection for the given table.
      Set objColumnPrivileges = GetColumnPrivileges(rsInfo!TableName)
      fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect

      ' If this column is from the Training Course table, then check that the user can read
      ' the start date and course title columns in the table.
      If rsInfo!TableID = glngCourseTableID Then
        fColumnOK = objColumnPrivileges.Item(gsCourseStartDateColumnName).AllowSelect And _
          objColumnPrivileges.Item(gsCourseTitleColumnName).AllowSelect
        fSelectFromCourseTableOK = fColumnOK
      End If

      Set objColumnPrivileges = Nothing

      If fColumnOK Then
        ' The column CAN be read from the Course table, or directly from a parent table.
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
        If rsInfo!TableID <> glngCourseTableID Then
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
      Else
        ' The column CANNOT be read from the Course table, or directly from a parent table.
        ' Try to read it from the views on the table.
        
        ' Loop through the views on the column's table, seeing if any have 'read' permission granted on them.
        ReDim asViews(0)
        For Each objTableView In gcoTablePrivileges.Collection
          If (Not objTableView.IsTable) And _
            (objTableView.TableID = rsInfo!TableID) And _
            (objTableView.AllowSelect) Then
              
            sRealSource = gcoTablePrivileges.Item(objTableView.ViewName).RealSource

            ' Get the column permission for the view.
            Set objColumnPrivileges = GetColumnPrivileges(objTableView.ViewName)

            fColumnOK = True
            ' If this column is from the Training Course table, then check that the user can read
            ' the start date and course title columns in the current view on this table.
            If rsInfo!TableID = glngCourseTableID Then
              fColumnOK = (objColumnPrivileges.IsValid(gsCourseStartDateColumnName) And _
                objColumnPrivileges.IsValid(gsCourseTitleColumnName))
              If fColumnOK Then
                fColumnOK = (objColumnPrivileges.Item(gsCourseStartDateColumnName).AllowSelect And _
                  objColumnPrivileges.Item(gsCourseTitleColumnName).AllowSelect)
              End If
            End If
            If fColumnOK Then
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

      rsInfo.MoveNext
    Loop

    ' Inform the user if they do not have permission to see the data.
    If fNoSelect Then
      COAMsgBox "You do not have 'read' permission on all of the columns in the selected order." & _
        vbCrLf & "Only permitted columns will be shown.", vbExclamation, "Security"
    End If
    
    mfFirstColumnsMatch = (lngFirstFindColumnID = lngFirstSortColumnID)

    If Len(sColumnList) > 0 Then
      ' Use the Course table as the base if it can be read.
      If (objCourseTable.AllowSelect) Or _
        (objCourseTable.TableType = tabTopLevel) Then
        
        sSQL = "SELECT " & sColumnList & ", " & _
          objCourseTable.RealSource & ".id" & _
          " FROM " & objCourseTable.RealSource
        
        sRecordCount = "SELECT COUNT(" & objCourseTable.RealSource & ".ID)" & _
          " FROM " & objCourseTable.RealSource
        
        ' Join any other tables and views that are used.
        For iNextIndex = 1 To UBound(alngTableViews, 2)
          If alngTableViews(1, iNextIndex) = 0 Then
            Set objTableView = gcoTablePrivileges.FindTableID(alngTableViews(2, iNextIndex))
          Else
            Set objTableView = gcoTablePrivileges.FindViewID(alngTableViews(2, iNextIndex))
          End If
          
          If objTableView.TableID = glngCourseTableID Then
            ' Join a view of the Course table.
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & objCourseTable.RealSource & ".ID = " & objTableView.RealSource & ".ID"
            sRecordCount = sRecordCount & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & objCourseTable.RealSource & ".ID = " & objTableView.RealSource & ".ID"
            If Not fSelectFromCourseTableOK Then
              sWhereCode = sWhereCode & _
                IIf(Len(sWhereCode) > 0, " OR (", "(") & _
                objCourseTable.RealSource & ".ID IN (SELECT ID FROM " & objTableView.RealSource & _
                  " WHERE " & gsCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "'" & _
                  " AND " & gsCourseCancelDateColumnName & " IS NULL " & _
                  " AND " & gsCourseStartDateColumnName & " >= " & sTodaysDate & "))"
                  
            End If
          Else
            ' Join a parent table/view.
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & objCourseTable.RealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
              " = " & objTableView.RealSource & ".ID"
          End If
          Set objTableView = Nothing
        Next iNextIndex

        sSQL = sSQL & _
          IIf(Len(sWhereCode) > 0, " WHERE " & sWhereCode, "")
        If fSelectFromCourseTableOK Then
          sSQL = sSQL & _
            IIf(Len(sWhereCode) > 0, " AND ", " WHERE ") & _
            objCourseTable.RealSource & "." & gsCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "'" & _
            " AND " & objCourseTable.RealSource & "." & gsCourseStartDateColumnName & " >= " & sTodaysDate & _
            " AND " & gsCourseCancelDateColumnName & " IS NULL "
        End If

        sRecordCount = sRecordCount & _
          IIf(Len(sWhereCode) > 0, " WHERE " & sWhereCode, "")
        If fSelectFromCourseTableOK Then
          sRecordCount = sRecordCount & _
            IIf(Len(sWhereCode) > 0, " AND ", " WHERE ") & _
            objCourseTable.RealSource & "." & gsCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "'" & _
            " AND " & objCourseTable.RealSource & "." & gsCourseStartDateColumnName & " >= " & sTodaysDate & _
            " AND " & gsCourseCancelDateColumnName & " IS NULL "
        End If
          
        ' Tag on the 'order by' code.
        sSQL = sSQL & _
          IIf(Len(sOrderString) > 0, " ORDER BY " & sOrderString, "")
      
        ' Get the required recordset.
        Set mrsCourseRecords = datGeneral.GetPersistentRecords(sSQL, adOpenStatic, adLockReadOnly)
          
        ' Get the recordset's record count.
        Set rsTemp = datGeneral.GetRecords(sRecordCount)
        If (rsTemp.EOF And rsTemp.BOF) Then
          mlngRecordCount = 0
        Else
          mlngRecordCount = rsTemp(0)
        End If
        rsTemp.Close
        Set rsTemp = Nothing

        ' Check we have course records.
        fOK = (mlngRecordCount > 0)
        If Not fOK Then
          COAMsgBox "No course records found.", vbExclamation, Me.Caption
        End If
        
        If fOK Then
          ' Configure the grid.
          ConfigureGrid
        End If
      Else
        ' Unable to read from the course table.
        COAMsgBox "You do not have permission to read the Course table." & _
          vbCrLf & "Unable to display records.", vbExclamation, "Security"
        fOK = False
      End If
    Else
      COAMsgBox "You do not have permission to read any of the columns in the Course table's default order." & _
        vbCrLf & "Unable to display records.", vbExclamation, "Security"
      fOK = False
    End If
  End If

  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  GetCourseRecords = fOK
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error reading Course records.", vbExclamation, Me.Caption
  fOK = False
  Resume TidyUpAndExit
  
End Function





Private Sub ConfigureGrid()
  ' Populate the grid.
  Dim iLoop As Integer
  Dim lngWidth As Long
  
  UI.LockWindow Me.hWnd
  
  lngWidth = 0
  mfFormattingGrid = True
  
  With grdCourses
    .Redraw = False
    .Columns.RemoveAll
    
    For iLoop = 0 To (mrsCourseRecords.Fields.Count - 1)
      .Columns.Add iLoop
      .Columns(iLoop).Name = mrsCourseRecords.Fields(iLoop).Name
      .Columns(iLoop).Visible = (UCase(mrsCourseRecords.Fields(iLoop).Name) <> "ID") And _
        (Left(mrsCourseRecords.Fields(iLoop).Name, 1) <> "?")
      .Columns(iLoop).Caption = RemoveUnderScores(mrsCourseRecords.Fields(iLoop).Name)
      .Columns(iLoop).Alignment = ssCaptionAlignmentLeft
      .Columns(iLoop).CaptionAlignment = ssColCapAlignUseColumnAlignment
    
      ' If the find column is a logic column then set the grid column style to be 'checkbox'.
      If mrsCourseRecords.Fields.Item(iLoop).Type = adBoolean Then
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

  If grdCourses.Rows > grdCourses.VisibleRows Then
    lngWidth = lngWidth + (UI.GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX) + 20
  End If

  Me.Width = lngWidth + 120

  UI.UnlockWindow
  
End Sub




Public Function Initialise(psCourseTitle As String, plngEmployeeID As Long, plngWaitListID As Long) As Boolean
  ' Initialise the form.
  Dim fOK As Boolean
  
  ' Validate the Training Booking module parameters.
  fOK = ValidateParameters
  
  If fOK Then
    msCourseTitle = psCourseTitle
    mlngEmployeeID = plngEmployeeID
    mlngWaitListID = plngWaitListID
    
    ' Get the required course records.
    fOK = GetCourseRecords
  End If
   
  Initialise = fOK
  
End Function


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
  
  If grdCourses.Rows = 0 Then
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
  
  With grdCourses
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




Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide

End Sub

Private Sub cmdSelect_Click()
  grdCourses_DblClick
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
  Hook Me.hWnd, dblFORM_MINWIDTH, dblFORM_MINHEIGHT
  RemoveIcon Me
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
    
  Const dblCOORD_XGAP = 150
  Const dblCOORD_YGAP = 150
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
'  ' Don't let the form get too narrow.
'  If Me.Width < dblFORM_MINWIDTH Then
'    Me.Width = dblFORM_MINWIDTH
'  End If
'
'  ' Don't let the form get too wide.
'  If Me.Width > Screen.Width Then
'    Me.Width = Screen.Width
'  End If
'
'  ' Set the height.
'  If Not mfSizing Then
'    mfSizing = True
'    Me.Height = Screen.Height / 3
'  End If
'
'  ' Don't let the form get too short.
'  If Me.Height < dblFORM_MINHEIGHT Then
'    mfSizing = True
'    Me.Height = dblFORM_MINHEIGHT
'  End If
'
'  ' Don't let the form get too tall.
'  If Me.Height > Screen.Height Then
'    Me.Height = Screen.Height
'  End If
      
  ' Size the grid.
  With grdCourses
    .Width = Me.ScaleWidth - (dblCOORD_XGAP * 2)
    .Height = Me.ScaleHeight - .Top - fraButtons.Height - (2 * dblCOORD_YGAP)
  End With
      
  fraButtons.Top = grdCourses.Top + grdCourses.Height + dblCOORD_YGAP
  fraButtons.Left = Me.ScaleWidth - fraButtons.Width - dblCOORD_XGAP

  ' Stretch the last find column to fit the grid.
  iLastColumnIndex = -1
  iMaxPosition = -1
  With grdCourses
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

End Sub


Private Sub Form_Unload(Cancel As Integer)
  'Tidy things up before unloading
  mrsCourseRecords.Close
  Set mrsCourseRecords = Nothing

  Unhook Me.hWnd
End Sub



Private Sub grdCourses_DblClick()
  Dim fOK As Boolean
  
'  If grdCourses.Row >= 0 And grdCourses.Row < grdCourses.Rows Then
  If grdCourses.SelBookmarks.Count > 0 Then
    ' Get the ID of the selected record.
    mrsCourseRecords.Bookmark = grdCourses.Bookmark
    mlngSelectedRecordID = mrsCourseRecords!ID

    ' Check that current employee has (or will have) satisfied the pre-requisite criteria.
    fOK = TrainingBooking_CheckPreRequisites(mlngSelectedRecordID, mlngEmployeeID)

    ' Check that the current employee is not unavailable for the selected course.
    If fOK Then
      fOK = TrainingBooking_CheckAvailability(mlngSelectedRecordID, mlngEmployeeID)
    End If

    ' Create booking record.
    If fOK Then
      fOK = CreateBooking
    End If

    If fOK Then
      mfCancelled = False
      Me.Hide
    End If
  End If
  
End Sub

Private Sub grdCourses_KeyPress(KeyAscii As Integer)
  Dim lngThistime As Long
  Static sFind As String
  Static lngLastTime As Long
  
  Select Case KeyAscii
    Case vbKeyReturn
      grdCourses_DblClick
    
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


Private Sub grdCourses_UnboundPositionData(StartLocation As Variant, ByVal NumberOfRowsToMove As Long, NewLocation As Variant)
  If IsNull(StartLocation) Then
    If NumberOfRowsToMove = 0 Then
      Exit Sub
    ElseIf NumberOfRowsToMove < 0 Then
      mrsCourseRecords.MoveLast
    Else
      mrsCourseRecords.MoveFirst
    End If
  Else
    mrsCourseRecords.Bookmark = StartLocation
  End If
  
  'JPD 20040803 Fault 9013
  If StartLocation + NumberOfRowsToMove <= 0 Then
    NumberOfRowsToMove = 0
  End If

  mrsCourseRecords.Move NumberOfRowsToMove
  NewLocation = mrsCourseRecords.Bookmark
  
End Sub


Private Sub grdCourses_UnboundReadData(ByVal RowBuf As SSDataWidgets_B.ssRowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
  ' Read the required data from the recordset to the grid.
  Dim iRowIndex As Integer
  Dim iFieldIndex As Integer
  Dim iRowsRead As Integer
  Dim strFormat As String
  
  iRowsRead = 0
  
  ' Do nothing if we a re just formatting the grid,
  ' or if there are no records to display.
  If (mfFormattingGrid) Or (mlngRecordCount = 0) Then Exit Sub
  
  If IsNull(StartLocation) Or (StartLocation = 0) Then
    If ReadPriorRows Then
      If Not mrsCourseRecords.EOF Then
        mrsCourseRecords.MoveLast
      End If
    Else
      If Not mrsCourseRecords.BOF Then
        mrsCourseRecords.MoveFirst
      End If
    End If
  Else
    mrsCourseRecords.Bookmark = StartLocation
    If ReadPriorRows Then
      mrsCourseRecords.MovePrevious
    Else
      mrsCourseRecords.MoveNext
    End If
  End If
  
  ' Read from the row buffer into the grid.
  For iRowIndex = 0 To (RowBuf.RowCount - 1)
    ' Do nothing if the begining of end of the recordset is Met.
    If mrsCourseRecords.BOF Or mrsCourseRecords.EOF Then Exit For
  
    ' Optimize the data read based on the ReadType.
    Select Case RowBuf.ReadType
      Case 0
        For iFieldIndex = 0 To (mrsCourseRecords.Fields.Count - 1)
          Select Case mrsCourseRecords.Fields(iFieldIndex).Type
            Case adDBTimeStamp
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsCourseRecords(iFieldIndex), DateFormat)
            
            Case adNumeric
              ' Are thousand separators used
              strFormat = "0"
              If mavFindColumns(3, iFieldIndex) Then
                strFormat = "#,0"
              End If
              If mavFindColumns(2, iFieldIndex) > 0 Then
                strFormat = strFormat & "." & String(mavFindColumns(2, iFieldIndex), "0")
              End If
              
              RowBuf.Value(iRowIndex, iFieldIndex) = Format(mrsCourseRecords(iFieldIndex), strFormat)
            
            Case Else
              RowBuf.Value(iRowIndex, iFieldIndex) = mrsCourseRecords(iFieldIndex)
          
          End Select
          
        Next iFieldIndex
        RowBuf.Bookmark(iRowIndex) = mrsCourseRecords.Bookmark
  
      Case 1
        RowBuf.Bookmark(iRowIndex) = mrsCourseRecords.Bookmark
  
    End Select
    
    If ReadPriorRows Then
      mrsCourseRecords.MovePrevious
    Else
      mrsCourseRecords.MoveNext
    End If
  
    iRowsRead = iRowsRead + 1
  Next iRowIndex
  
  RowBuf.RowCount = iRowsRead

End Sub



Private Function ValidateParameters() As Boolean
  ' Validate the Training Booking module parameters
  ' used by the 'BookCourse' function.
  Dim iLoop As Integer
  Dim fValid As Boolean
  Dim objColumn As CColumnPrivilege
  Dim objColumns As CColumnPrivileges
  Dim objTable As CTablePrivilege
  Dim alngRelatedColumns() As Long

  ' Check that the Training Booking module is installed.
  fValid = gfTrainingBookingEnabled

  ' Check that the user has the required permissions on the Training Bookings table.
  If fValid Then
    Set objTable = gcoTablePrivileges.Item(gsTrainBookTableName)
    fValid = objTable.AllowInsert
    If Not fValid Then
      COAMsgBox "You do not have 'new' permission on the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
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
          COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
        End If
      End If
      
      ' Check the user has permission to edit all of the related Training Booking columns.
      If fValid Then
        alngRelatedColumns = RelatedColumns
        
        For iLoop = 1 To UBound(alngRelatedColumns, 2)
          Set objColumn = objColumns.FindColumnID(alngRelatedColumns(1, iLoop))
          fValid = Not objColumn Is Nothing
          
          If Not fValid Then
            COAMsgBox "Unable to find all related columns in the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
            Exit For
          Else
            fValid = objColumn.AllowUpdate
            If Not fValid Then
              COAMsgBox "You do not have 'edit' permission on the '" & objColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
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
      COAMsgBox "You do not have 'delete' permission on the '" & gsWaitListTableName & "' table.", vbOKOnly, App.ProductName
    End If
    
    If fValid Then
      Set objColumns = GetColumnPrivileges(gsWaitListTableName)
      
      ' Check the user has permission to edit all of the related Training Booking columns.
      For iLoop = 1 To UBound(alngRelatedColumns, 2)
        Set objColumn = objColumns.FindColumnID(alngRelatedColumns(2, iLoop))
        fValid = Not objColumn Is Nothing
        
        If Not fValid Then
          COAMsgBox "Unable to find all related columns in the '" & gsWaitListTableName & "' table.", vbOKOnly, App.ProductName
          Exit For
        Else
          fValid = objColumn.AllowSelect
          If Not fValid Then
            COAMsgBox "You do not have 'read' permission on the '" & objColumn.ColumnName & "' column in the '" & gsTrainBookTableName & "' table.", vbOKOnly, App.ProductName
            Exit For
          End If
        End If
        Set objColumn = Nothing
      Next iLoop
      
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
  Dim objTrainingBookingTable As CTablePrivilege
  Dim objWaitingListTable As CTablePrivilege
  Dim objTBColumn As CColumnPrivilege
  Dim objWLColumn As CColumnPrivilege
  Dim objTBColumns As CColumnPrivileges
  Dim objWLColumns As CColumnPrivileges
  Dim alngRelatedColumns() As Long
  Dim asAddedColumns() As String
  
  fOK = True
  fInTransaction = False
  
  ' Prompt the user wheter they want to Book the course,
  ' or just Provisionally Book it.
  If gfTrainBookStatus_P Then
    Set frmPrompt = New frmTrainingBookingPrompt
    With frmPrompt
      .Show vbModal
  
      fOK = Not .Cancelled
      fBooked = .Booked
      Unload frmPrompt
    End With
    Set frmPrompt = Nothing
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
      fOK = TrainingBooking_CheckOverbooking(mlngSelectedRecordID, 0)
    End If
  End If

  ' Check that we are not over-lapping another booking.
  If fOK Then
    fOK = TrainingBooking_CheckOverlappedBooking(mlngSelectedRecordID, mlngEmployeeID, 0)
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
      Trim(Str(mlngEmployeeID)) & ", " & _
      Trim(Str(mlngSelectedRecordID)) & ", " & _
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
      " WHERE id = " & Trim(Str(mlngWaitListID))
    
    Screen.MousePointer = vbHourglass

    sErrorMsg = ""
    fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)

    If Not fOK Then
      Screen.MousePointer = vbDefault
      COAMsgBox "Unable to create booking record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly, App.ProductName
      Screen.MousePointer = vbHourglass
      
      gADOCon.RollbackTrans
      fInTransaction = False
    End If
  End If
  
  If fOK Then
    ' Delete the record in the Waiting List table.
    sSQL = "DELETE FROM " & objWaitingListTable.RealSource & _
      " WHERE id = " & Trim(Str(mlngWaitListID))

    sErrorMsg = ""
    fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)

    If Not fOK Then
      Screen.MousePointer = vbDefault
      COAMsgBox "Unable to delete waiting list record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly, App.ProductName
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



