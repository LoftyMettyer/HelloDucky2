VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmTransferCourseBookings 
   Caption         =   "Transfer Bookings"
   ClientHeight    =   4425
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
   HelpContextID   =   1060
   Icon            =   "frmTransferCourseBookings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   5790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   400
      Left            =   3100
      TabIndex        =   3
      Top             =   3900
      Width           =   2600
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Default         =   -1  'True
         Height          =   400
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   1350
         TabIndex        =   2
         Top             =   0
         Width           =   1200
      End
   End
   Begin SSDataWidgets_B.SSDBGrid grdCourses 
      Height          =   3600
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5505
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
      _ExtentX        =   9701
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
Attribute VB_Name = "frmTransferCourseBookings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Employee record variables.
Private mlngCurrentEmployeeID As Long

' Course recordset variables.
Private mrsCourseRecords As New ADODB.Recordset
Private mlngRecordCount As Long

' Course record variables.
Private mlngCourseID As Long
Private msCourseTitle As String

' Course record variables.
Private mlngSelectedRecordID As Long

' Course recordset location variables.
Private mfFirstColumnsMatch As Boolean
Private mfFirstColumnAscending As Boolean
Private miFirstColumnDataType As Integer

' Form handling variables.
Private mfSizing As Boolean
Private mfCancelled As Boolean
Private mfFormattingGrid As Boolean
Private mfPreReqsOverridden As Boolean
Private mfUnavailOverridden  As Boolean
Private mfOverlapOverridden  As Boolean
Private mfErrorTransferring  As Boolean

Private mavFindColumns() As Variant        ' Find columns details

Private Const dblFORM_MINWIDTH = 4000
Private Const dblFORM_MINHEIGHT = 4000

Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property


Private Function CheckAvailability() As Boolean
  ' Check that the selected employee is available for the selected course.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  fOK = True

  ' If no Unavailability table is defined then do nothing.
  If Len(gsUnavailTableName) > 0 Then
    ' Check for the existence of the sp_ASR_TBCheckUnavailability.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckUnavailability')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the delegate is available.
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = "sp_ASR_TBCheckUnavailability"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon

        Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = mlngSelectedRecordID

        Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = mlngCurrentEmployeeID

        Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
        .Parameters.Append pmADO
    
        Set pmADO = Nothing

        cmADO.Execute

        Select Case .Parameters("result").Value
          Case 1    ' Employee unavailable (error).
            fOK = False
            COAMsgBox "Some transferred delegates are unavailable for the selected course." & vbCrLf & _
              "Unable to make the bookings.", vbExclamation + vbOKOnly, App.ProductName
              
          Case 2    ' Employee unavailable (over-rideable by the user).
            If mfUnavailOverridden Then
              fOK = True
            Else
              fOK = (COAMsgBox("Some transferred delegates are unavailable for the selected course." & vbCrLf & _
                "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
              mfUnavailOverridden = fOK
            End If
            
          Case Else ' Employee available.
            fOK = True
        End Select
      
        Set cmADO = Nothing
      End With
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If

  CheckAvailability = fOK

End Function
Private Function CheckPreRequisites() As Boolean
  ' Check that current employee has (or will have) satisfied the pre-requisite criteria.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  fOK = True

  ' If no prerequisite table is defined then do nothing.
  If Len(gsPreReqTableName) > 0 Then
  
    ' Check for the existence of the sp_ASR_TBCheckPreRequisites.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckPreRequisites')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the prerequisites have been met.
      Set cmADO = New ADODB.Command
      With cmADO
        .CommandText = "sp_ASR_TBCheckPreRequisites"
        .CommandType = adCmdStoredProc
        .CommandTimeout = 0
        Set .ActiveConnection = gADOCon

        Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = mlngSelectedRecordID

        Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
        .Parameters.Append pmADO
        pmADO.Value = mlngCurrentEmployeeID

        Set pmADO = .CreateParameter("preReqsMet", adInteger, adParamOutput)
        .Parameters.Append pmADO
    
        Set pmADO = Nothing

        cmADO.Execute

        Select Case .Parameters("preReqsMet").Value
          Case 1    ' Pre-requisites not satisfied (error).
            fOK = False
            COAMsgBox "The pre-requisites for the selected course have not been met." & vbCrLf & _
              "Unable to transfer the booking(s).", vbExclamation + vbOKOnly, App.ProductName
              
          Case 2    ' Pre-requisites not satisfied (over-rideable by the user).
            If mfPreReqsOverridden Then
              fOK = True
            Else
              fOK = (COAMsgBox("The pre-requisites for the selected course have not been met." & vbCrLf & _
                "Do you still want to make the booking(s) ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
              mfPreReqsOverridden = fOK
            End If
          
          Case Else ' Pre-requisites satisfied.
            fOK = True
        End Select
      
        Set cmADO = Nothing
      End With
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
  CheckPreRequisites = fOK

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





Private Sub TransferBookings()
  ' Create the booking record.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim objTBTable As CTablePrivilege
  Dim sErrorMsg As String

  Screen.MousePointer = vbHourglass
    
  Set objTBTable = gcoTablePrivileges.Item(gsTrainBookTableName)
  
  ' Create the new booking records.
  sSQL = "INSERT INTO " & objTBTable.RealSource & _
    " (" & gsTrainBookStatusColumnName & ", " & _
    "id_" & Trim(Str(glngEmployeeTableID)) & ", " & _
    "id_" & Trim(Str(glngCourseTableID)) & ")" & _
    " (SELECT " & _
    gsTrainBookStatusColumnName & ", " & _
    "id_" & Trim(Str(glngEmployeeTableID)) & ", " & _
    Trim(Str(mlngSelectedRecordID)) & _
    " FROM " & objTBTable.RealSource & _
    " WHERE id_" & glngCourseTableID & " = " & Trim(Str(mlngCourseID))
    
  If gfCourseTransferProvisionals Then
    sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
      " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P'))"
  Else
    sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B')"
  End If

  sErrorMsg = ""
  fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
  
  Screen.MousePointer = vbDefault

  If Not fOK Then
    mfErrorTransferring = True
    COAMsgBox "Unable to transfer the bookings." & vbCrLf & vbCrLf & sErrorMsg, vbExclamation + vbOKOnly, App.ProductName
  End If

  Set objTBTable = Nothing
  
End Sub

Private Function CheckOverlappedBooking() As Boolean
  ' Check that the selected course is not already fully booked.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  fOK = True

  ' Check for the existence of the sp_ASR_TBCheckOverlappedBooking.
  sSQL = "SELECT COUNT(*) AS objectCount" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('sp_ASR_TBCheckOverlappedBooking')" & _
    "     AND sysstat & 0xf = 4"
  Set rsInfo = datGeneral.GetRecords(sSQL)
  
  If rsInfo!objectCount > 0 Then
    ' If it exists then run it to see if the prerequisites have been met.
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASR_TBCheckOverlappedBooking"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = mlngSelectedRecordID

      Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = mlngCurrentEmployeeID

      Set pmADO = .CreateParameter("bookingRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = 0

      Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
      .Parameters.Append pmADO
  
      Set pmADO = Nothing

      cmADO.Execute

      Select Case .Parameters("result").Value
        Case 1    ' Overlapped booking (error).
          fOK = False
          COAMsgBox "A delegate is already booked on a course that overlaps with the selected course." & vbCrLf & _
            "Unable to transfer the booking.", vbExclamation + vbOKOnly, App.ProductName
            
        Case 2    ' Overlapped booking (over-rideable by the user).
          If mfOverlapOverridden Then
            fOK = True
          Else
            fOK = (COAMsgBox("A delegate is already booked on a course that overlaps with the selected course." & vbCrLf & _
              "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
            mfOverlapOverridden = fOK
          End If
          
        Case Else ' Course NOT fully booked.
          fOK = True
      End Select
    
      Set cmADO = Nothing
    End With
  End If
  
  rsInfo.Close
  Set rsInfo = Nothing

  CheckOverlappedBooking = fOK
  
End Function





Private Function CheckOverbooking(plngNumberBooked As Long) As Boolean
  ' Check that the selected course is not already fully booked.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  fOK = True

  ' Check for the existence of the sp_ASR_TBCheckOverbooking.
  sSQL = "SELECT COUNT(*) AS objectCount" & _
    "   FROM sysobjects" & _
    "   WHERE id = object_id('sp_ASR_TBCheckOverbooking')" & _
    "     AND sysstat & 0xf = 4"
  Set rsInfo = datGeneral.GetRecords(sSQL)
    
  If rsInfo!objectCount > 0 Then
    ' If it exists then run it to see if the prerequisites have been met.
    Set cmADO = New ADODB.Command
    With cmADO
      .CommandText = "sp_ASR_TBCheckOverbooking"
      .CommandType = adCmdStoredProc
      .CommandTimeout = 0
      Set .ActiveConnection = gADOCon

      Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = mlngSelectedRecordID

      Set pmADO = .CreateParameter("bookingID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = 0

      Set pmADO = .CreateParameter("newBookings", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngNumberBooked

      Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
      .Parameters.Append pmADO
  
      Set pmADO = Nothing

      cmADO.Execute

      Select Case .Parameters("result").Value
        Case 1    ' Course fully booked (error).
          fOK = False
          COAMsgBox "The selected course is already fully booked." & vbCrLf & _
            "Unable to transfer the bookings.", vbExclamation + vbOKOnly, App.ProductName
            
        Case 2    ' Course fully booked (over-rideable by the user).
          fOK = (COAMsgBox("The selected course is already fully booked." & vbCrLf & _
            "Do you still want to transfer the bookings ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
                  
        Case Else ' Course NOT fully booked.
          fOK = True
      End Select
    
      Set cmADO = Nothing
    End With
  End If
    
  rsInfo.Close
  Set rsInfo = Nothing

  CheckOverbooking = fOK
  
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

Public Function Initialise(plngCourseID As Long, psCourseTitle As String) As Boolean
  ' Initialise the form.
  Dim fOK As Boolean
  
  fOK = ValidateParameters
  
  If fOK Then
    mlngCourseID = plngCourseID
    msCourseTitle = psCourseTitle
    
    ' Get the required course records.
    fOK = GetCourseRecords
  End If
  
  mfErrorTransferring = False
  
  Initialise = fOK
  
End Function
Private Function GetCourseRecords() As Boolean
  ' Construct a recordset of the courses that match the required title, and have a start date
  ' on or after the system date.
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
  Set rsInfo = datGeneral.GetOrderDefinition(objCourseTable.DefaultOrderID)

  fOK = Not (rsInfo.EOF And rsInfo.BOF)
  If Not fOK Then
    COAMsgBox "No default order defined for the course table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      ' Get the column privileges collection for the given table.
      sRealSource = gcoTablePrivileges.Item(rsInfo!TableName).RealSource
      
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
                  " AND " & gsCourseStartDateColumnName & " >= " & sTodaysDate & _
                  " AND id <> " & Trim(Str(mlngCourseID)) & "))"
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
            " AND " & gsCourseCancelDateColumnName & " IS NULL " & _
            " AND " & objCourseTable.RealSource & ".id <> " & Trim(Str(mlngCourseID))
        End If

        sRecordCount = sRecordCount & _
          IIf(Len(sWhereCode) > 0, " WHERE " & sWhereCode, "")
        If fSelectFromCourseTableOK Then
          sRecordCount = sRecordCount & _
            IIf(Len(sWhereCode) > 0, " AND ", " WHERE ") & _
            objCourseTable.RealSource & "." & gsCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "'" & _
            " AND " & objCourseTable.RealSource & "." & gsCourseStartDateColumnName & " >= " & sTodaysDate & _
            " AND " & gsCourseCancelDateColumnName & " IS NULL " & _
            " AND " & objCourseTable.RealSource & ".id <> " & Trim(Str(mlngCourseID))
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
Private Function ValidateParameters() As Boolean
  Dim fValid As Boolean
  Dim objColumns As CColumnPrivileges

  ' Check that the Training Booking module is installed.
  fValid = gfTrainingBookingEnabled

  ' Validate the required Training Bookings table parameters.
  If fValid Then
    ' Check that the user has permission to insert records into the Training Bookings table.
    fValid = gcoTablePrivileges.Item(gsTrainBookTableName).AllowInsert
    If Not fValid Then
      COAMsgBox "You do not have 'new' permission on the '" & gsTrainBookTableName & "' table.", vbExclamation + vbOKOnly, App.ProductName
    End If
  End If

  If fValid Then
    ' Get the column privileges for the Training Bookings table.
    Set objColumns = GetColumnPrivileges(gsTrainBookTableName)

'''    ' Check that the user has permission to edit the Training Bookings Course Title column.
'''    fValid = objColumns.Item(gsTrainBookCourseTitleName).AllowUpdate
'''    If Not fValid Then
'''      COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCourseTitleName & "' column.", vbExclamation + vbOKOnly, App.ProductName
'''    End If

    If fValid Then
      ' Check that the user has permission to edit the Training Bookings Status column.
      fValid = objColumns.Item(gsTrainBookStatusColumnName).AllowUpdate
      If Not fValid Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbExclamation + vbOKOnly, App.ProductName
      End If
    End If

    Set objColumns = Nothing
  End If
  
  ValidateParameters = fValid

End Function


Private Sub cmdCancel_Click()
  mfCancelled = True
  Me.Hide

End Sub


Private Sub cmdSelect_Click()
  grdCourses_DblClick

End Sub


Private Sub Form_Load()
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
  Dim sSQL As String
  Dim lngNumberBooked As Long
  Dim rsBookings As ADODB.Recordset
  Dim objTBTable As CTablePrivilege
  
  fOK = True
  
  If (grdCourses.Row >= 0) And (grdCourses.Row < grdCourses.Rows) Then
    ' Get the ID of the selected record.
    mrsCourseRecords.Bookmark = grdCourses.Bookmark
    mlngSelectedRecordID = mrsCourseRecords!ID

    Set objTBTable = gcoTablePrivileges.Item(gsTrainBookTableName)
    
    ' Get a recordset of the the bookings to be transferred from the current course.
    sSQL = "SELECT id_" & Trim(Str(glngEmployeeTableID)) & ", " & gsTrainBookStatusColumnName & _
      " FROM " & objTBTable.RealSource & _
      " WHERE id_" & Trim(Str(glngCourseTableID)) & " = " & Trim(Str(mlngCourseID))
    If gfCourseTransferProvisionals Then
      sSQL = sSQL & " AND (LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'" & _
        " OR LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'P')"
    Else
      sSQL = sSQL & " AND LEFT(UPPER(" & gsTrainBookStatusColumnName & "), 1) = 'B'"
    End If
    Set rsBookings = datGeneral.GetRecords(sSQL)
    
    Set objTBTable = Nothing
    
    mfPreReqsOverridden = False
    mfUnavailOverridden = False
    mfOverlapOverridden = False
    lngNumberBooked = 0
    
    With rsBookings
      Do While (Not .EOF) And fOK
        mlngCurrentEmployeeID = .Fields("id_" & Trim(Str(glngEmployeeTableID)))
        
        ' Check that current employee has (or will have) satisfied the pre-requisite criteria.
        fOK = CheckPreRequisites

        ' Check that the current employee is not unavailable for the selected course.
        If fOK Then
          fOK = CheckAvailability
        End If
    
        ' Check that the current employee is not unavailable for the selected course.
        If fOK Then
          fOK = CheckOverlappedBooking
        End If
    
        ' Total the number of delegates booked.
        If (UCase(Trim(.Fields(gsTrainBookStatusColumnName))) = "B") Or gfCourseIncludeProvisionals Then
          lngNumberBooked = lngNumberBooked + 1
        End If
        
        .MoveNext
      Loop
      
      .Close
    End With
    Set rsBookings = Nothing
    
    If fOK Then
      fOK = CheckOverbooking(lngNumberBooked)
    End If
      
    ' Transfer the bookings.
    If fOK Then
      TransferBookings
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
  ' ot if there a re no records to display.
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





Public Property Get ErrorTransferring() As Boolean
  ErrorTransferring = mfErrorTransferring
  
End Property


