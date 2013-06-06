VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBulkBooking 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bulk Booking"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1012
   Icon            =   "frmBulkBooking.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboBookingStatus 
      Height          =   315
      ItemData        =   "frmBulkBooking.frx":000C
      Left            =   1665
      List            =   "frmBulkBooking.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2925
   End
   Begin VB.Frame fraAddRemoveButtons 
      BorderStyle     =   0  'None
      Caption         =   "fraButtons"
      Height          =   2950
      Left            =   4800
      TabIndex        =   11
      Top             =   650
      Width           =   1515
      Begin VB.CommandButton cmdPicklistAdd 
         Caption         =   "&Picklist Add..."
         Height          =   400
         Left            =   0
         TabIndex        =   5
         Top             =   1500
         Width           =   1425
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&Add..."
         Height          =   450
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1425
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   400
         Left            =   0
         TabIndex        =   6
         Top             =   2000
         Width           =   1425
      End
      Begin VB.CommandButton cmdDeleteAll 
         Caption         =   "Re&move All"
         Enabled         =   0   'False
         Height          =   400
         Left            =   0
         TabIndex        =   7
         Top             =   2500
         Width           =   1425
      End
      Begin VB.CommandButton cmdAddAll 
         Caption         =   "A&dd All"
         Height          =   400
         Left            =   0
         TabIndex        =   3
         Top             =   500
         Width           =   1425
      End
      Begin VB.CommandButton cmdAddFilter 
         Caption         =   "&Filtered Add..."
         Height          =   400
         Left            =   0
         TabIndex        =   4
         Top             =   1000
         Width           =   1425
      End
   End
   Begin VB.Frame fraOKCancelButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   900
      Left            =   4800
      TabIndex        =   10
      Top             =   3750
      Width           =   1515
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   400
         Left            =   0
         TabIndex        =   9
         Top             =   500
         Width           =   1425
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   400
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1425
      End
   End
   Begin MSComctlLib.ListView lvRecords 
      Height          =   4000
      Left            =   150
      TabIndex        =   1
      Top             =   650
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblBookingStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Status :"
      Height          =   195
      Left            =   150
      TabIndex        =   12
      Top             =   210
      Width           =   1440
   End
End
Attribute VB_Name = "frmBulkBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Course record variables.
Private mlngCourseID As Long
Private msCourseTitle As String

' Form handling variables.
Private mfCancelled As Boolean
Private mfLoading As Boolean
Private mfSizing As Boolean

' Delegate record variables.
Private mSelectDelegateSQL As String
Private mavSelectedRecords() As Variant
Private msDelegateRealSource As String

Const giSTATUS_BOOKED = 1
Const giSTATUS_PROVISIONAL = 2

Const giCHECK_PASSED = 1
Const giCHECK_FAILED = 2
Const giCHECK_FAILEDQUERY = 3

' Array holding the User Defined functions that are needed for this report
Private mastrUDFsRequired() As String
Private mavOrderDefinition() As Variant

Private Const dblFORM_MINWIDTH = 6285
Private Const dblFORM_MINHEIGHT = 5400

Private Function ValidateParameters() As Boolean
  ' Validate the training Booking module parameters.
  Dim fValid As Boolean
  Dim objTBColumnPrivileges As CColumnPrivileges
  
  ' Check that the Training Booking module is installed.
  fValid = gfTrainingBookingEnabled
  
  ' Validate the required Training Bookings table parameters.
  If fValid Then
    ' Get the column privileges for the Training Bookings table.
    Set objTBColumnPrivileges = GetColumnPrivileges(gsTrainBookTableName)

'''    ' Check that the user has permission to update the Training Bookings Course Title column.
'''    fValid = objTBColumnPrivileges.Item(gsTrainBookCourseTitleName).AllowUpdate
'''    If Not fValid Then
'''      COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookCourseTitleName & "' column.", vbOKOnly, App.ProductName
'''    End If

    If fValid Then
      ' Check that the user has permission to update the Training Bookings Status column.
      fValid = objTBColumnPrivileges.Item(gsTrainBookStatusColumnName).AllowUpdate
      If Not fValid Then
        COAMsgBox "You do not have 'edit' permission on the '" & gsTrainBookStatusColumnName & "' column.", vbOKOnly + vbInformation, App.ProductName
      End If
    End If

    Set objTBColumnPrivileges = Nothing
  End If

  ValidateParameters = fValid
  
End Function

Public Function Initialise(plngCourseID As Long, pobjCourseTableView As CTablePrivilege) As Boolean
  ' Initialise the Bulk Booking form.
  Dim fOK As Boolean
  Dim sSQL As String
  Dim rsTemp As ADODB.Recordset
  Dim objColumns As CColumnPrivileges
  
  ' Validate the Training Booking module parameters.
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
    mlngCourseID = plngCourseID
    
    ' Get the selected course's title.
    sSQL = "SELECT " & gsCourseTitleColumnName & _
      " FROM " & pobjCourseTableView.RealSource & _
      " WHERE id = " & Trim(Str(mlngCourseID))
    Set rsTemp = datGeneral.GetRecords(sSQL)
    fOK = Not (rsTemp.BOF And rsTemp.EOF)
  
    If fOK Then
      msCourseTitle = rsTemp.Fields(gsCourseTitleColumnName)
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
  End If
  
  If fOK Then
    ' Initialise the booking status combo.
    cboBookingStatus_Initialise
    
    ' Initialise the employee list view.
    fOK = lvRecords_Configure
  End If
  
  Initialise = fOK

End Function

Private Function lvRecords_Configure() As Boolean
  ' Construct a recordset of the delegates that have the given course title
  ' on their Waiting List.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fNoSelect As Boolean
  Dim fColumnOK As Boolean
  Dim fFound As Boolean
  Dim iNextIndex As Integer
  Dim sSQL As String
  Dim sRealSource As String
  Dim sColumnCode As String
  Dim sColumnList As String
  Dim sJoinCode As String
  Dim sWhereCode As String
  Dim objColumnPrivileges As CColumnPrivileges
  Dim rsInfo As Recordset
  Dim objTableView As CTablePrivilege
  Dim objDelegateTable As CTablePrivilege
  Dim alngTableViews() As Long
  Dim asViews() As String
    
  Screen.MousePointer = vbHourglass
  
  fNoSelect = False
  
  sJoinCode = ""
  sColumnList = ""
  sWhereCode = ""

  ' Initialise the order definition array.
  ' Index 1 = column name.
  ' Index 2 = table name.
  ' Index 3 = table ID.
  ' Index 4 = column size.
  ' Index 5 = decimals.
  ' Index 6 = uses separator.
  ReDim mavOrderDefinition(6, 0)
  
  ' Clear the listview headers.
  lvRecords.ColumnHeaders.Clear
  
  ' Dimension an array of tables/views joined to the base table/view.
  ' Column 1 = 0 if this row is for a table, 1 if it is for a view.
  ' Column 2 = table/view ID.
  ReDim alngTableViews(2, 0)

  ' Get the Delegate table object.
  Set objDelegateTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)
  msDelegateRealSource = objDelegateTable.RealSource
  
  ' Get the default order items from the database.
  Set rsInfo = datGeneral.GetOrderDefinition(objDelegateTable.DefaultOrderID)

  fOK = Not (rsInfo.EOF And rsInfo.BOF)
  If Not fOK Then
    COAMsgBox "No default order defined for the delegate table." & _
      vbCrLf & "Unable to display the records.", vbExclamation, "Security"
  Else
    ' Check the user's privilieges on the order columns.
    Do While Not rsInfo.EOF
      If rsInfo!Type = "F" Then
        ' Get the column privileges collection for the given table.
        sRealSource = gcoTablePrivileges.Item(rsInfo!TableName).RealSource
        
        Set objColumnPrivileges = GetColumnPrivileges(rsInfo!TableName)
        fColumnOK = objColumnPrivileges.Item(rsInfo!ColumnName).AllowSelect
        Set objColumnPrivileges = Nothing
  
        If fColumnOK Then
          ' The column CAN be read from the Delegate table, or directly from a parent table.
          ' Add the column to the column list.
          sColumnList = sColumnList & _
            IIf(Len(sColumnList) > 0, ", ", "") & _
            sRealSource & "." & Trim(rsInfo!ColumnName)
          
          ' Add the column name to the listview headers.
          lvRecords.ColumnHeaders.Add , , RemoveUnderScores(rsInfo!ColumnName)
          
          ' Add the column to the order definition array.
          iNextIndex = UBound(mavOrderDefinition, 2) + 1
          ReDim Preserve mavOrderDefinition(6, iNextIndex)
          mavOrderDefinition(1, iNextIndex) = Trim(rsInfo!ColumnName)
          mavOrderDefinition(2, iNextIndex) = Trim(rsInfo!TableName)
          mavOrderDefinition(3, iNextIndex) = rsInfo!TableID
          mavOrderDefinition(4, iNextIndex) = datGeneral.GetDataSize(rsInfo!ColumnID)
          mavOrderDefinition(5, iNextIndex) = datGeneral.GetDecimalsSize(rsInfo!ColumnID)
          mavOrderDefinition(6, iNextIndex) = datGeneral.DoesColumnUseSeparators(rsInfo!ColumnID)
          
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
        Else
          ' The column CANNOT be read from the Delegate table, or directly from a parent table.
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
              ' Add the column name to the listview headers.
              lvRecords.ColumnHeaders.Add , , RemoveUnderScores(rsInfo!ColumnName)
              
              ' Add the column to the order definition array.
              iNextIndex = UBound(mavOrderDefinition, 2) + 1
              ReDim Preserve mavOrderDefinition(6, iNextIndex)
              mavOrderDefinition(1, iNextIndex) = Trim(rsInfo!ColumnName)
              mavOrderDefinition(2, iNextIndex) = Trim(rsInfo!TableName)
              mavOrderDefinition(3, iNextIndex) = rsInfo!TableID
              mavOrderDefinition(4, iNextIndex) = datGeneral.GetDataSize(rsInfo!ColumnID)
              mavOrderDefinition(5, iNextIndex) = datGeneral.GetDecimalsSize(rsInfo!ColumnID)
              mavOrderDefinition(6, iNextIndex) = datGeneral.DoesColumnUseSeparators(rsInfo!ColumnID)
              
              sColumnCode = sColumnCode & _
                " ELSE NULL" & _
                " END AS " & _
                rsInfo!ColumnName
                
              sColumnList = sColumnList & _
                IIf(Len(sColumnList) > 0, ", ", "") & _
                sColumnCode
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
    
    If Len(sColumnList) > 0 Then
      ' Use the Delegate table as the base if it can be read.
      If (objDelegateTable.AllowSelect) Or _
        (objDelegateTable.TableType = tabTopLevel) Then
        
        sSQL = "SELECT " & sColumnList & ", " & _
          objDelegateTable.RealSource & ".id" & _
          " FROM " & objDelegateTable.RealSource
        
        ' Join any other tables and views that are used.
        For iNextIndex = 1 To UBound(alngTableViews, 2)
          If alngTableViews(1, iNextIndex) = 0 Then
            Set objTableView = gcoTablePrivileges.FindTableID(alngTableViews(2, iNextIndex))
          Else
            Set objTableView = gcoTablePrivileges.FindViewID(alngTableViews(2, iNextIndex))
          End If
          
          If objTableView.TableID = glngEmployeeTableID Then
            ' Join a view of the Delegate table.
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & objDelegateTable.RealSource & ".ID = " & objTableView.RealSource & ".ID"
            If Not objDelegateTable.AllowSelect Then
              sWhereCode = sWhereCode & _
                IIf(Len(sWhereCode) > 0, " OR (", "(") & _
                objDelegateTable.RealSource & ".ID IN (SELECT ID FROM " & objTableView.RealSource & "))"
            End If
          Else
            ' Join a parent table/view.
            sSQL = sSQL & _
              " LEFT OUTER JOIN " & objTableView.RealSource & _
              " ON " & objDelegateTable.RealSource & ".ID_" & Trim(Str(objTableView.TableID)) & _
              " = " & objTableView.RealSource & ".ID"
          End If
          Set objTableView = Nothing
        Next iNextIndex

        sSQL = sSQL & _
          IIf(Len(sWhereCode) > 0, " WHERE (" & sWhereCode & ")", "")
        
        mSelectDelegateSQL = sSQL
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

  rsInfo.Close
  Set rsInfo = Nothing

TidyUpAndExit:
  Screen.MousePointer = vbDefault
  lvRecords_Configure = fOK
  Exit Function
  
ErrorTrap:
  COAMsgBox "Error reading Delegate records.", vbExclamation, Me.Caption
  fOK = False
  Resume TidyUpAndExit

End Function

Private Function CheckAvailability() As Boolean
  ' Check that the selected employee is available for the selected course.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iReply As Integer
  Dim iPassedCount As Integer
  Dim iFailedCount As Integer
  Dim iFailedQueryCount As Integer
  Dim lngSelectedRecordID As Long
  Dim strCheckFailedPeople As String
  Dim strCheckFailedOverrideable As String
  Dim sMessage As String
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter

  strCheckFailedPeople = ""
  strCheckFailedOverrideable = ""

  fOK = True

  ' If no Unavailability table is defined then do nothing.
  If Len(gsUnavailTableName) > 0 Then
    iPassedCount = lvRecords.ListItems.Count
    iFailedCount = 0
    iFailedQueryCount = 0
  
    ' Check for the existence of the sp_ASR_TBCheckUnavailability.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckUnavailability')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the delegate is available.
      For iLoop = 1 To UBound(mavSelectedRecords, 2)
        lngSelectedRecordID = mavSelectedRecords(1, iLoop)
  
        Set cmADO = New ADODB.Command
        With cmADO
          .CommandText = "sp_ASR_TBCheckUnavailability"
          .CommandType = adCmdStoredProc
          .CommandTimeout = 0
          Set .ActiveConnection = gADOCon
  
          Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
          .Parameters.Append pmADO
          pmADO.Value = mlngCourseID
  
          Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
          .Parameters.Append pmADO
          pmADO.Value = lngSelectedRecordID
  
          Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
          .Parameters.Append pmADO
      
          Set pmADO = Nothing
  
          cmADO.Execute
  
          Select Case .Parameters("result").Value
            Case 1    ' Employee unavailable (error).
              mavSelectedRecords(3, iLoop) = giCHECK_FAILED
              iFailedCount = iFailedCount + 1
              iPassedCount = iPassedCount - 1
              
              'NHRD17012007 Fault 9017
              If Len(strCheckFailedPeople) > 0 Then
                strCheckFailedPeople = strCheckFailedPeople & ", " & "'" & (mavSelectedRecords(4, iLoop)) & "'"
              Else
                strCheckFailedPeople = "'" & (mavSelectedRecords(4, iLoop)) & "'"
              End If
               
            Case 2    ' Employee unavailable (over-rideable by the user).
              mavSelectedRecords(3, iLoop) = giCHECK_FAILEDQUERY
              iFailedQueryCount = iFailedQueryCount + 1
              iPassedCount = iPassedCount - 1
            
              'NHRD17012007 Fault 9017
              If Len(strCheckFailedOverrideable) > 0 Then
                strCheckFailedOverrideable = strCheckFailedOverrideable & ", " & "'" & (mavSelectedRecords(4, iLoop)) & "'"
              Else
                strCheckFailedOverrideable = "'" & (mavSelectedRecords(4, iLoop)) & "'"
              End If

            Case Else ' Employee available.
          End Select
        
          Set cmADO = Nothing
        End With
      Next iLoop

      ' Tell the user of any availability failures.
      If iFailedCount > 0 Then
        'NHRD17012007 Fault 9017
        sMessage = strCheckFailedPeople & IIf(iFailedCount = 1, " is", " are") & _
          " unavailable for the current course." & vbCrLf & _
          IIf(iFailedCount = lvRecords.ListItems.Count, "Unable to book the selected delegates.", "Continue booking the selected delegates that are available ?")
      
        fOK = (COAMsgBox(sMessage, IIf(iFailedCount = lvRecords.ListItems.Count, vbOKOnly + vbInformation, vbOKCancel), App.ProductName) = vbOK) And _
          (iFailedCount < lvRecords.ListItems.Count)
      End If
      
      If fOK Then
        If iFailedQueryCount > 0 Then
          'NHRD17012007 Fault 9017
          sMessage = strCheckFailedOverrideable & IIf(iFailedQueryCount = 1, " is", " are") & _
            " unavailable for the current course." & vbCrLf & _
            "Do you still want to book " & IIf(iFailedQueryCount = 1, "this delegate", "these delegates") & " on the course ?"
      
          iReply = COAMsgBox(sMessage, vbYesNo + vbQuestion, App.ProductName)
      
          ' Modify the check status with respect to the user's choice.
          For iLoop = 1 To UBound(mavSelectedRecords, 2)
            If mavSelectedRecords(3, iLoop) = giCHECK_FAILEDQUERY Then
              mavSelectedRecords(3, iLoop) = IIf(iReply = vbYes, giCHECK_PASSED, giCHECK_FAILED)
            End If
          Next iLoop
      
          ' Stop the bulk booking.
          fOK = (iReply <> vbNo)
        End If
      End If
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If

  CheckAvailability = fOK

End Function

Private Function CheckPreRequisites() As Boolean
  ' Check that the selected employees have (or will have) satisfied the pre-requisite criteria.
  ' Check that selected employee has (or will have) satisfied the pre-requisite criteria.
  Dim fOK As Boolean
  Dim iLoop As Integer
  Dim iReply As Integer
  Dim iPassedCount As Integer
  Dim iFailedCount As Integer
  Dim iFailedQueryCount As Integer
  Dim lngSelectedRecordID As Long
  Dim sMessage As String
  Dim strCheckFailedPeople As String
  Dim strCheckFailedOverrideable As String
  Dim sSQL As String
  Dim rsInfo As ADODB.Recordset
  Dim cmADO As ADODB.Command
  Dim pmADO As ADODB.Parameter
  
  strCheckFailedPeople = ""
  strCheckFailedOverrideable = ""
  
  fOK = True
  
  ' If no prerequisite table is defined then do nothing.
  If Len(gsPreReqTableName) > 0 Then
  
    iPassedCount = lvRecords.ListItems.Count
    iFailedCount = 0
    iFailedQueryCount = 0
    
    ' Check for the existence of the sp_ASR_TBCheckPreRequisites.
    sSQL = "SELECT COUNT(*) AS objectCount" & _
      "   FROM sysobjects" & _
      "   WHERE id = object_id('sp_ASR_TBCheckPreRequisites')" & _
      "     AND sysstat & 0xf = 4"
    Set rsInfo = datGeneral.GetRecords(sSQL)
    
    If rsInfo!objectCount > 0 Then
      ' If it exists then run it to see if the prerequisites have been met.
      For iLoop = 1 To UBound(mavSelectedRecords, 2)
        lngSelectedRecordID = mavSelectedRecords(1, iLoop)
      
        Set cmADO = New ADODB.Command
        With cmADO
          .CommandText = "sp_ASR_TBCheckPreRequisites"
          .CommandType = adCmdStoredProc
          .CommandTimeout = 0
          Set .ActiveConnection = gADOCon
    
          Set pmADO = .CreateParameter("courseRecordID", adInteger, adParamInput)
          .Parameters.Append pmADO
          pmADO.Value = mlngCourseID
  
          Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
          .Parameters.Append pmADO
          pmADO.Value = lngSelectedRecordID
  
          Set pmADO = .CreateParameter("preReqsMet", adInteger, adParamOutput)
          .Parameters.Append pmADO
      
          Set pmADO = Nothing
  
          cmADO.Execute
          
          Select Case .Parameters("preReqsMet").Value
            Case 1    ' Pre-requisites not satisfied (error).
              mavSelectedRecords(2, iLoop) = giCHECK_FAILED
              iFailedCount = iFailedCount + 1
              iPassedCount = iPassedCount - 1
            
              If Len(strCheckFailedPeople) > 0 Then
                strCheckFailedPeople = strCheckFailedPeople & ", " & "'" & (mavSelectedRecords(4, iLoop)) & "'"
              Else
                strCheckFailedPeople = "'" & (mavSelectedRecords(4, iLoop)) & "'"
              End If
              
            Case 2    ' Pre-requisites not satisfied (over-rideable by the user).
              mavSelectedRecords(2, iLoop) = giCHECK_FAILEDQUERY
              iFailedQueryCount = iFailedQueryCount + 1
              iPassedCount = iPassedCount - 1
          
              If Len(strCheckFailedOverrideable) > 0 Then
                strCheckFailedOverrideable = strCheckFailedOverrideable & ", " & "'" & (mavSelectedRecords(4, iLoop)) & "'"
              Else
                strCheckFailedOverrideable = "'" & (mavSelectedRecords(4, iLoop)) & "'"
              End If

            Case Else ' Pre-requisites satisfied.
          End Select
          Set cmADO = Nothing
        End With
      Next iLoop
    
      ' Tell the user of any pre-requisite failures.
      If iFailedCount > 0 Then
        'NHRD17012007 Fault 9017 replace numbers with meaningful names
        sMessage = strCheckFailedPeople & _
          " failed to meet the pre-requisites for the current course." & vbCrLf & _
          IIf(iFailedCount = lvRecords.ListItems.Count, "Unable to book the selected delegates.", "Continue booking the selected delegates that met the pre-requisite criteria ?")
      
        fOK = (COAMsgBox(sMessage, IIf(iFailedCount = lvRecords.ListItems.Count, vbOKOnly + vbQuestion, vbOKCancel + vbQuestion), App.ProductName) = vbOK) And _
          (iFailedCount < lvRecords.ListItems.Count)
      End If
          
      If fOK Then
        If iFailedQueryCount > 0 Then
          'NHRD17012007 Fault 9017 replace numbers with meaningful names
          sMessage = strCheckFailedOverrideable & _
            " failed to meet the pre-requisites for the current course." & vbCrLf & _
            "Do you still want to book " & IIf(iFailedQueryCount = 1, "this delegate", "these delegates") & " on the course ?" & vbCrLf & vbCrLf & _
            "Click Yes to continue or click No to go back and remove " & IIf(iFailedQueryCount = 1, "this delegate", "these delegates") & " from the Bulk Booking list ?"
      
          iReply = COAMsgBox(sMessage, vbYesNo + vbQuestion, App.ProductName)
          
          ' Modify the check status with respect to the user's choice.
          For iLoop = 1 To UBound(mavSelectedRecords, 2)
            If mavSelectedRecords(2, iLoop) = giCHECK_FAILEDQUERY Then
              mavSelectedRecords(2, iLoop) = IIf(iReply = vbYes, giCHECK_PASSED, giCHECK_FAILED)
            End If
          Next iLoop
      
          ' Stop the bulk booking.
          fOK = (iReply <> vbNo)
        End If
      End If
    End If
    
    rsInfo.Close
    Set rsInfo = Nothing
  End If
  
  CheckPreRequisites = fOK
  
End Function
Private Function CreateBookings() As Boolean
  ' Create the booking record.
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  Dim fOverBooked As Boolean
  Dim fBooked As Boolean
  Dim fInTransaction As Boolean
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim iBulkBookingCount As Integer
  Dim iNumberBooked As Integer
  Dim sSQL As String
  Dim sErrorMsg As String
  Dim sDelegateDescription As String
  Dim objTrainingBookingTable As CTablePrivilege
  Dim objWaitingListTable As CTablePrivilege
  Dim objWLColumns As CColumnPrivileges
  Dim iBookedCount As Integer
  
  fOK = True
  fInTransaction = False
  fOverBooked = False
  iBookedCount = 0
  
  Set objTrainingBookingTable = gcoTablePrivileges.Item(gsTrainBookTableName)
  Set objWaitingListTable = gcoTablePrivileges.Item(gsWaitListTableName)
  Screen.MousePointer = vbHourglass

  With cboBookingStatus
    fBooked = (.ItemData(.ListIndex) = giSTATUS_BOOKED)
  End With
    
  ' Get the number of bookings being made.
  iBulkBookingCount = 0
  iNumberBooked = 0
  For iLoop = 1 To UBound(mavSelectedRecords, 2)
    ' Only book delegates who have passed the pre-requisite and
    ' availability criteria.
    If (mavSelectedRecords(2, iLoop) = giCHECK_PASSED) And _
      (mavSelectedRecords(3, iLoop) = giCHECK_PASSED) Then
      iBulkBookingCount = iBulkBookingCount + 1
      
      If fBooked Or gfCourseIncludeProvisionals Then
        iNumberBooked = iNumberBooked + 1
      End If
    End If
  Next iLoop
  
  ' Progress bar
  With gobjProgress
    .AVI = dbLoadUsers
    .MainCaption = "Bulk Booking"
    .NumberOfBars = 1
    .Caption = "Training Booking"
    .Time = False
    .Cancel = True
    .Bar1Caption = "Bulk Booking..."
    .OpenProgress
    .Bar1MaxValue = iBulkBookingCount
  End With
  
  If iBulkBookingCount > 0 Then
    ' Check that we are not over-booking a course.
    If fOK Then
      fOK = CheckOverbooking(iNumberBooked)
      fOverBooked = Not fOK
    End If

    If fOK Then
      For iLoop = 1 To UBound(mavSelectedRecords, 2)
        ' Only book delegates who have passed the pre-requisite and
        ' availability criteria.
        If (mavSelectedRecords(2, iLoop) = giCHECK_PASSED) And _
          (mavSelectedRecords(3, iLoop) = giCHECK_PASSED) Then
    
          fOK = True
        
          If fOK Then
            'NHRD18012007 Fault 9017
            sDelegateDescription = CStr(mavSelectedRecords(4, iLoop))
'            sDelegateDescription = lvRecords.ListItems(iLoop).Text
'            For iLoop2 = 1 To lvRecords.ListItems(iLoop).ListSubItems.Count
'              sDelegateDescription = sDelegateDescription & ", " & _
'                lvRecords.ListItems(iLoop).SubItems(iLoop2)
'            Next iLoop2
            fOK = CheckOverlappedBooking(CLng(mavSelectedRecords(1, iLoop)), sDelegateDescription)
            gobjProgress.Visible = True
          End If
    
          If fOK Then
            ' Create the booking records.
            gADOCon.BeginTrans
            fInTransaction = True
            
            sSQL = "INSERT INTO " & objTrainingBookingTable.RealSource & _
              " (" & gsTrainBookStatusColumnName & ", " & _
              "id_" & Trim(Str(glngEmployeeTableID)) & ", " & _
              "id_" & Trim(Str(glngCourseTableID)) & ")" & _
              " VALUES" & _
              "(" & IIf(fBooked, "'B'", "'P'") & ", " & _
              Trim(Str(mavSelectedRecords(1, iLoop))) & ", " & _
              Trim(Str(mlngCourseID)) & ")"
              
            sErrorMsg = ""
            fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
            gobjProgress.UpdateProgress

            If Not fOK Then
              Screen.MousePointer = vbDefault
              If gobjProgress.Visible Then gobjProgress.CloseProgress
              COAMsgBox "Unable to create booking record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbInformation, App.ProductName
              Screen.MousePointer = vbHourglass
              
              gADOCon.RollbackTrans
              fInTransaction = False
            End If
        
            If fOK Then
              ' Delete any matching records in the Waiting List table
              ' if the user has permission to.
              If objWaitingListTable.AllowDelete Then
                Set objWLColumns = GetColumnPrivileges(gsWaitListTableName)
  
                If objWLColumns.Item(gsWaitListCourseTitleColumnName).AllowSelect Then
                  sSQL = "DELETE FROM " & objWaitingListTable.RealSource & _
                    " WHERE id_" & Trim(Str(glngEmployeeTableID)) & " = " & Trim(Str(mavSelectedRecords(1, iLoop))) & _
                    " AND " & gsWaitListCourseTitleColumnName & " = '" & Replace(msCourseTitle, "'", "''") & "'"

                  sErrorMsg = ""
                  fOK = datGeneral.ExecuteSql(sSQL, sErrorMsg)
                  
                  If Not fOK Then
                    Screen.MousePointer = vbDefault
                    gobjProgress.Visible = False
                    COAMsgBox "Unable to delete waiting list record." & vbCrLf & vbCrLf & sErrorMsg, vbOKOnly + vbInformation, App.ProductName
                    Screen.MousePointer = vbHourglass
                      
                    gADOCon.RollbackTrans
                    fInTransaction = False
                  End If
                End If
              
                Set objWLColumns = Nothing
              End If
            End If
            
            If fInTransaction Then
              If fOK Then
                iBookedCount = iBookedCount + 1
                gADOCon.CommitTrans
              Else
                gADOCon.RollbackTrans
              End If
              fInTransaction = False
            End If
          End If
        End If
      Next iLoop
    
      If iBookedCount > 0 Then
        ' JPD20011101 Fault 3466
        ' JPD20011101 Fault 3082
        ' JPD20021115 Fault 4754
        If gobjProgress.Visible Then gobjProgress.CloseProgress
        COAMsgBox Trim(Str(iBookedCount)) & " booking" & IIf(iBookedCount = 1, "", "s") & " made successfully.", vbOKOnly & vbInformation, App.ProductName
      End If
    End If
  End If
  
TidyUpAndExit:
  
  If gobjProgress.Visible Then gobjProgress.CloseProgress
 
  If fInTransaction Then
    If fOK Then
      gADOCon.CommitTrans
      ' JPD20011101 Fault 3082
      COAMsgBox "Booking(s) made successfully.", vbOKOnly & vbInformation, App.ProductName
    Else
      gADOCon.RollbackTrans
    End If
    fInTransaction = False
  End If

  Set objTrainingBookingTable = Nothing
  Set objWaitingListTable = Nothing
  
  Screen.MousePointer = vbDefault
  
  CreateBookings = fOverBooked
  Exit Function
  
ErrorTrap:
  fOK = False
  COAMsgBox Err.Description, vbExclamation + vbOKOnly, Application.Name
  Resume TidyUpAndExit

End Function




Private Function CheckOverlappedBooking(plngEmployeeID As Long, psDescription As String) As Boolean
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
      pmADO.Value = mlngCourseID

      Set pmADO = .CreateParameter("employeeRecordID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = plngEmployeeID

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
          gobjProgress.Visible = False
          COAMsgBox "'" & psDescription & "'  is already booked on a course that overlaps with this course." & vbCrLf & _
            "Unable to make the booking.", vbOKOnly + vbInformation, App.ProductName
            
        Case 2    ' Overlapped booking (over-rideable by the user).
          gobjProgress.Visible = False
          fOK = (COAMsgBox("'" & psDescription & "' is already booked on a course that overlaps with this course." & vbCrLf & _
            "Do you still want to make the booking ?", vbYesNo + vbQuestion, App.ProductName) = vbYes)
                  
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




Private Function CheckOverbooking(piNewBookings As Integer) As Boolean
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
      pmADO.Value = mlngCourseID

      Set pmADO = .CreateParameter("bookingID", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = 0

      Set pmADO = .CreateParameter("newBookings", adInteger, adParamInput)
      .Parameters.Append pmADO
      pmADO.Value = piNewBookings

      Set pmADO = .CreateParameter("result", adInteger, adParamOutput)
      .Parameters.Append pmADO
  
      Set pmADO = Nothing

      cmADO.Execute

      Select Case .Parameters("result").Value
        Case 1    ' Course fully booked (error).
          fOK = False
          COAMsgBox "The number of delegates selected would exceed the maximum number allowed on the course." & vbCrLf & _
            "Unable to make the booking" & IIf(piNewBookings = 1, ".", "s."), vbOKOnly + vbInformation, App.ProductName
            
        Case 2    ' Course fully booked (over-rideable by the user).
          fOK = (COAMsgBox("The number of delegates selected would exceed the maximum number allowed on the course." & vbCrLf & _
            "Do you still want to make the booking" & IIf(piNewBookings = 1, "?", "s ?"), vbYesNo + vbQuestion, App.ProductName) = vbYes)
                  
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

Private Function ItemInList(plngID As Long) As Boolean
  ' Return TRUE if the given ID is already in the picklist.
  Dim fInList As Boolean
  Dim objItem As MSComctlLib.ListItem
  
  fInList = False

  For Each objItem In lvRecords.ListItems
    If Trim(objItem.Tag) = Trim(Str(plngID)) Then
      fInList = True
      Exit For
    End If
  Next objItem
  Set objItem = Nothing

  ItemInList = fInList
  
End Function

Private Sub lvRecords_ClearSelections()
  ' Deselect any currently selected items in the listview.
  Dim objNode As MSComctlLib.ListItem
  
  For Each objNode In lvRecords.ListItems
    objNode.Selected = False
  Next objNode
  Set objNode = Nothing

End Sub

Private Function lvRecords_SelectedItemsCount() As Long
  ' Return the count of selected items in the listview.
  Dim lngCount As Long
  Dim objNode As MSComctlLib.ListItem
  
  lngCount = 0
  For Each objNode In lvRecords.ListItems
    If objNode.Selected Then lngCount = lngCount + 1
  Next objNode
  Set objNode = Nothing
  
  lvRecords_SelectedItemsCount = lngCount
  
End Function


Private Sub RefreshControls()
  ' Enable/Disable controls as required.
  'cmdDelete.Enabled = (lvRecords_SelectedItemsCount > 0)
  cmdDelete.Enabled = (lvRecords_SelectedItemsCount > 0) Or (lvRecords.ListItems.Count > 0)
  cmdDeleteAll.Enabled = (lvRecords.ListItems.Count > 0)
  cmdOK.Enabled = (lvRecords.ListItems.Count > 0)
  
End Sub


Public Property Get Cancelled() As Boolean
  Cancelled = mfCancelled

End Property
Public Property Let Cancelled(ByVal pfCancelled As Boolean)
  mfCancelled = pfCancelled

End Property


Private Sub cmdAddAll_Click()
  ' Add all records for the selected table into the bulk booking selection.
  On Error GoTo ErrorTrap

  Dim lngCount As Long
  Dim sSQL As String
  Dim sRecord As String
  Dim rsItem As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim strFormat As String

  Screen.MousePointer = vbHourglass

  lvRecords_ClearSelections

  Set rsItem = datGeneral.GetRecords(mSelectDelegateSQL)
  With rsItem
    Do While Not .EOF

      sRecord = ""
      If IsNull(rsItem(0)) Then
        sRecord = ""
      Else
        If rsItem.Fields(0).Type = adDBTimeStamp Then
          sRecord = Format(rsItem(0), DateFormat)
        ElseIf rsItem.Fields(0).Type = adNumeric Then
          ' Are thousand separators used
          strFormat = "0"
          If mavOrderDefinition(6, 1) Then
            strFormat = "#,0"
          End If
          If mavOrderDefinition(5, 1) > 0 Then
            strFormat = strFormat & "." & String(mavOrderDefinition(5, 1), "0")
          End If
          
          sRecord = Format(rsItem(0), strFormat)
        Else
          sRecord = rsItem(0)
        End If
      End If

      ' Check if the current item is already in the bulk booking selection.
      If Not ItemInList(rsItem(rsItem.Fields.Count - 1)) Then
        Set objNode = lvRecords.ListItems.Add(, , sRecord)
        objNode.Selected = True

        For lngCount = 1 To (rsItem.Fields.Count - 1)
          If IsNull(rsItem(lngCount)) Then
            sRecord = ""
          Else
            If rsItem.Fields(lngCount).Type = adDBTimeStamp Then
              sRecord = Format(rsItem(lngCount), DateFormat)
            ElseIf rsItem.Fields(lngCount).Type = adNumeric Then
              ' Are thousand separators used
              strFormat = "0"
              If mavOrderDefinition(6, lngCount + 1) Then
                strFormat = "#,0"
              End If
              If mavOrderDefinition(5, lngCount + 1) > 0 Then
                strFormat = strFormat & "." & String(mavOrderDefinition(5, lngCount + 1), "0")
              End If
              
              sRecord = Format(rsItem(lngCount), strFormat)
            Else
              sRecord = rsItem(lngCount)
            End If
          End If

          If lngCount < (rsItem.Fields.Count - 1) Then
            objNode.SubItems(lngCount) = sRecord
          Else
            objNode.Tag = sRecord
          End If
        Next lngCount
      End If

      .MoveNext
    Loop

    .Close
  End With
  Set rsItem = Nothing

  RefreshControls

  Screen.MousePointer = vbDefault

  Exit Sub

ErrorTrap:

End Sub

Private Sub cmdAddFilter_Click()
  ' Add a filtered set of records into the Bulk booking selection.
  On Error GoTo ErrorTrap

  Dim fApply As Boolean
  Dim lngCount As Long
  Dim sSQL As String
  Dim sRecord As String
  Dim sFilteredIDs As String
  Dim rsItem As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim objFilter As clsExprExpression
  Dim strFormat As String

  ReDim mastrUDFsRequired(0)

  Screen.MousePointer = vbHourglass
  fApply = False
  
  lvRecords_ClearSelections

  sSQL = mSelectDelegateSQL

  ' Add the filter 'where' clause code.
  Set objFilter = New clsExprExpression
  With objFilter
    ' Initialise the filter expression object.
    If .Initialise(glngEmployeeTableID, 0, giEXPR_RUNTIMEFILTER, giEXPRVALUE_LOGIC) Then
      .SelectExpression True
  
      fApply = (.ExpressionID > 0)
  
      If fApply Then
        fApply = datGeneral.FilteredIDs(.ExpressionID, sFilteredIDs)
      End If
      
      ' Generate any UDFs that are used in this filter
      If fApply Then
        datGeneral.FilterUDFs .ExpressionID, mastrUDFsRequired()
      End If
      
      If fApply Then
        If Len(sFilteredIDs) > 0 Then
          sSQL = sSQL & _
            IIf(InStr(UCase(sSQL), " WHERE ") > 0, " AND ", " WHERE ") & msDelegateRealSource & ".id IN (" & sFilteredIDs & ")"
        End If
      End If
    End If
  End With
  Set objFilter = Nothing

  If fApply Then
    
    If fApply Then fApply = UDFFunctions(mastrUDFsRequired, True)

    Set rsItem = datGeneral.GetRecords(sSQL)

    If fApply Then fApply = UDFFunctions(mastrUDFsRequired, False)

    With rsItem
      Do While Not .EOF

        sRecord = ""
        If IsNull(rsItem(0)) Then
          sRecord = ""
        Else
          If rsItem.Fields(0).Type = adDBTimeStamp Then
            sRecord = Format(rsItem(0), DateFormat)
          ElseIf rsItem.Fields(0).Type = adNumeric Then
            ' Are thousand separators used
            strFormat = "0"
            If mavOrderDefinition(6, 1) Then
              strFormat = "#,0"
            End If
            If mavOrderDefinition(5, 1) > 0 Then
              strFormat = strFormat & "." & String(mavOrderDefinition(5, 1), "0")
            End If
            
            sRecord = Format(rsItem(0), strFormat)
          Else
            sRecord = rsItem(0)
          End If
        End If

        ' Check if the current item is already in the Bulk Booking selection.
        fApply = Not ItemInList(rsItem(rsItem.Fields.Count - 1))
        If fApply Then
          Set objNode = lvRecords.ListItems.Add(, , sRecord)
          objNode.Selected = True

          For lngCount = 1 To (rsItem.Fields.Count - 1)
            If IsNull(rsItem(lngCount)) Then
              sRecord = ""
            Else
              If rsItem.Fields(lngCount).Type = adDBTimeStamp Then
                sRecord = Format(rsItem(lngCount), DateFormat)
              ElseIf rsItem.Fields(lngCount).Type = adNumeric Then
                ' Are thousand separators used
                strFormat = "0"
                If mavOrderDefinition(6, lngCount + 1) Then
                  strFormat = "#,0"
                End If
                If mavOrderDefinition(5, lngCount + 1) > 0 Then
                  strFormat = strFormat & "." & String(mavOrderDefinition(5, lngCount + 1), "0")
                End If
                
                sRecord = Format(rsItem(lngCount), strFormat)
              Else
                sRecord = rsItem(lngCount)
              End If
            End If

            If lngCount < (rsItem.Fields.Count - 1) Then
              objNode.SubItems(lngCount) = sRecord
            Else
              objNode.Tag = sRecord
            End If
          Next lngCount
        Else
          'NHRD19012007 Fault 4559
          COAMsgBox "This delegate is already in the Bulk Booking list", vbExclamation
          lvRecords.SelectedItem.Selected = 1
        End If

        .MoveNext
      Loop

      .Close
    End With
    Set rsItem = Nothing

    RefreshControls
  End If

  Screen.MousePointer = vbDefault

  Exit Sub

ErrorTrap:

End Sub


Private Sub cmdCancel_Click()
  ' Exit the form, without saving changes.
  Cancelled = True
  Me.Hide
  Screen.MousePointer = vbDefault

End Sub


Private Sub cmdDelete_Click()
  ' Remove the selected item from the picklist definition.
  Dim lngIndex As Long
  Dim lngNextIndex As Long
  Dim objNode As MSComctlLib.ListItem
  Dim alngIndices() As Long

  Screen.MousePointer = vbHourglass

  ' Construct an array of the item indices to be deleted.
  ReDim alngIndices(0)
  For Each objNode In lvRecords.ListItems
    If objNode.Selected Then
      lngNextIndex = UBound(alngIndices) + 1
      ReDim Preserve alngIndices(lngNextIndex)
      alngIndices(lngNextIndex) = objNode.Index
    End If
  Next objNode
  Set objNode = Nothing

  For lngIndex = UBound(alngIndices) To 1 Step -1
    lvRecords.ListItems.Remove alngIndices(lngIndex)
  Next lngIndex

  If lvRecords.ListItems.Count > 0 Then
    lvRecords.SelectedItem = lvRecords.ListItems(1)
    lvRecords.SelectedItem.EnsureVisible
  End If

  Screen.MousePointer = vbDefault

  RefreshControls

End Sub


Private Sub cmdDeleteAll_Click()
  ' Remove all items from the Bulk Booking selection.
  If COAMsgBox("Remove all records from the Bulk Booking selection, are you sure ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
    Screen.MousePointer = vbHourglass
    lvRecords.ListItems.Clear
    Screen.MousePointer = vbDefault

    RefreshControls
  End If

End Sub


Private Sub cmdNew_Click()
  ' Display the find form for selecting items to add to the Bulk Booking selection.
  Dim fApply As Boolean
  Dim lngCount As Long
  Dim lngIndex As Long
  Dim sSQL As String
  Dim sRecord As String
  Dim alngRecordIDs() As Long
  Dim frmBBFind As frmBulkBookingFind
  Dim rsItem As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim strFormat As String

  ' Display the form for the user to selected the required records.
  Set frmBBFind = New frmBulkBookingFind
  With frmBBFind
    If .Initialise Then
      .Show vbModal

      If Not .Cancelled Then
        ReDim alngRecordIDs(0)
        alngRecordIDs = .SelectedRecordIDs
  
        ' Add the selected records to the Bulk Booking selection listbox.
        If UBound(alngRecordIDs) > 0 Then
          Screen.MousePointer = vbHourglass
  
          lvRecords_ClearSelections
  
          sSQL = mSelectDelegateSQL & _
            IIf(InStr(UCase(mSelectDelegateSQL), " WHERE ") > 0, " AND ", " WHERE ") & msDelegateRealSource & ".id IN (0"
          For lngIndex = 1 To UBound(alngRecordIDs)
            sSQL = sSQL & ", " & alngRecordIDs(lngIndex)
          Next lngIndex
          sSQL = sSQL & ")"
          
          Set rsItem = datGeneral.GetRecords(sSQL)
          
          With rsItem
            Do While Not .EOF
          
              sRecord = ""
              If IsNull(rsItem(0)) Then
                sRecord = ""
              Else
                If rsItem.Fields(0).Type = adDBTimeStamp Then
                  sRecord = Format(rsItem(0), DateFormat)
                ElseIf rsItem.Fields(0).Type = adNumeric Then
                  ' Are thousand separators used
                  strFormat = "0"
                  If mavOrderDefinition(6, 1) Then
                    strFormat = "#,0"
                  End If
                  If mavOrderDefinition(5, 1) > 0 Then
                    strFormat = strFormat & "." & String(mavOrderDefinition(5, 1), "0")
                  End If
                  
                  sRecord = Format(rsItem(0), strFormat)
                Else
                  sRecord = rsItem(0)
                End If
              End If
          
              ' Check if the current item is already in the Bulk Booking selection.
              fApply = Not ItemInList(rsItem(rsItem.Fields.Count - 1))
              If fApply Then
                Set objNode = lvRecords.ListItems.Add(, , sRecord)
                objNode.Selected = True
          
                For lngCount = 1 To (rsItem.Fields.Count - 1)
                  If IsNull(rsItem(lngCount)) Then
                    sRecord = ""
                  Else
                    If rsItem.Fields(lngCount).Type = adDBTimeStamp Then
                      sRecord = Format(rsItem(lngCount), DateFormat)
                    ElseIf rsItem.Fields(lngCount).Type = adNumeric Then
                      ' Are thousand separators used
                      strFormat = "0"
                      If mavOrderDefinition(6, lngCount + 1) Then
                        strFormat = "#,0"
                      End If
                      If mavOrderDefinition(5, lngCount + 1) > 0 Then
                        strFormat = strFormat & "." & String(mavOrderDefinition(5, lngCount + 1), "0")
                      End If
                      
                      sRecord = Format(rsItem(lngCount), strFormat)
                    Else
                      sRecord = rsItem(lngCount)
                    End If
                  End If
          
                  If lngCount < (rsItem.Fields.Count - 1) Then
                    objNode.SubItems(lngCount) = sRecord
                  Else
                    objNode.Tag = sRecord
                  End If
                Next lngCount
              
                Set objNode = Nothing
              Else
                'NHRD15012007 Fault 10610
                COAMsgBox "This delegate is already in the Bulk Booking list", vbExclamation
                lvRecords.SelectedItem.Selected = 1
              End If
          
              .MoveNext
            Loop
            
            .Close
          End With
          Set rsItem = Nothing

          Screen.MousePointer = vbDefault
          RefreshControls
        End If
      End If
    End If
  End With

  Unload frmBBFind
  Set frmBBFind = Nothing

End Sub


Private Sub cmdOK_Click()
  ' Bulk Book the selected records on the current course.
  Dim fOK As Boolean
  Dim iNextIndex As Integer
  Dim objNode As MSComctlLib.ListItem
  Dim strSortOrderColums As String
  Dim objTableView As CTablePrivilege
  Dim objDelegateTable As CTablePrivilege
  ' Get the Delegate table object to determine the order
  Set objDelegateTable = gcoTablePrivileges.FindTableID(glngEmployeeTableID)
  ' Create an array of the selected record IDs and whether or not they
  ' pass the booking criteria.
  ' Index 1 = record ID.
  ' Index 2 = pre-requisite check status.
  ' Index 3 = availibility check status.
  ' Check status = giCHECK_PASSED (1) if the check passed.
  '              = giCHECK_FAILED (2) if the check failed.
  '              = giCHECK_FAILEDQUERY (3) if the check failed but can be over-ridden.
  ReDim mavSelectedRecords(4, 0)
  For Each objNode In lvRecords.ListItems
    strSortOrderColums = EvaluateRecordDescription(objNode.Tag, objDelegateTable.RecordDescriptionID)
    iNextIndex = UBound(mavSelectedRecords, 2) + 1
    ReDim Preserve mavSelectedRecords(4, iNextIndex)
    mavSelectedRecords(1, iNextIndex) = Val(objNode.Tag)
    mavSelectedRecords(2, iNextIndex) = giCHECK_PASSED
    mavSelectedRecords(3, iNextIndex) = giCHECK_PASSED
    mavSelectedRecords(4, iNextIndex) = strSortOrderColums
  Next objNode
  
  Set objNode = Nothing
  Set objDelegateTable = Nothing
  ' Check that selected employees have (or will have) satisfied the pre-requisite criteria.
  fOK = CheckPreRequisites
  ' Check that the selected employees are not unavailable for the selected course.
  If fOK Then
    fOK = CheckAvailability
  End If
  ' Create booking records.
  If fOK Then
    If Not CreateBookings Then
      Cancelled = False
      Me.Hide
    End If
  End If
End Sub


Private Sub cmdPicklistAdd_Click()
  ' Add a picklist set of records into the Bulk booking selection.
  On Error GoTo ErrorTrap

  Dim fExit As Boolean
  Dim fApply As Boolean
  Dim lngCount As Long
  Dim lngPicklistID As Long
  Dim sSQL As String
  Dim sPicklistSQL As String
  Dim sRecord As String
  Dim sPickListIDs As String
  Dim rsItem As Recordset
  Dim rsWhere As Recordset
  Dim objNode As MSComctlLib.ListItem
  Dim frmPick As frmPicklists
  Dim frmSelection As frmDefSel
  Dim strFormat As String
  
  Screen.MousePointer = vbHourglass

  lvRecords_ClearSelections

  sPicklistSQL = mSelectDelegateSQL

  ' Display the picklist selection form.
  lngPicklistID = 0
  Set frmSelection = New frmDefSel
  
  With frmSelection
    Do While Not fExit
      
      .TableID = glngEmployeeTableID
      .TableComboVisible = True
      .TableComboEnabled = False
      If lngPicklistID > 0 Then
        .SelectedID = lngPicklistID
      End If
      
      If .ShowList(utlPicklist) Then
        .Show vbModal
          
        Select Case .Action
          Case edtAdd
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(True, False, glngEmployeeTableID) Then
              frmPick.Show vbModal
            End If
            Set frmPick = Nothing
          
          Case edtEdit
            Set frmPick = New frmPicklists
            If frmPick.InitialisePickList(False, .FromCopy, glngEmployeeTableID, .SelectedID) Then
              frmPick.Show vbModal
            End If
            Set frmPick = Nothing
          
          'MH20050728 Fault 10232
          Case edtPrint
            Set frmPick = New frmPicklists
            frmPick.PrintDef .TableID, .SelectedID
            Unload frmPick
            Set frmPick = Nothing
          
          Case edtSelect
            lngPicklistID = .SelectedID
            fExit = True
              
          Case 0
            fExit = True
                
        End Select
      End If
    Loop
  End With
  
  Set frmSelection = Nothing

  If lngPicklistID > 0 Then
    sPickListIDs = ""
    sSQL = "EXEC sp_ASRGetPickListRecords " & Trim(Str(lngPicklistID))
    Set rsWhere = datGeneral.GetRecords(sSQL)
    
    Do While Not rsWhere.EOF
      sPickListIDs = sPickListIDs & IIf(Len(sPickListIDs) > 0, ", ", "") & Trim(Str(rsWhere!ID))
      rsWhere.MoveNext
    Loop
    rsWhere.Close
    Set rsWhere = Nothing
    
    If Len(sPickListIDs) > 0 Then
      sPicklistSQL = sPicklistSQL & _
        IIf(InStr(UCase(sPicklistSQL), " WHERE ") > 0, " AND ", " WHERE ") & msDelegateRealSource & ".id IN (" & sPickListIDs & ")"
      
      Set rsItem = datGeneral.GetRecords(sPicklistSQL)
  
      With rsItem
        Do While Not .EOF
          sRecord = ""
          If IsNull(rsItem(0)) Then
            sRecord = ""
          Else
            If rsItem.Fields(0).Type = adDBTimeStamp Then
              sRecord = Format(rsItem(0), DateFormat)
            ElseIf rsItem.Fields(0).Type = adNumeric Then
              ' Are thousand separators used
              strFormat = "0"
              If mavOrderDefinition(6, 1) Then
                strFormat = "#,0"
              End If
              If mavOrderDefinition(5, 1) > 0 Then
                strFormat = strFormat & "." & String(mavOrderDefinition(5, 1), "0")
              End If
              
              sRecord = Format(rsItem(0), strFormat)
            Else
              sRecord = rsItem(0)
            End If
          End If
  
          ' Check if the current item is already in the Bulk Booking selection.
          fApply = Not ItemInList(rsItem(rsItem.Fields.Count - 1))
          If fApply Then
            Set objNode = lvRecords.ListItems.Add(, , sRecord)
            objNode.Selected = True
  
            For lngCount = 1 To (rsItem.Fields.Count - 1)
              If IsNull(rsItem(lngCount)) Then
                sRecord = ""
              Else
                If rsItem.Fields(lngCount).Type = adDBTimeStamp Then
                  sRecord = Format(rsItem(lngCount), DateFormat)
                ElseIf rsItem.Fields(lngCount).Type = adNumeric Then
                  ' Are thousand separators used
                  strFormat = "0"
                  If mavOrderDefinition(6, lngCount + 1) Then
                    strFormat = "#,0"
                  End If
                  If mavOrderDefinition(5, lngCount + 1) > 0 Then
                    strFormat = strFormat & "." & String(mavOrderDefinition(5, lngCount + 1), "0")
                  End If
                  
                  sRecord = Format(rsItem(lngCount), strFormat)
                Else
                  sRecord = rsItem(lngCount)
                End If
              End If
  
              If lngCount < (rsItem.Fields.Count - 1) Then
                objNode.SubItems(lngCount) = sRecord
              Else
                objNode.Tag = sRecord
              End If
            Next lngCount
          Else
            'NHRD19012007 Fault 4559
            COAMsgBox "This delegate is already in the Bulk Booking list", vbExclamation
            lvRecords.SelectedItem.Selected = 1          '
          End If
  
          .MoveNext
        Loop
  
        .Close
      End With
      Set rsItem = Nothing
  
      RefreshControls
    End If
  End If
  
  Screen.MousePointer = vbDefault

  Exit Sub

ErrorTrap:

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
  Cancelled = False

  Hook Me.hWnd, dblFORM_MINWIDTH, dblFORM_MINHEIGHT
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    cmdCancel_Click
    Exit Sub
  End If
    
  If Cancelled = True Then
    Exit Sub
  End If
    
End Sub


Private Sub Form_Resize()
  ' Resize the form's controls as the form is itself resized.
  Dim lCount As Long
  Dim lWidth As Long
  Dim iLastColumnIndex As Integer
  Dim iMaxPosition As Integer
  
  Const dblCOORD_XGAP = 200
  Const dblCOORD_SMALLXGAP = 150
  Const dblCOORD_YGAP = 200
  Const dblCOORD_SMALLYGAP = 100
  
  'JPD 20030908 Fault 5756
  DisplayApplication
  
  If Me.WindowState = vbNormal Then
  
'    ' Ensure the form does not get narrower than the defined minimum for a Find window.
'    If Me.Width < dblFORM_MINWIDTH Then
'      Me.Width = dblFORM_MINWIDTH
'    End If
'
'    ' Ensure the form does not get wider than the screen.
'    If Me.Width > Screen.Width Then
'      Me.Width = Screen.Width
'    End If
'
'    ' Initialise the form height.
'    If Not mfSizing Then
'      mfSizing = True
'      Me.Height = Screen.Height / 3
'    End If
'
'    ' Ensure the form does not get shorter than the defined minimum for a Find window.
'    If Me.Height < dblFORM_MINHEIGHT Then
'      mfSizing = True
'      Me.Height = dblFORM_MINHEIGHT
'    End If
'
'    ' Ensure the form does not get taller than the screen.
'    If Me.Height > Screen.Height Then
'      Me.Height = Screen.Height
'    End If
            
    ' Size the Booking Status controls.
    cboBookingStatus.Width = Me.ScaleWidth - cboBookingStatus.Left - fraAddRemoveButtons.Width - (2 * dblCOORD_XGAP)
    
    ' Size the listview.
    With lvRecords
      .Width = Me.ScaleWidth - .Left - fraAddRemoveButtons.Width - dblCOORD_XGAP - dblCOORD_SMALLXGAP
      .Height = Me.ScaleHeight - .Top - dblCOORD_YGAP
    End With
        
    ' Size the frame with the Add/Remove command buttons in.
    With fraAddRemoveButtons
      .Left = lvRecords.Left + lvRecords.Width + dblCOORD_XGAP
    End With
    
    ' Size the frame with the OK/Cancel command buttons in.
    With fraOKCancelButtons
      .Top = Me.ScaleHeight - .Height - dblCOORD_YGAP
      .Left = fraAddRemoveButtons.Left
    End With
    
    ' Stretch the last find column to fit the listview.
    iLastColumnIndex = -1
    iMaxPosition = -1
    With lvRecords
      If .ColumnHeaders.Count > 0 Then
        If .ColumnHeaders(.ColumnHeaders.Count).Left + .ColumnHeaders(.ColumnHeaders.Count).Width < .Width Then
          If .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - _
            ((UI.GetSystemMetrics(SM_CXFRAME) * 6) * Screen.TwipsPerPixelX) > 0 Then
            .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - _
              ((UI.GetSystemMetrics(SM_CXFRAME) * 6) * Screen.TwipsPerPixelX)
          End If
        End If
      End If
    End With
  End If

End Sub


Private Sub cboBookingStatus_Initialise()
  ' Initialise the Booking Status combo.
  With cboBookingStatus
    .Clear
    
    .AddItem "Booked"
    .ItemData(.NewIndex) = giSTATUS_BOOKED
  
    If gfTrainBookStatus_P Then
      .AddItem "Provisional"
      .ItemData(.NewIndex) = giSTATUS_PROVISIONAL
    Else
      .Enabled = False
    End If
    
    .ListIndex = 0
  End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unhook Me.hWnd
End Sub

