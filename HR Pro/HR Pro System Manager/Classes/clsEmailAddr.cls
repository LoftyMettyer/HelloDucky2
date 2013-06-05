VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Email definition variables.
Private mlngEmailID As Long
Private mstrEmailName As String
Private mlngTableID As Long
Private mintEmailType As Integer
Private mstrFixed As String
Private mlngColumnID As Long
Private mlngExpressionID As Long

' Object handing variables.
Private mfNew As Boolean
Private mfChanged As Boolean
Private mfDeleted As Boolean

Private mfrmUse As frmUsage
Private mcolEmailIDs As Collection

Public Property Get EmailType() As Integer
  EmailType = mintEmailType
End Property

Public Property Let EmailType(ByVal intNewValue As Integer)
  mintEmailType = intNewValue
End Property

Public Property Get Fixed() As String
  Fixed = mstrFixed
End Property

Public Property Let Fixed(ByVal strNewValue As String)
  mstrFixed = strNewValue
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(ByVal lngNewValue As Long)
  mlngColumnID = lngNewValue
End Property

Public Property Get ExpressionID() As Long
  ExpressionID = mlngExpressionID
End Property

Public Property Let ExpressionID(ByVal lngNewValue As Long)
  mlngExpressionID = lngNewValue
End Property

Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property

Public Property Get EmailID() As Long
  EmailID = mlngEmailID
End Property

Public Property Let EmailID(ByVal lngNewValue As Long)
  mlngEmailID = lngNewValue
End Property


Public Property Get EmailIDs() As Collection
  Set EmailIDs = mcolEmailIDs
End Property

Public Property Let EmailIDs(ByVal colNewValue As Collection)
  Set mcolEmailIDs = colNewValue
End Property


Public Property Get EmailName() As String
  EmailName = mstrEmailName
End Property

Public Property Let EmailName(psName As String)
  mstrEmailName = psName
End Property


Public Property Get IsNew() As Boolean
  IsNew = mfNew
End Property

Public Property Let IsNew(pfSetting As Boolean)
  mfNew = pfSetting
End Property

Public Property Get IsChanged() As Boolean
  IsChanged = mfChanged
End Property

Public Property Let IsChanged(pfValue As Boolean)
  mfChanged = pfValue
End Property

Public Property Let IsDeleted(pfSetting As Boolean)
  mfDeleted = pfSetting
End Property

Public Property Get IsDeleted() As Boolean
  IsDeleted = mfDeleted
End Property


'Public Function CopyEmail() As Boolean
'  On Error GoTo ErrorTrap
'
'  Dim fOK As Boolean
'  Dim sName As String
'  Dim frmEdit As frmEmailAddr
'
'  ' Begin the transaction of data to the local database.
'  daoWS.BeginTrans
'
'  'fOK = ConstructEmail
'
'  'Find & read table record to copy
'  'If fOK Then
'    mfNew = True
'    mfChanged = False
'    mfDeleted = False
'
'    mlngEmailID = 0
'    sName = mstrEmailName
'    mstrEmailName = "Copy_of_" & Trim(sName)
'
'    Set frmEdit = New frmEmailAddr
'    Set frmEdit.Email = Me
'    frmEdit.Show vbModal
'    fOK = Not frmEdit.Cancelled
'    Set frmEdit = Nothing
'
'    If fOK Then
'      fOK = WriteEmail
'    End If
'
'    If Not fOK Then
'      mfNew = True
'      mfChanged = False
'      mfDeleted = False
'      mlngEmailID = 0
'      mstrEmailName = sName
'    End If
'
'  'End If
'
'ExitCopyEmail:
'  ' Commit the data transaction if everything was okay.
'  If fOK Then
'    daoWS.CommitTrans dbForceOSFlush
'    Application.Changed = True
'  Else
'    daoWS.Rollback
'  End If
'  CopyEmail = fOK
'  Exit Function
'
'ErrorTrap:
'  fOK = False
'  Resume ExitCopyEmail
'
'End Function

Public Function DeleteEmail() As Boolean
  ' Delet the current Email object.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  fOK = True

  If fOK Then
    Set mfrmUse = New frmUsage
    mfrmUse.ResetList
    If EmailIsUsed(mfrmUse) Then
      Screen.MousePointer = vbDefault
      mfrmUse.ShowMessage Me.EmailName & " Email", "The email cannot be deleted as the email is used by the following:", UsageCheckObject.Email
      fOK = False
    End If
    UnLoad mfrmUse
    Set mfrmUse = Nothing
  End If

  If fOK Then
    'Find & read Email record to delete
    fOK = ConstructEmail

    If fOK Then
      'Check if table is already deleted
      If Not IsDeleted Then
        'Set record deleted flag
        mfDeleted = True

        'Write Email record
        fOK = WriteEmail
      End If
    End If

  End If

TidyUpAndExit:
  If fOK Then
    mfDeleted = True
  End If
  DeleteEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function EditEmail() As Boolean
  ' Edit the Email.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim frmEdit As frmEmailAddr

  ' Begin the transaction of data to the local database.
  'daoWS.BeginTrans   'MH20071108 Fault 12595

  fOK = ConstructEmail

  If fOK Then
    Set frmEdit = New frmEmailAddr

    frmEdit.Initialise Me
    frmEdit.Show vbModal

    fOK = Not frmEdit.Cancelled

  End If

  daoWS.BeginTrans    'MH20071108 Fault 12595

  If fOK Then

    With frmEdit.Email
      EmailName = .EmailName
      EmailType = .EmailType
      Fixed = .Fixed
      ColumnID = .ColumnID
      ExpressionID = .ExpressionID
      IsChanged = True
    End With

    fOK = WriteEmail
  End If

  Set frmEdit = Nothing

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  EditEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function FindEmail() As Boolean
  On Error GoTo ErrorTrap

  With recEmailAddrEdit
    .Index = "idxID"
    .Seek "=", EmailID
    FindEmail = (Not .NoMatch)
  End With

  Exit Function

ErrorTrap:
  FindEmail = False

End Function

Public Function NewEmail() As Boolean
  ' Create a new Email.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim frmEdit As frmEmailAddr

  ' Begin the transaction of data to the local database.
  'daoWS.BeginTrans     'MH20071108 Fault 12595

  mfNew = True
  mfChanged = False
  mfDeleted = False
  mlngEmailID = 0
  mstrEmailName = vbNullString
  EmailType = 0
  Fixed = vbNullString

  Set frmEdit = New frmEmailAddr
  frmEdit.Initialise Me
  frmEdit.Show vbModal
  fOK = Not frmEdit.Cancelled

  daoWS.BeginTrans      'MH20071108 Fault 12595

  If fOK Then

    With frmEdit.Email
      EmailID = .EmailID
      EmailName = .EmailName
      EmailType = .EmailType
      Fixed = .Fixed
      ColumnID = .ColumnID
      ExpressionID = .ExpressionID
      IsNew = True
    End With

    fOK = WriteEmail

    If Not fOK Then
      mfNew = True
      mfChanged = False
      mfDeleted = False
      mlngEmailID = 0
      mstrEmailName = vbNullString
    End If
  End If
  Set frmEdit = Nothing

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  NewEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function

Public Function ConstructEmail() As Boolean
  ' Read the Email definition from the local tables.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sIndex As String
  Dim sSQL As String

  With recEmailAddrEdit
    'Store current index
    sIndex = .Index

    'Find the table record
    fOK = FindEmail

    If fOK Then
      'Set table object properties
      mstrEmailName = !Name

      'MH20030819
      'Don't reset the table id at this stage in case the
      'user is changing a fixed email to a column or calc
      'mlngTableID = !TableID

      mintEmailType = IIf(IsNull(!Type), 0, !Type)
      mstrFixed = IIf(IsNull(!Fixed), vbNullString, !Fixed)
      mlngColumnID = IIf(IsNull(!ColumnID), 0, !ColumnID)
      mlngExpressionID = IIf(IsNull(!ExprID), 0, !ExprID)
      mfNew = !New
      mfChanged = !Changed
      mfDeleted = !Deleted

    Else
      'Set table object properties
      mstrEmailName = vbNullString
      mlngTableID = 0
      mintEmailType = 0
      mstrFixed = vbNullString
      mlngColumnID = 0
      mlngExpressionID = 0
      mfNew = True
      mfChanged = False
      mfDeleted = False
    End If
  End With

ExitConstructEmail:
  'Restore original index
  recEmailAddrEdit.Index = sIndex
  ConstructEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitConstructEmail

End Function


Public Function SelectEmail(Optional pvarReadOnly As Variant, Optional pvarSelectMultiple As Variant) As Boolean
  ' Display the Email selection form.
  Dim fExit As Boolean
  Dim lngOldEmailID As Long
  Dim frmSelectEmail As frmSelection
  Dim lngTableID As Long
  Dim strSQL As String
  Dim fReadOnly As Boolean
  Dim fSelectMultiple As Boolean
  Dim lngTempID As Long

  If IsMissing(pvarReadOnly) Then
    fReadOnly = False
  Else
    fReadOnly = CBool(pvarReadOnly)
  End If

  If IsMissing(pvarSelectMultiple) Then
    fSelectMultiple = False
  Else
    fSelectMultiple = CBool(pvarSelectMultiple)
  End If



  'JPD 20060512 Want to be able to choose just email addresses that are not tied
  ' to any table
  If TableID >= 0 Then

    fExit = False
    lngOldEmailID = mlngEmailID
    lngTableID = Me.TableID
    Set frmSelectEmail = New frmSelection


    Do While Not fExit
      ' Set the selection form type.
      ' Set the selection form properties.
      frmSelectEmail.SelectedIDs = mcolEmailIDs
      frmSelectEmail.ForcedReadOnly = fReadOnly
      frmSelectEmail.SelectMultiple = fSelectMultiple
      Me.TableID = lngTableID
      Set frmSelectEmail.Email = Me
      frmSelectEmail.Show vbModal

      Select Case frmSelectEmail.Action
        Case edtAdd
          NewEmail
          'frmSelectEmail.SelectedIDs.Add frmSelectEmail.SelectedID, CStr(frmSelectEmail.SelectedID)

        Case edtEdit
          EmailID = frmSelectEmail.SelectedID
          EditEmail

        Case edtCopy
          EmailID = frmSelectEmail.SelectedID
          CopyEmail

        Case edtDelete
          'TM20020528 Fault 3924
          If MsgBox("Delete this email address, are you sure?", vbQuestion + vbYesNo, "Confirm Delete") = vbYes Then
            EmailID = frmSelectEmail.SelectedID
            DeleteEmail_Transaction
          End If

        Case edtPrint
          EmailID = frmSelectEmail.SelectedID
          PrintEmail

        Case edtSelect
          EmailID = frmSelectEmail.SelectedID
          SelectEmail = ConstructEmail

          EmailIDs = frmSelectEmail.SelectedIDs
          fExit = True

        ' Deselect all expressions and return to the calling form.
        Case edtDeselect
          EmailID = -1
          EmailIDs = Nothing
          EmailIDs = New Collection
          ConstructEmail
          SelectEmail = True
          fExit = True

        Case edtCancel
          ' Check if the original Email still exists.
          With recEmailAddrEdit
            .Index = "idxID"
            .Seek "=", lngOldEmailID

            If .NoMatch Then
              EmailID = 0
            Else
              If !Deleted Then
                EmailID = 0
              End If
            End If
          End With

          SelectEmail = False
          fExit = True

      End Select
    Loop

    Set frmSelectEmail = Nothing
  End If

End Function


Private Function DeleteEmail_Transaction() As Boolean
  ' Transaction wrapper for the 'DeleteEmail' function.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean

  ' Begin the transaction of data to the local database.
  daoWS.BeginTrans

  fOK = DeleteEmail

TidyUpAndExit:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  DeleteEmail_Transaction = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function CopyEmail() As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sName As String
  Dim frmEdit As frmEmailAddr

  ' Begin the transaction of data to the local database.
  'daoWS.BeginTrans     'MH20071108 Fault 12595

  fOK = ConstructEmail

  'Find & read table record to copy
  If fOK Then
    mfNew = True
    mfChanged = False
    mfDeleted = False

    mlngEmailID = 0
    sName = mstrEmailName
    mstrEmailName = "Copy_of_" & Trim(sName)

    Set frmEdit = New frmEmailAddr
    frmEdit.Initialise Me
    'TM20020320 Fault 3548 - let the email Form's changed be true... let it be... let it be...
    frmEdit.Changed = True
    frmEdit.Show vbModal
    fOK = Not frmEdit.Cancelled
    Set frmEdit = Nothing
  End If

  daoWS.BeginTrans      'MH20071108 Fault 12595

  If fOK Then
    fOK = WriteEmail
  End If

  If Not fOK Then
    mfNew = True
    mfChanged = False
    mfDeleted = False
    mlngEmailID = 0
    mstrEmailName = sName
  End If


ExitCopyEmail:
  ' Commit the data transaction if everything was okay.
  If fOK Then
    daoWS.CommitTrans dbForceOSFlush
    Application.Changed = True
  Else
    daoWS.Rollback
  End If
  CopyEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume ExitCopyEmail

End Function


Public Function WriteEmail() As Boolean
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim sIndex As String

  With recEmailAddrEdit
    'Store current index
    sIndex = .Index

    If (IsNew And Not IsChanged) Then
      'Get new Email id
'      EmailID = UniqueColumnValue("tmpEmailAddresses", "EmailID")
      EmailID = GetUniqueID("emailaddress", "tmpEmailAddresses", "emailID")

      'Add a record for new Email
      fOK = (EmailID > 0)

      If fOK Then
        .AddNew
        !EmailID = EmailID
      End If
    Else
      'Find existing table record
      fOK = FindEmail
      If fOK Then
        'Edit table record
        .Edit
      End If
    End If

    ' Write the Email definition to the local tables.
    If fOK Then
      mfChanged = True
      !TableID = IIf(EmailType > 0, TableID, 0)
      !Name = EmailName
      !Type = EmailType
      !Fixed = Fixed
      !ColumnID = ColumnID
      !ExprID = ExpressionID
      !New = IsNew
      !Changed = IsChanged
      !Deleted = IsDeleted
      .Update
    End If
  End With

TidyUpAndExit:
  recEmailAddrEdit.Index = sIndex
  WriteEmail = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function CloneEmail(pavCloneRegister As Variant) As clsEmailAddr
  ' Clone the current Email.
  ' Used when copying table definitions.
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim iIndex As Integer
  Dim lngNewColumnID As Long
  Dim objClonedEmail As clsEmailAddr
  Dim lngNewExprID As Long

  ' Read the Email definition.
  fOK = ConstructEmail

  If fOK Then
    ' Instantiate the clone Email.
    Set objClonedEmail = New clsEmailAddr
    ' Clone the current Email's properties.
    With objClonedEmail
      .IsNew = True
      .IsDeleted = False
      .IsChanged = False
      .EmailID = 0
      .EmailName = EmailName

      .EmailName = Me.EmailName
      .EmailType = Me.EmailType

      lngNewColumnID = 0
      lngNewExprID = 0
      If Me.ColumnID > 0 Then
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "COLUMN" And _
            pavCloneRegister(2, iIndex) = Me.ColumnID Then
            lngNewColumnID = pavCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
      End If

      If Me.ExpressionID > 0 Then
        For iIndex = 1 To UBound(pavCloneRegister, 2)
          If pavCloneRegister(1, iIndex) = "EXPRESSION" And _
            pavCloneRegister(2, iIndex) = Me.ExpressionID Then
            lngNewExprID = pavCloneRegister(3, iIndex)
            Exit For
          End If
        Next iIndex
      End If

      .ColumnID = lngNewColumnID
      .Fixed = Me.Fixed
      .ExpressionID = lngNewExprID

    End With
  End If

TidyUpAndExit:
  If fOK Then
    Set CloneEmail = objClonedEmail
  Else
    Set CloneEmail = Nothing
  End If
  ' Disassociate object variables.
  Set objClonedEmail = Nothing
  Exit Function

ErrorTrap:
  fOK = False
  Resume TidyUpAndExit

End Function


Public Function PrintEmail() As Boolean

  Dim objPrintDef As clsPrintDef
  Dim strSelectedPrinter As String
  Dim blnOK As Boolean

  Set objPrintDef = New clsPrintDef

  Screen.MousePointer = vbHourglass
  strSelectedPrinter = vbNullString

  blnOK = (objPrintDef.IsOK And ConstructEmail)

  If blnOK Then

    With objPrintDef
      If .PrintStart(True) Then

        strSelectedPrinter = Printer.DeviceName

        .PrintHeader "Email Address : " & Me.EmailName

        Select Case Me.EmailType
        Case 0
          .PrintNormal "Type: Fixed"
          .PrintNormal "Value: '" & Trim(Me.Fixed) & "'"
        Case 1
          .PrintNormal "Type: Column"
          .PrintNormal "Column: '" & GetColumnName(Me.ColumnID) & "'"
        Case 2
          .PrintNormal "Type: Calculated"
          .PrintNormal "Calculation: '" & GetExpressionName(Me.ExpressionID) & "'"
        End Select

        .PrintEnd
      
      End If
    End With

  End If

TidyAndExit:
  Screen.MousePointer = vbDefault
  If strSelectedPrinter <> vbNullString Then
    If blnOK Then
      'MH20040127 Fault 7991
      'MsgBox "Email Address : " & Me.EmailName & " printing complete." & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")", vbInformation, "Email Address"
      MsgBox "Email Address : " & Me.EmailName & " printing complete." & vbCrLf & vbCrLf & "(" & strSelectedPrinter & ")", vbInformation, "Email Address"
    Else
      MsgBox "Printing Email Address Definition Failed", vbExclamation, "Email Address"
    End If
  End If

Exit Function

LocalErr:
  blnOK = False
  Resume TidyAndExit

End Function


Public Function EmailIsUsed(pfrmUsage As frmUsage) As Boolean

  Dim fUsed As Boolean
  Dim sSQL As String
  Dim rsDAOCheck As DAO.Recordset
  Dim adoRes As New ADODB.Recordset
  Dim rsModules As DAO.Recordset
  Dim sMessage As String
  Dim bWF As Boolean
  Dim sModuleName As String

  On Error GoTo ErrorTrap

  Set mfrmUse = pfrmUsage

  sMessage = vbNullString

  fUsed = False

  'TM20011003 Fault 2191 - 5
  'Make check for the email definition being used in email link definitions.
    sSQL = "SELECT DISTINCT tmpEmailLinks.Title, tmpTables.Tablename " & vbCrLf & _
           "FROM tmpEmailLinks, tmpEmailLinksRecipients, tmpTables " & vbCrLf & _
           "WHERE (tmpEmailLinksRecipients.LinkID = tmpEmailLinks.LinkID) " & _
           "  AND (tmpEmailLinks.TableID = tmpTables.TableID) " & _
           "  AND (tmpEmailLinksRecipients.RecipientID = " & Me.EmailID & ")"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Email Link : " & !Title & " <" & !TableName & ">"
          .MoveNext
        Loop
        .Close
      End If
    End With
    Set rsDAOCheck = Nothing


  'TM20020211 Fault 3487
  'Make check for the email definition being used as default email on a table.
    sSQL = "SELECT DISTINCT tmpEmailAddresses.EMailID, tmpTables.tableID, tmpTables.TableName " & _
            "FROM tmpEmailAddresses, tmpTables " & _
            "WHERE (tmpTables.DefaultEmailID = tmpEmailAddresses.EMailID) " & _
            " AND (tmpEmailAddresses.EmailID = " & Me.EmailID & ")"
    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
    With rsDAOCheck
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList "Default Email Address : " & !TableName
          .MoveNext
        Loop
        .Close
      End If
    End With
    Set rsDAOCheck = Nothing


'    'Make check for the email definition being used as Audit email notifications.
'    sSQL = "SELECT DISTINCT tmpTables.tableID, tmpTables.TableName " & _
'            "FROM tmpTables " & _
'            "WHERE (tmpTables.EmailInsert = " & Me.EmailID & ")" & _
'            " OR (tmpTables.EmailDelete = " & Me.EmailID & ")"
'    Set rsDAOCheck = daoDb.OpenRecordset(sSQL)
'    With rsDAOCheck
'      If Not (.EOF And .BOF) Then
'        fUsed = True
'        Do Until .EOF
'          mfrmUse.AddToList "Audit Email Notification : " & !TableName & " table"
'          .MoveNext
'        Loop
'        .Close
'      End If
'    End With
'    Set rsDAOCheck = Nothing



  'MH20030922 Fault 6139
  'Make check for the email definition being used in an email group.
    sSQL = "SELECT DISTINCT 'Email Group' as UType, ASRSysEmailGroupName.Name " & _
          "FROM ASRSysEmailGroupItems " & _
          "JOIN ASRSysEmailGroupName on ASRSysEmailGroupItems.EmailGroupID = ASRSysEmailGroupName.EmailGroupID " & _
          "WHERE ASRSysEmailGroupItems.EmailDefID = " & CStr(Me.EmailID) & vbCrLf
    sSQL = sSQL & "UNION" & vbCrLf & _
          "SELECT DISTINCT CASE WHEN ASRSysMailMergeName.IsLabel = 1 THEN 'Envelope & Label'" & _
          " ELSE 'Mail Merge' END as UType, ASRSysMailMergeName.Name " & _
          "FROM ASRSysMailMergeName " & _
          "WHERE ASRSysMailMergeName.EmailAddrID = " & CStr(Me.EmailID)
    adoRes.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
    With adoRes
      If Not (.EOF And .BOF) Then
        fUsed = True
        Do Until .EOF
          mfrmUse.AddToList .Fields("UType").value & " : " & .Fields("Name").value
          .MoveNext
        Loop
        .Close
      End If
    End With
    Set adoRes = Nothing

  ' Check that it is not used by an email element in a workflow.
  sSQL = "SELECT DISTINCT tmpWorkflowElements.workflowID," & _
    "   tmpWorkflowElements.caption" & _
    " FROM tmpWorkflowElements" & _
    " WHERE tmpWorkflowElements.emailID = " & Trim(Str(Me.EmailID)) & _
    "   OR tmpWorkflowElements.emailCCID = " & Trim(Str(Me.EmailID))

  Set rsDAOCheck = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsDAOCheck.BOF And rsDAOCheck.EOF) Then
    Do Until rsDAOCheck.EOF
      recWorkflowEdit.Index = "idxWorkflowID"
      recWorkflowEdit.Seek "=", rsDAOCheck.Fields("workflowID")

      If Not recWorkflowEdit.NoMatch Then
        If recWorkflowEdit.Fields("deleted").value = False Then
          fUsed = True
          mfrmUse.AddToList ("Workflow : " & recWorkflowEdit.Fields("name").value & " <'" & rsDAOCheck.Fields("caption") & "' email element>")
        End If
      End If

      rsDAOCheck.MoveNext
    Loop
  End If
  'Close temporary recordset
  rsDAOCheck.Close

  ' Check that the email is not used in any Module definitions.
  sSQL = "SELECT DISTINCT moduleKey, parameterkey" & _
    " FROM tmpModuleSetup" & _
    " WHERE parameterType = '" & gsPARAMETERTYPE_EMAILID & "'" & _
    " AND parameterValue = '" & Trim(Str(Me.EmailID)) & "'"
  Set rsModules = daoDb.OpenRecordset(sSQL, _
    dbOpenForwardOnly, dbReadOnly)
  If Not (rsModules.BOF And rsModules.EOF) Then
    fUsed = True

    Do Until rsModules.EOF
      If (rsModules!moduleKey = gsMODULEKEY_WORKFLOW) And (Not bWF) Then
          bWF = True
          sModuleName = "Workflow"
          mfrmUse.AddToList (sModuleName & " module setup")
      End If

      rsModules.MoveNext
    Loop
  End If
  ' Close the recordset.
  rsModules.Close

TidyUpAndExit:
  ' Disassociate object variables.
  Set rsDAOCheck = Nothing

  EmailIsUsed = fUsed
  Exit Function

ErrorTrap:
  fUsed = True
  Resume TidyUpAndExit

End Function


