VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngLinkID As Long
Private mstrTitle As String
Private mlngFilterID As Long
Private mdtEffectiveDate As Date
Private mstrAttachment As String
Private mlngTableID As Long
Private mintType As Integer

Private mblnRecordInsert As Boolean
Private mblnRecordUpdate As Boolean
Private mblnRecordDelete As Boolean

Private mlngDateColumnID As Long
Private mlngDateOffset As Long
Private mintDatePeriod As Integer
Private mblnDateAmendment As Boolean

Private mobjSubjectContent As clsLinkContent
Private mobjBodyContent As clsLinkContent
Private mcolRecipients(2) As Collection

Private mcolColumns As Collection

Public Function IsColumnSelected(lngColumnID As Long) As Boolean

  On Local Error GoTo LocalErr

  Dim lngTest As Long
  lngTest = Columns(CStr(lngColumnID))

  IsColumnSelected = True

Exit Function

LocalErr:
  IsColumnSelected = False

End Function


Public Property Get Columns() As Collection
  If mcolColumns Is Nothing Then
    Set mcolColumns = New Collection
  End If
  Set Columns = mcolColumns
End Property

Public Property Let Columns(value As Collection)
  Set mcolColumns = value
End Property

Public Property Get RecipientsTo() As Collection
  If mcolRecipients(0) Is Nothing Then
    Set mcolRecipients(0) = New Collection
  End If
  Set RecipientsTo = mcolRecipients(0)
End Property

Public Property Let RecipientsTo(value As Collection)
  Set mcolRecipients(0) = value
End Property

Public Property Get RecipientsCc() As Collection
  If mcolRecipients(1) Is Nothing Then
    Set mcolRecipients(1) = New Collection
  End If
  Set RecipientsCc = mcolRecipients(1)
End Property

Public Property Let RecipientsCc(value As Collection)
  Set mcolRecipients(1) = value
End Property

Public Property Get RecipientsBcc() As Collection
  If mcolRecipients(2) Is Nothing Then
    Set mcolRecipients(2) = New Collection
  End If
  Set RecipientsBcc = mcolRecipients(2)
End Property

Public Property Let RecipientsBcc(value As Collection)
  Set mcolRecipients(2) = value
End Property


Public Property Get TableID() As Long
  TableID = mlngTableID
End Property

Public Property Let TableID(ByVal lngNewValue As Long)
  mlngTableID = lngNewValue
End Property

Public Property Get LinkType() As Integer
  LinkType = mintType
End Property

Public Property Let LinkType(ByVal intNewValue As Integer)
  mintType = intNewValue
End Property


Public Property Get LinkID() As Long
  LinkID = mlngLinkID
End Property

Public Property Let LinkID(ByVal lngNewValue As Long)
  mlngLinkID = lngNewValue
End Property

Public Property Get DateColumnID() As Long
  DateColumnID = mlngDateColumnID
End Property

Public Property Let DateColumnID(ByVal lngNewValue As Long)
  mlngDateColumnID = lngNewValue
End Property

Public Property Get Title() As String
  Title = mstrTitle
End Property

Public Property Let Title(ByVal strNewValue As String)
  mstrTitle = strNewValue
End Property

Public Property Get FilterID() As Long
  FilterID = mlngFilterID
End Property

Public Property Let FilterID(ByVal lngNewValue As Long)
  mlngFilterID = lngNewValue
End Property

'Public Property Get Immediate() As Boolean
'  Immediate = mblnImmediate
'End Property
'
'Public Property Let Immediate(ByVal blnNewValue As Boolean)
'  mblnImmediate = blnNewValue
'End Property

Public Property Get DateOffset() As Long
  DateOffset = mlngDateOffset
End Property

Public Property Let DateOffset(ByVal lngNewValue As Long)
  mlngDateOffset = lngNewValue
End Property

Public Property Get DatePeriod() As Integer
  DatePeriod = mintDatePeriod
End Property

Public Property Let DatePeriod(ByVal intNewValue As Integer)
  mintDatePeriod = intNewValue
End Property

Public Property Get DateAmendment() As Boolean
  DateAmendment = mblnDateAmendment
End Property

Public Property Let DateAmendment(ByVal blnNewValue As Boolean)
  mblnDateAmendment = blnNewValue
End Property

Public Property Get EffectiveDate() As Date
  EffectiveDate = mdtEffectiveDate
End Property

Public Property Let EffectiveDate(ByVal dtNewValue As Date)
  mdtEffectiveDate = dtNewValue
End Property

'Public Property Get Subject() As String
'  Subject = mstrSubject
'End Property
'
'Public Property Let Subject(ByVal strNewValue As String)
'  mstrSubject = strNewValue
'End Property


'Public Property Get Importance() As Integer
'  Importance = mintImportance
'End Property
'Public Property Let Importance(ByVal intNewValue As Integer)
'  mintImportance = intNewValue
'End Property
'
'
'Public Property Get Sensitivity() As Integer
'  Sensitivity = mintSensitivity
'End Property
'Public Property Let Sensitivity(ByVal intNewValue As Integer)
'  mintSensitivity = intNewValue
'End Property
'
'
'Public Property Get IncUsername() As Boolean
'  IncUsername = mblnIncUserName
'End Property
'Public Property Let IncUsername(ByVal blnNewValue As Boolean)
'  mblnIncUserName = blnNewValue
'End Property
'
'
'Public Property Get IncRecordDesc() As Boolean
'  IncRecordDesc = mblnIncRecordDesc
'End Property
'Public Property Let IncRecordDesc(ByVal blnNewValue As Boolean)
'  mblnIncRecordDesc = blnNewValue
'End Property
'
'
'Public Property Get IncColumnDetails() As Boolean
'  IncColumnDetails = mblnIncColumnDetails
'End Property
'Public Property Let IncColumnDetails(ByVal blnNewValue As Boolean)
'  mblnIncColumnDetails = blnNewValue
'End Property
'
'
'Public Property Get Text() As String
'  Text = mstrText
'End Property
'Public Property Let Text(ByVal strNewValue As String)
'  mstrText = strNewValue
'End Property


Public Property Get Attachment() As String
  Attachment = mstrAttachment
End Property
Public Property Let Attachment(ByVal vNewValue As String)
  mstrAttachment = vNewValue
End Property


'Public Property Get Recipients() As Variant
'  Recipients = mlngRecipients()
'End Property
'Public Property Let Recipients(ByVal vNewValue As Variant)
'  mlngRecipients() = vNewValue
'End Property
'
'
'Public Property Get SendTypes() As Variant
'  SendTypes = mlngSendTypes()
'End Property
'Public Property Let SendTypes(ByVal vNewValue As Variant)
'  mlngSendTypes() = vNewValue
'End Property



Public Function ReadEmailLink() As Boolean
  ' Read the Email Link definition from the Email Link recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  Dim intCount As Integer
  
  fReadOk = False
  
  ' Read the basic component details.
  'If FindEmailLink Then
    
    With recEmailLinksEdit

      mlngLinkID = IIf(IsNull(!LinkID), 0, !LinkID)
      mstrTitle = IIf(IsNull(!Title), vbNullString, !Title)
      mlngFilterID = IIf(IsNull(!FilterID), 0, !FilterID)
      mdtEffectiveDate = IIf(IsNull(!EffectiveDate), vbNullString, !EffectiveDate)
      mstrAttachment = IIf(IsNull(!Attachment), vbNullString, !Attachment)
      mlngTableID = IIf(IsNull(!TableID), 0, !TableID)
      mintType = IIf(IsNull(!Type), 0, !Type)
      
      mlngDateColumnID = IIf(IsNull(!DateColumnID), 0, !DateColumnID)
      mlngDateOffset = IIf(IsNull(!DateOffset), 0, !DateOffset)
      mintDatePeriod = IIf(IsNull(!DatePeriod), 0, !DatePeriod)
      mblnDateAmendment = IIf(IsNull(!DateAmendment), 1, !DateAmendment)
  
      mblnRecordUpdate = IIf(IsNull(!RecordUpdate), True, !RecordUpdate)
      mblnRecordInsert = IIf(IsNull(!RecordInsert), True, !RecordInsert)
      mblnRecordDelete = IIf(IsNull(!RecordDelete), True, !RecordDelete)
      
      'mstrSubject = IIf(IsNull(!Subject), vbNullString, !Subject)
      'mlngOffset = IIf(IsNull(!Offset), 0, !Offset)
      'mintOffsetPeriod = IIf(IsNull(!Period), 0, !Period)
      'mblnImmediate = IIf(IsNull(!Immediate), True, !Immediate)
      'mblnIncRecordDesc = IIf(IsNull(!IncRecDesc), True, !IncRecDesc)
      'mblnIncColumnDetails = IIf(IsNull(!IncColDetail), True, !IncColDetail)
      'mblnIncUserName = IIf(IsNull(!IncUsername), True, !IncUsername)
      'mstrText = IIf(IsNull(!Body), vbNullString, !Body)

      'MH20090520
      SubjectContent.ReadDetail !SubjectContentID
      BodyContent.ReadDetail !BodyContentID

    End With
  
  
    ClearRecipients
    With recEmailRecipientsEdit

      .Index = "idxLinkID"
      .Seek ">=", mlngLinkID
      
      intCount = 0
      If Not .NoMatch Then
        
        Do While Not .EOF
          
          If !LinkID <> mlngLinkID Then
            Exit Do
          End If
    
          mcolRecipients(!Mode).Add CLng(!RecipientID), CStr(!RecipientID)
          .MoveNext
        Loop
    
      End If

    End With


    ClearColumns
    With recEmailLinksColumnsEdit

      .Index = "idxLinkID"
      .Seek ">=", mlngLinkID
      
      intCount = 0
      If Not .NoMatch Then
        
        Do While Not .EOF
          
          If !LinkID <> mlngLinkID Then
            Exit Do
          End If
    
          Columns.Add CLng(!ColumnID), CStr(!ColumnID)
          .MoveNext
        Loop
    
      End If

    End With



'    With recEmailAttachmentsEdit
'
'      .Index = "idxLinkID"
'      .Seek ">=", mlngLinkID
'
'      intCount = 0
'      ReDim mstrAttachment(0) As String
'
'      If Not .NoMatch Then
'
'        Do While Not .EOF
'
'          If !LinkID <> mlngLinkID Then
'            Exit Do
'          End If
'
'          intCount = intCount + 1
'          ReDim Preserve mstrAttachment(intCount)
'          mstrAttachment(intCount) = !FileName
'          .MoveNext
'        Loop
'
'      End If
'
'    End With


    fReadOk = True
  
  'End If
   
TidyUpAndExit:
  ReadEmailLink = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function


Private Function FindEmailLink() As Boolean
  ' Locate the current Email Link in the recordset.
  On Error GoTo ErrorTrap
  
  With recEmailLinksEdit
    .Index = "idxID"
    .Seek "=", mlngLinkID
    FindEmailLink = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindEmailLink = False

End Function


Public Function WriteEMailLink() As Boolean
  ' Write the Email Link to the recordset.
  On Error GoTo ErrorTrap
  
  Dim lngNewID As Long
  Dim fOK As Boolean
  Dim intCount As Integer
  
  fOK = True
      
      
  With recEmailLinksEdit

    If mlngLinkID = 0 Then
      mlngLinkID = UniqueColumnValue("tmpEmailLinks", "LinkID")
    End If

    .AddNew
    !Deleted = False

    !LinkID = mlngLinkID
    !Title = mstrTitle
    !FilterID = mlngFilterID
    !EffectiveDate = mdtEffectiveDate
    !Attachment = mstrAttachment
    !TableID = mlngTableID
    !Type = mintType

    !SubjectContentID = SubjectContent.WriteDetail
    !BodyContentID = BodyContent.WriteDetail

    !DateColumnID = mlngDateColumnID
    !DateOffset = mlngDateOffset
    !DatePeriod = mintDatePeriod
    !DateAmendment = mblnDateAmendment
    
    !RecordUpdate = mblnRecordUpdate
    !RecordDelete = mblnRecordDelete
    !RecordInsert = mblnRecordInsert

    .Update
    
  End With


  With recEmailRecipientsEdit

'    For intCount = 0 To UBound(mlngRecipients) Step 1
'      .AddNew
'      !LinkID = mlngLinkID
'      !RecipientID = mlngRecipients(intCount)
'      !Mode = mlngSendTypes(intCount)
'      .Update
'    Next

    Dim lngModeIndex As Long
    Dim lngRecipientIndex As Long
    Dim strSQL As String
    
    strSQL = "DELETE FROM tmpEmailLinksRecipients WHERE LinkID = " & CStr(mlngLinkID)
    daoDb.Execute strSQL, dbFailOnError
    
    For lngModeIndex = 0 To 2
      For lngRecipientIndex = 1 To mcolRecipients(lngModeIndex).Count
        .AddNew
        !LinkID = mlngLinkID
        !RecipientID = CLng(mcolRecipients(lngModeIndex).Item(lngRecipientIndex))
        !Mode = lngModeIndex
        .Update
      Next
    Next


  End With


  With recEmailLinksColumnsEdit

'    For intCount = 0 To UBound(mlngRecipients) Step 1
'      .AddNew
'      !LinkID = mlngLinkID
'      !RecipientID = mlngRecipients(intCount)
'      !Mode = mlngSendTypes(intCount)
'      .Update
'    Next

    Dim lngColumnIndex As Long
    'Dim strSQL As String
    
    strSQL = "DELETE FROM tmpEmailLinksColumns WHERE LinkID = " & CStr(mlngLinkID)
    daoDb.Execute strSQL, dbFailOnError
    
    For lngColumnIndex = 1 To Columns.Count
      .AddNew
      !LinkID = mlngLinkID
      !ColumnID = CLng(Columns.Item(lngColumnIndex))
      .Update
    Next
  
  End With
  
  
  '  With recEmailAttachmentsEdit
'
'    For intCount = 1 To UBound(mstrAttachment)
'      .AddNew
'      !LinkID = mlngLinkID
'      !FileName = mstrAttachment(intCount)
'      .Update
'    Next
'
'  End With

TidyUpAndExit:
  WriteEMailLink = fOK
  Exit Function

ErrorTrap:
  If ASRDEVELOPMENT Then
    MsgBox Err.Description, vbCritical, "ASRDEVELOPMENT"
    Stop
  End If
  
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function

Private Sub Class_Initialize()
  'ReDim mstrAttachment(0)
  'ReDim mlngRecipients(0)
  'ReDim mlngSendTypes(0)
  ClearRecipients
  'mblnImmediate = True
End Sub


Private Sub ClearRecipients()
  Set mcolRecipients(0) = Nothing
  Set mcolRecipients(1) = Nothing
  Set mcolRecipients(2) = Nothing
  Set mcolRecipients(0) = New Collection
  Set mcolRecipients(1) = New Collection
  Set mcolRecipients(2) = New Collection
End Sub


Private Sub ClearColumns()
  Set mcolColumns = Nothing
  Set mcolColumns = New Collection
End Sub


Public Function CloneLink(Optional pavCloneRegister As Variant) As clsEmailLink

  Dim objNewLink As clsEmailLink
  
  Set objNewLink = New clsEmailLink

  With objNewLink
    .Title = Me.Title
    '.FilterID = Me.FilterID
    .EffectiveDate = Me.EffectiveDate
    .Attachment = Me.Attachment
    '.TableID = Me.TableID
    .LinkType = Me.LinkType
    '.DateColumnID = Me.DateColumnID
    .DateOffset = Me.DateOffset
    .DatePeriod = Me.DatePeriod
    .DateAmendment = Me.DateAmendment
    .RecordDelete = Me.RecordDelete
    .RecordInsert = Me.RecordInsert
    .RecordUpdate = Me.RecordUpdate
    
    '.SubjectContent.ReadDetail(
    
  End With
  
  Set CloneLink = objNewLink
  Set objNewLink = Nothing

End Function

Public Property Get RecordInsert() As Boolean
  RecordInsert = mblnRecordInsert
End Property

Public Property Let RecordInsert(ByVal bNewValue As Boolean)
  mblnRecordInsert = bNewValue
End Property

Public Property Get RecordUpdate() As Boolean
  RecordUpdate = mblnRecordUpdate
End Property

Public Property Let RecordUpdate(ByVal bNewValue As Boolean)
  mblnRecordUpdate = bNewValue
End Property

Public Property Get RecordDelete() As Boolean
  RecordDelete = mblnRecordDelete
End Property

Public Property Let RecordDelete(ByVal bNewValue As Boolean)
  mblnRecordDelete = bNewValue
End Property


Public Function GetNewLinkID(mvarEmailLinks As Collection) As Long

  Dim objNewLink As clsEmailLink
  Dim lngNewID As Long

  On Local Error GoTo LocalErr

  Set objNewLink = New clsEmailLink

  lngNewID = UniqueColumnValue("tmpEmailLinks", "LinkID")

  
  If Not mvarEmailLinks Is Nothing Then
    Do While lngNewID < 99999
      Set objNewLink = mvarEmailLinks("ID" & lngNewID)
      lngNewID = lngNewID + 1
    Loop
  End If

  GetNewLinkID = lngNewID

Exit Function

LocalErr:
  GetNewLinkID = lngNewID

End Function


'MH20090520
Public Property Get SubjectContent() As clsLinkContent
  If mobjSubjectContent Is Nothing Then
    Set mobjSubjectContent = New clsLinkContent
  End If
  Set SubjectContent = mobjSubjectContent
End Property

Public Property Let SubjectContent(value As clsLinkContent)
  Set mobjSubjectContent = value
End Property

Public Property Get BodyContent() As clsLinkContent
  If mobjBodyContent Is Nothing Then
    Set mobjBodyContent = New clsLinkContent
  End If
  Set BodyContent = mobjBodyContent
End Property

Public Property Let BodyContent(value As clsLinkContent)
  Set mobjBodyContent = value
End Property

Private Sub Class_Terminate()
  Set mcolColumns = Nothing
  Set mcolRecipients(0) = Nothing
  Set mcolRecipients(1) = Nothing
  Set mcolRecipients(2) = Nothing
End Sub
