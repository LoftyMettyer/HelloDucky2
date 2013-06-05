VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDiaryLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Properties.
Private gLngDiaryLinkID As Long
Private gLngColumnID As Long
Private gsComment As String
Private giOffset As Integer
Private giPeriod As Integer
Private gfReminder As Boolean
Private glFilterID As Long
Private gdtEffectiveDate As Date
Private mfCheckLeavingDate As Boolean


Public Property Get ColumnID() As Long
  ' Return the Column ID property.
  ColumnID = gLngColumnID
  
End Property

Public Property Let ColumnID(ByVal pLngNewValue As Long)
  ' Set the Column ID property.
  gLngColumnID = pLngNewValue

End Property

Public Property Get DiaryLinkId() As Long
  ' Return the Diary Link ID property.
  DiaryLinkId = gLngDiaryLinkID

End Property

Public Property Let DiaryLinkId(ByVal pLngNewValue As Long)
  ' Set the Diary Link ID property.
  gLngDiaryLinkID = pLngNewValue

End Property

Public Property Get Comment() As String
  ' Return the Comment property.
  Comment = gsComment

End Property

Public Property Let Comment(ByVal psNewValue As String)
  ' Set the Comment property.
  gsComment = psNewValue

End Property

Public Property Get Offset() As Integer
  ' Return the Offset property.
  Offset = giOffset

End Property

Public Property Let Offset(ByVal piNewValue As Integer)
  ' Set the Offset property.
  giOffset = piNewValue

End Property

Public Property Get Period() As Integer
  ' Return the Period property.
  Period = giPeriod

End Property

Public Property Let Period(ByVal piNewValue As Integer)
  ' Set the Period property.
  giPeriod = piNewValue

End Property

Public Property Get Reminder() As Boolean
  ' Return the Reminder property.
  Reminder = gfReminder

End Property

Public Property Let Reminder(ByVal pfNewValue As Boolean)
  ' Set the Reminder property.
  gfReminder = pfNewValue

End Property
Public Function ReadDiaryLink() As Boolean
  ' Read the Diary Link definition from the Diary Link recordset.
  On Error GoTo ErrorTrap
  
  Dim fReadOk As Boolean
  
  fReadOk = False
  
  ' Read the basic component details.
  If FindDiaryLink Then

    With recDiaryEdit

      gLngColumnID = IIf(IsNull(.Fields("ColumnID")), 0, .Fields("ColumnID"))
      gsComment = IIf(IsNull(.Fields("comment")), vbNullString, .Fields("comment"))
      giOffset = IIf(IsNull(.Fields("offset")), 0, .Fields("offset"))
      giPeriod = IIf(IsNull(.Fields("period")), iTimePeriodDays, .Fields("period"))
      gfReminder = IIf(IIf(IsNull(.Fields("reminder")), False, .Fields("reminder")), True, False)
      glFilterID = IIf(Trim(!FilterID) = vbNullString, 0, !FilterID)
      gdtEffectiveDate = IIf(IsNull(!EffectiveDate), #1/1/1980#, !EffectiveDate)
      mfCheckLeavingDate = IIf(IsNull(!CheckLeavingDate), True, !CheckLeavingDate)

      fReadOk = True

    End With
  End If
   
TidyUpAndExit:
  ReadDiaryLink = fReadOk
  Exit Function

ErrorTrap:
  fReadOk = False
  Err = False
  Resume TidyUpAndExit

End Function

Public Function WriteDiaryLink() As Boolean
  ' Write the Diary Link to the recordset.
  On Error GoTo ErrorTrap
  
  Dim lngNewID As Long
  Dim fOK As Boolean
  
  fOK = True
  
  lngNewID = UniqueColumnValue("tmpDiary", "diaryID")
  gLngDiaryLinkID = lngNewID
  
  With recDiaryEdit

    'Add a record for new expression
    .AddNew

    !diaryID = gLngDiaryLinkID
    !ColumnID = gLngColumnID
    !Comment = gsComment
    !Offset = giOffset
    !Period = giPeriod
    !Reminder = gfReminder
    !FilterID = glFilterID
    !EffectiveDate = gdtEffectiveDate
    !CheckLeavingDate = mfCheckLeavingDate

    .Update

  End With
  
TidyUpAndExit:
  WriteDiaryLink = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  Resume TidyUpAndExit

End Function

Private Function FindDiaryLink() As Boolean
  ' Locate the current Diary Link in the recordset.
  On Error GoTo ErrorTrap
  
  With recDiaryEdit
    .Index = "idxDiaryLinkID"
    .Seek "=", gLngDiaryLinkID
    FindDiaryLink = (Not .NoMatch)
  End With
  
  Exit Function
  
ErrorTrap:
  FindDiaryLink = False

End Function


Public Property Get FilterID() As Long
  FilterID = glFilterID
End Property

Public Property Let FilterID(ByVal lngNewValue As Long)
  glFilterID = lngNewValue
End Property

Public Property Get EffectiveDate() As Date
  EffectiveDate = gdtEffectiveDate
End Property

Public Property Let EffectiveDate(ByVal dtNewValue As Date)
  gdtEffectiveDate = dtNewValue
End Property
  
Public Property Get CheckLeavingDate() As Boolean
  CheckLeavingDate = mfCheckLeavingDate
End Property
  
Public Property Let CheckLeavingDate(ByVal blnNewValue As Boolean)
  mfCheckLeavingDate = blnNewValue
End Property
