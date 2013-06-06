VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBankHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' Hold local collection
Private mCol As Collection

Public Function Add(pstrRegion As String, pstrDescription As String, pdtHolidayDate As Date) As clsBankHoliday
                    
  ' Add a new object to the collection
  
  Dim objNewMember As clsBankHoliday
  Set objNewMember = New clsBankHoliday
  
  objNewMember.Region = pstrRegion
  objNewMember.Description = pstrDescription
  objNewMember.HolidayDate = pdtHolidayDate
  
  mCol.Add objNewMember
  
  Set Add = objNewMember
  
  Set objNewMember = Nothing
                    
End Function

Public Property Get Collection() As Collection
  Set Collection = mCol
End Property

Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
End Property

Public Property Get Item(sColTypeAndID As String) As clsBankHoliday
    
  ' Provide a reference to a specific item in the collection
  
  On Error Resume Next
  
  Set Item = mCol(sColTypeAndID)

End Property

Public Property Get Count() As Long
    
  ' Provide number of objects in the collection
  
  Count = mCol.Count
    
End Property

Public Sub Remove(sColTypeAndID As String)

  ' Remove a specific object from the collection
  
  mCol.Remove sColTypeAndID
    
End Sub

Public Sub RemoveAll()

  ' Remove all object from the collection

  Do While mCol.Count > 1
    mCol.Remove mCol.Count
  Loop

End Sub

Private Sub Class_Initialize()
    
  Set mCol = New Collection
    
End Sub

Private Sub Class_Terminate()

  Set mCol = Nothing
    
End Sub

Public Property Get NewEnum() As IUnknown
    
  Set NewEnum = mCol.[_NewEnum]

End Property



