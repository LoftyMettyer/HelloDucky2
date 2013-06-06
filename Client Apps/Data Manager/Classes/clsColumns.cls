VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class for use with the CustomReports functionality
' Works with the CustomReportColDetail Class definition
'
' RH Jan 2000

Option Explicit

' Hold local collection
Private mCol As Collection

Public Property Get Collection() As Collection
  Set Collection = mCol
End Property
Public Property Let Collection(ByVal pcolNewValue As Collection)
  Set mCol = pcolNewValue
End Property


Public Function Add(sType As String, lID As Long, sHeading As String, _
                lSize As Long, iDecPlaces As Integer, bAverage As Boolean, _
                bCount As Boolean, bTotal As Boolean, bIsNumeric As Boolean, _
                bHidden As Boolean, bGroupWithNext As Boolean) As clsColumn
                    
  ' Add a new object to the collection
  ' Called when a user adds a column/expression to a report definition
  
  Dim obNewMember As New clsColumn
  
  obNewMember.ColType = sType
  obNewMember.ID = lID
  obNewMember.Heading = sHeading
  obNewMember.Size = lSize
  obNewMember.DecPlaces = iDecPlaces
  obNewMember.Average = bAverage
  obNewMember.Count = bCount
  obNewMember.Total = bTotal
  obNewMember.Hidden = bHidden
  obNewMember.GroupWithNext = bGroupWithNext
  obNewMember.IsNumeric = bIsNumeric
  
  mCol.Add obNewMember, sType + CStr(lID)
  
  Set Add = obNewMember
  
  Set obNewMember = Nothing
                    
End Function

Public Property Get Item(sColTypeAndID As String) As clsColumn
    
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
  ' Called when a user removes a column/expression from a report definition
  
  mCol.Remove sColTypeAndID
    
End Sub

Public Sub RemoveAll()

  ' Remove all object from the collection
  ' Called when a user removes all columns/expressions from a report definition

  Do While mCol.Count > 1
    'mCol.Remove 0
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

