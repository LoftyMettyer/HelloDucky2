VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmailRecipients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCol As Collection

Public Function Populate_Collection(pstrEmailGroupIDs As String) As Boolean
  
  Dim strSQL As String
  Dim rsRecipients As ADODB.Recordset
  Dim objEmailRecipient As clsEmailRecipient
  
  strSQL = vbNullString
  strSQL = strSQL & "SELECT [ASRSysEmailGroupName].[Name] AS 'GroupName', " & vbCrLf
  strSQL = strSQL & "       [ASRSysEmailAddress].[Fixed], " & vbCrLf
  strSQL = strSQL & "       [ASRSysEmailAddress].[Name] " & vbCrLf
  strSQL = strSQL & "FROM [ASRSysEmailGroupName]" & vbCrLf
  strSQL = strSQL & "       INNER JOIN [ASRSysEmailGroupItems] " & vbCrLf
  strSQL = strSQL & "       ON [ASRSysEmailGroupItems].[EmailGroupID] = [ASRSysEmailGroupName].[EmailGroupID]" & vbCrLf
  strSQL = strSQL & "         INNER JOIN [ASRSysEmailAddress] " & vbCrLf
  strSQL = strSQL & "         ON [ASRSysEmailAddress].[EmailID] = [ASRSysEmailGroupItems].[EmailDefID]" & vbCrLf
  strSQL = strSQL & "WHERE [ASRSysEmailGroupName].[EmailGroupID] IN (" & pstrEmailGroupIDs & ")" & vbCrLf
  strSQL = strSQL & "ORDER BY [ASRSysEmailAddress].[Name]"
 
  Set rsRecipients = datGeneral.GetReadOnlyRecords(strSQL)
  
  With rsRecipients
  
    If Not (.BOF And .EOF) Then
    
      Do Until .EOF
        Add .Fields("Fixed").Value, .Fields("Name").Value
        
        .MoveNext
      Loop
    
    End If
    .Close
  End With
  
  Set rsRecipients = Nothing
  Set objEmailRecipient = Nothing
  strSQL = vbNullString
  
  Populate_Collection = True
  
End Function

Public Function Add(pstrEmailAddr As String, pstrEmailName As String) As clsEmailRecipient

  Dim objNewMember As New clsEmailRecipient
  
  objNewMember.FixedEmail = pstrEmailAddr
  objNewMember.EmailName = pstrEmailName
  
  mCol.Add objNewMember, CStr(pstrEmailAddr)
  
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
  
  On Error Resume Next
  
  ' Provide a reference to a specific item in the collection
  Set Item = mCol(sColTypeAndID)

End Property

Public Property Get Count() As Long
    
  ' Provide number of objects in the collection
  Count = mCol.Count
    
End Property

Public Sub Remove(pstrFixedEmail As String)

  ' Remove a specific object from the collection
  mCol.Remove pstrFixedEmail
    
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

