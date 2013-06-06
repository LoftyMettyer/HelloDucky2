VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemPermissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Local variable to hold collection.
Private mvar_colPermissions As Collection

Public Sub Clear()
  Do While mvar_colPermissions.Count > 0
    mvar_colPermissions.Remove 1
  Loop
  
End Sub
Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
'
  Dim fOK As Boolean
  Dim objPermission As clsSystemPermission
'
  fOK = True

  For Each objPermission In mvar_colPermissions
    objPermission.PrintSecurity piLevel + 1
    'Exit For
  Next objPermission
  
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
Public Function Add(plngItemID As Long, _
  pfAllowed As Boolean, _
  psItemKey As String, _
  psCategoryKey As String) As clsSystemPermission
  ' Create a new System Permission object.
  Dim objNewMember As clsSystemPermission
    
  Set objNewMember = New clsSystemPermission

  ' Set the properties passed into the method.
  With objNewMember
    .ItemID = plngItemID
    .Allowed = pfAllowed
    .ItemKey = psItemKey
    .CategoryKey = psCategoryKey
  End With

  mvar_colPermissions.Add objNewMember, "P_" & psCategoryKey & "_" & psItemKey

  ' Return the object created.
  Set Add = objNewMember
    
  ' Disassociate object variables.
  Set objNewMember = Nothing

End Function

Public Function IsValid(ByRef pvIndexKey As Variant) As Boolean
  ' Return TRUE if the given system privilege exists in the collection.
  Dim Item As clsSystemPermission
  
  On Error GoTo err_IsValid
  
  Set Item = mvar_colPermissions(pvIndexKey)
  IsValid = True
  Set Item = Nothing
  
  Exit Function
  
err_IsValid:
  IsValid = False
  
End Function


Public Property Get Item(ByRef pvIndexKey As Variant) As clsSystemPermission
  Set Item = mvar_colPermissions(pvIndexKey)
  
End Property

Public Property Get Count() As Long
  Count = mvar_colPermissions.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colPermissions.Remove pvIndexKey
  
End Sub


Private Sub Class_Initialize()
  ' Create the collection when this class is created.
  Set mvar_colPermissions = New Collection

End Sub


Private Sub Class_Terminate()
  ' Destroy collection when this class is terminated.
    Set mvar_colPermissions = Nothing

End Sub


Public Function Clone() As clsSystemPermissions

  Dim objSrcPermission As clsSystemPermission
  Dim objDestPermission As clsSystemPermission

  Set Clone = New clsSystemPermissions

  For Each objSrcPermission In mvar_colPermissions
    With objSrcPermission
      Set objDestPermission = Clone.Add(.ItemID, .Allowed, .ItemKey, .CategoryKey)
    End With
  Next

  Set objSrcPermission = Nothing
  
End Function

