VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"SecurityGroup"
Attribute VB_Ext_KEY = "Member0" ,"SecurityGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Local variable to hold collection.
Private mvar_colGroups As Collection
Private mlngBottom As Long

Public Function Add(psName As String, _
  pfChanged As Boolean, _
  pfDeleteGroup As Boolean, _
  pfNewGroup As Boolean, _
  Optional psKey As String) As SecurityGroup

  Dim objNewMember As SecurityGroup
  
  Dim objUsers As SecurityUsers
  Dim objTables As SecurityTables
  Dim objViews As SecurityTables
  Dim objSystemPermissions As clsSystemPermissions
  
  ' Instantiate new collections for the group definition.
  Set objUsers = New SecurityUsers
  Set objTables = New SecurityTables
  Set objViews = New SecurityTables
  Set objSystemPermissions = New clsSystemPermissions
  
  ' Instantiate a new security group.
  Set objNewMember = New SecurityGroup

  With objNewMember
    ' Set the properties passed into the method.
    Set .Users = objUsers
    Set .Tables = objTables
    Set .Views = objViews
    Set .SystemPermissions = objSystemPermissions
    
    .Name = psName
    .Changed = pfChanged
    .DeleteGroup = pfDeleteGroup
    .RequireLogout = pfDeleteGroup    'MH20010410
    .NewGroup = pfNewGroup
    .Initialised = False
  End With
  
  If Len(psKey) = 0 Then
    mvar_colGroups.Add objNewMember
  Else
    mvar_colGroups.Add objNewMember, psKey
  End If

  ' Return the new group object.
  Set Add = objNewMember
  Set objNewMember = Nothing
  Set objUsers = Nothing
  Set objTables = Nothing
  Set objViews = Nothing
  Set objSystemPermissions = Nothing

End Function

Public Property Get Item(ByRef pvIndexKey As Variant) As SecurityGroup
Attribute Item.VB_UserMemId = 0
  Set Item = mvar_colGroups(pvIndexKey)
  
End Property

Public Property Get Count() As Long
  Count = mvar_colGroups.Count
  
End Property

Public Sub Remove(pvIndexKey As Variant)
  mvar_colGroups.Remove pvIndexKey
  
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mvar_colGroups.[_NewEnum]
  
End Property

Private Sub Class_Initialize()
  'creates the collection when this class is created
  Set mvar_colGroups = New Collection
  
End Sub

Private Sub Class_Terminate()
  ' Destroy collection when this class is terminated.
  Set mvar_colGroups = Nothing
  
End Sub

Public Sub Paste(ByVal pobjSecurityGroup As SecurityGroup)

  mvar_colGroups.Add pobjSecurityGroup, pobjSecurityGroup.Name

End Sub


Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
  
  'Dim mvar_objUsers As SecurityUsers
  
  Dim objGroup As SecurityGroup
  Dim objPrintDef As clsPrintDef
  
  Dim fOK As Boolean
  Dim iCount As Integer
  Dim oLoop As Integer
  Dim jLoop As Integer
  Dim iLoop As Integer
  Dim iLoop2 As Integer
  Dim temp As String
  Dim asGroups() As String
  Dim fFound As Boolean
  Dim fPrintGroup As Boolean
  
  fOK = True
  
  If fOK Then
    Set objPrintDef = New clsPrintDef
'
    If objPrintDef.IsOK Then
        With objPrintDef
    '        If .PrintStart(False) Then
                
                'NHRD28062004 Fault 8520
                '.PrintHeader "Group Permissions"
                
                ReDim asGroups(gObjGroups.Count)
                
                For iCount = 1 To gObjGroups.Count
                    asGroups(iCount) = gObjGroups.Item(iCount).Name
                Next
                
                For iCount = 1 To gObjGroups.Count
                    'inner loop
                    For jLoop = 1 To gObjGroups.Count - iCount
                        'compare adjacent pair of numbers to see if they are in wrong order
                        If UCase(asGroups(jLoop)) > UCase(asGroups(jLoop + 1)) Then
                        'pair in wrong order, swap pair of numbers
                        temp = asGroups(jLoop)
                        asGroups(jLoop) = asGroups(jLoop + 1)
                        asGroups(jLoop + 1) = temp
                        End If
                    Next jLoop
                        'end of inner loop
                Next iCount
                
                For iCount = 1 To gObjGroups.Count
                    
                    Set objGroup = gObjGroups(asGroups(iCount))
                    If Not objGroup.DeleteGroup Then
                        fPrintGroup = Not gasPrintOptions(1).PrintRPaneGROUPS
                        
                        If Not fPrintGroup Then
                            For iLoop2 = 1 To UBound(gasPrintGroups)
                                If gasPrintGroups(iLoop2) = objGroup.Name Then
                                    fPrintGroup = True
                                    Exit For
                                End If
                            Next iLoop2
                        End If
                        
                        If fPrintGroup Then
                          
                          objGroup.PrintSecurity piLevel + 1
                          
                          If iCount <> gObjGroups.Count Then
                             'NHRD28062004 Fault 8520
'                            mlngBottom = CalculateBottomOfPage
'                            Printer.CurrentY = mlngBottom + 1
'                            CheckEndOfPage2 mlngBottom, True
                          End If
                        End If
                    End If
                Next iCount
       '      End If
        End With
     End If
  End If
  
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function
