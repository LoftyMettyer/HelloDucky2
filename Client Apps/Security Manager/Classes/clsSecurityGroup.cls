VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SecurityMenus"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member1" ,"SecuritySysPrivileges"
Attribute VB_Ext_KEY = "Member2" ,"SecurityTables"
Attribute VB_Ext_KEY = "Member3" ,"SecurityUsers"
Attribute VB_Ext_KEY = "Member4" ,"SecurityViews"
 
Option Explicit

'set this to 0 to disable debug code in this class
#Const DebugMode = 0

#If DebugMode Then
  ' Local variable to hold the serialized class ID that was created in Class_Initialize.
  Private mvar_lngClassDebugID As Long
#End If

' Local variables to hold property values.
Private mvar_sName As String
Private mvar_objUsers As SecurityUsers
Private mvar_objTables As SecurityTables
Private mvar_objViews As SecurityTables
Private mvar_objSystemPermissions As clsSystemPermissions
Private mvar_objColumns As SecurityColumns
Private mvar_sOriginalName As String

Private mvar_fNewGroup As Boolean
Private mvar_fChanged As Boolean
Private mvar_fDeleteGroup As Boolean
Private mvar_fInitialised As Boolean
Private mvar_fUsersInitialised As Boolean
Private mvar_fRequireLogout As Boolean
Dim iCount As Integer

Private Type PermissionInfo
   ItemID As Integer
   Category As String
   Item As String
   Allowed As Boolean
End Type

Private mvar_sAccessCopyGroup As String
Private mvar_avAccessConfiguration As Variant
Private mvar_sCopyGroup As String

Public Property Get Users_Initialised() As Boolean
  Users_Initialised = mvar_fUsersInitialised
  
End Property
Public Function PrintSecurity(piLevel As Integer) As Boolean
  ' Print the component definition to the printer object.
  On Error GoTo ErrorTrap
'
  Dim fOK As Boolean
  Dim objSecurityTable As SecurityTable
  Dim objSecurityTable2 As SecurityTable
  Dim objSecurityView As SecurityTable
  Dim objSecurityView2 As SecurityTable
  Dim objTable As SecurityTable
  
  Dim avTablesViews() As Variant
  Dim asViewInfo() As String
  Dim asUsersName() As String
  Dim asSysPerms() As PermissionInfo
  Dim asTemp() As PermissionInfo
  
  Dim sTableName As String
  Dim iIndex As Integer
  Dim sSQL As String
  Dim rsTables As New ADODB.Recordset
  Dim sFullViewName As String
  Dim iLoop As Integer
  Dim intSameLine As Integer
  Dim mlngBottom As Long
  Dim intOldFontSize As Integer
  Dim rsPermissionCategory As New ADODB.Recordset
  Dim iLoopItems As Integer
  Dim sKey As String
  Dim miColumnNumber As Integer
  Dim objIcon As IPictureDisp
  Dim intTopY As Integer
 
  fOK = True
  
    ' Position the printing.
    With Printer
      .FontSize = 8
      .CurrentX = giPRINT_XINDENT + (piLevel * giPRINT_XSPACE)
      .CurrentY = giPRINT_YINDENT '+ .CurrentY
      
      gstrPrintGroupName = mvar_sName
      .FontBold = True
      
      'NHRD19012006 Fault 10735
      If CheckToPrintHeader() Then
        If gasPrintOptions(1).PrintBlankVersion Then
            Printer.Print "Blank Version For Admin Purposes - GROUPNAME:  " + mvar_sName
        Else
            Printer.Print "GROUPNAME:  " + mvar_sName
        End If
      End If
      
      .FontBold = False
    End With
            
    '*******************************************************************************************
    'PRINT THE USER HEADING
    
    'JPD 20040308 Fault 8098
    'If (gasPrintOptions(1).PrintLPaneUSERS = True) Or _
            (gasPrintOptions(1).PrintRPaneUSERS = True) Or _
            (gasPrintOptions(1).PrintBlankVersion = True) Then
     If (gasPrintOptions(1).PrintLPaneUSERS = True) Or _
      (gasPrintOptions(1).PrintRPaneUSERS = True) Then
        
        With Printer
            .FontSize = 8
            .FontUnderline = True
            .CurrentX = giPRINT_XINDENT + ((piLevel + 1) * giPRINT_XSPACE)
            .CurrentY = .CurrentY + giPRINT_YSPACE
            Printer.Print "Users"
            .FontUnderline = False
            
            If mvar_objUsers.Count > 0 Then
                ReDim Preserve asUsersName(mvar_objUsers.Count)
                Dim jLoop As Integer
                Dim temp As String
                
                For iCount = 1 To mvar_objUsers.Count
                    asUsersName(iCount) = UCase(mvar_objUsers.Item(iCount).UserName)
                Next
                
                For iCount = 1 To mvar_objUsers.Count
                    'inner loop
                    For jLoop = 1 To mvar_objUsers.Count - iCount
                        'compare adjacent pair of numbers to see if they are in wrong order
                        If asUsersName(jLoop) > asUsersName(jLoop + 1) Then
                          'pair in wrong order, swap pair of numbers
                          temp = asUsersName(jLoop)
                          asUsersName(jLoop) = asUsersName(jLoop + 1)
                          asUsersName(jLoop + 1) = temp
                        End If
                        Next jLoop
                        'end of inner loop
                    Next iCount
                'Print USERS in alphabetical order using the pre-ordered loop
                For iLoop = 1 To mvar_objUsers.Count
                     mvar_objUsers.Item(asUsersName(iLoop)).PrintSecurity piLevel + 1
                Next
            Else
                .CurrentX = giPRINT_XINDENT + ((piLevel + 2) * giPRINT_XSPACE)
                .CurrentY = .CurrentY + giPRINT_YSPACE
                .FontSize = 6
                Printer.Print "***There are no current users***"
            End If
        End With
        
        mlngBottom = CalculateBottomOfPage
        Printer.CurrentY = mlngBottom + 1
        CheckEndOfPage2 mlngBottom, False
    End If
    
    '*********************************************************************************************
    'PRINT THE TABLEVIEWS
    If (gasPrintOptions(1).PrintLPaneTABLESVIEWS = True) Or (gasPrintOptions(1).PrintRPaneTABLESVIEWS = True) Then
          'Order the TABLE VIEWS
          ReDim asViewInfo(2, 0)
          sSQL = "SELECT ASRSysTables.tableName, ASRSysViews.viewName" & _
            " FROM ASRSysTables " & _
            " INNER JOIN ASRSysViews ON ASRSysViews.viewTableID = ASRSysTables.tableID"
          rsTables.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
          
          Do While Not rsTables.EOF
            ReDim Preserve asViewInfo(2, UBound(asViewInfo, 2) + 1)
            asViewInfo(1, UBound(asViewInfo, 2)) = rsTables!TableName
            asViewInfo(2, UBound(asViewInfo, 2)) = rsTables!ViewName
            
            rsTables.MoveNext
          Loop
          rsTables.Close
          
          ReDim avTablesViews(mvar_objTables.Count + mvar_objViews.Count, 3)
          
          For Each objSecurityTable In mvar_objTables
            sTableName = objSecurityTable.Name
            iIndex = 1
            
            For Each objSecurityTable2 In mvar_objTables
              If objSecurityTable2.Name < sTableName Then
                iIndex = iIndex + 1
              End If
            Next objSecurityTable2
            
            For Each objSecurityView2 In mvar_objViews
              For iLoop = 1 To UBound(asViewInfo, 2)
                If asViewInfo(2, iLoop) = objSecurityView2.Name Then
                  sFullViewName = asViewInfo(1, iLoop) & " - '" & objSecurityView2.Name & "' - VIEW"
                  Exit For
                End If
              Next iLoop
              
              If sFullViewName < sTableName Then
                iIndex = iIndex + 1
              End If
            Next objSecurityView2
            
            avTablesViews(iIndex, 1) = sTableName
            avTablesViews(iIndex, 2) = 1
            avTablesViews(iIndex, 3) = sTableName
          Next objSecurityTable
          
          For Each objSecurityView In mvar_objViews
            For iLoop = 1 To UBound(asViewInfo, 2)
              If asViewInfo(2, iLoop) = objSecurityView.Name Then
                sTableName = asViewInfo(1, iLoop) & " - '" & objSecurityView.Name & "' - VIEW"
                Exit For
              End If
            Next iLoop
            iIndex = 1
            
            For Each objSecurityTable2 In mvar_objTables
              If objSecurityTable2.Name < sTableName Then
                iIndex = iIndex + 1
              End If
            Next objSecurityTable2
            
            For Each objSecurityView2 In mvar_objViews
              For iLoop = 1 To UBound(asViewInfo, 2)
                If asViewInfo(2, iLoop) = objSecurityView2.Name Then
                  sFullViewName = asViewInfo(1, iLoop) & " - '" & objSecurityView2.Name & "' - VIEW"
                  Exit For
                End If
              Next iLoop
              
              If sFullViewName < sTableName Then
                iIndex = iIndex + 1
              End If
            Next objSecurityView2
            
            avTablesViews(iIndex, 1) = sTableName
            avTablesViews(iIndex, 2) = 0
            avTablesViews(iIndex, 3) = objSecurityView.Name
          Next objSecurityView
        
          '*************************************************************************************************************
          'PRINT THE TABLES AND VIEWS AFTER BEING READ FROM ALPHABETICALLY ORDERED ARRAYS.
          For iLoop = 1 To (mvar_objTables.Count + mvar_objViews.Count)
            If avTablesViews(iLoop, 2) = 1 Then
              ' Table
              mvar_objTables.Item(avTablesViews(iLoop, 3)).PrintHeadings = (iLoop = 1)
              mvar_objTables.Item(avTablesViews(iLoop, 3)).PrintSecurity piLevel + 2, (avTablesViews(iLoop, 1))
              'Exit For ' shouldn't be activated, testing purposes only
            Else
              ' view
              mvar_objViews.Item(avTablesViews(iLoop, 3)).PrintHeadings = (iLoop = 1)
              mvar_objViews.Item(avTablesViews(iLoop, 3)).PrintSecurity piLevel + 2, (avTablesViews(iLoop, 1))
             ' Exit For ' shouldn't be activated, testing purposes only
            End If
            mlngBottom = CalculateBottomOfPage
          Next iLoop
          
          mlngBottom = CalculateBottomOfPage
          Printer.CurrentY = mlngBottom + 1
          CheckEndOfPage2 mlngBottom, False
    End If
        
    '**************************************************************************************************************
    'PRINT THE SYSTEM PERMSISSIONS + headings
    If (gasPrintOptions(1).PrintLPaneSYSTEM = True) Or (gasPrintOptions(1).PrintRPaneSYSTEM = True) Then
      
      'Reorder system permissions into correct order
      ReDim Preserve asSysPerms(mvar_objSystemPermissions.Count)
      ReDim Preserve asTemp(mvar_objSystemPermissions.Count)
      
      For iCount = 1 To mvar_objSystemPermissions.Count
        asSysPerms(iCount).ItemID = UCase(mvar_objSystemPermissions.Item(iCount).ItemID)
        asSysPerms(iCount).Category = UCase(mvar_objSystemPermissions.Item(iCount).CategoryKey)
        asSysPerms(iCount).Item = UCase(mvar_objSystemPermissions.Item(iCount).ItemKey)
        asSysPerms(iCount).Allowed = UCase(mvar_objSystemPermissions.Item(iCount).Allowed)
      Next
      
      With Printer
        intOldFontSize = .FontSize
        .FontSize = 8
        .FontUnderline = True

        '.CurrentY = giPRINT_YINDENT
        .CurrentX = giPRINT_XINDENT
        Printer.Print "System Permissions" & vbCrLf
        intTopY = .CurrentY
        .FontUnderline = False
        mlngBottom = CalculateBottomOfPage - 400
        
        miColumnNumber = 1
  
        ' Loop through each of the categories
        sSQL = "SELECT DISTINCT CategoryKey, Description, ListOrder FROM ASRSysPermissionCategories" _
              & " ORDER BY ListOrder, Description"
        
        rsPermissionCategory.Open sSQL, gADOCon, adOpenForwardOnly, adLockReadOnly
        Do While Not rsPermissionCategory.EOF
               
          sKey = rsPermissionCategory!CategoryKey
          
          'NHRD20100428 JIRA HRPRO 884
          If Not (sKey = "CMG" And Not IsModuleEnabled(modCMG)) _
            And Not (sKey = "NINEBOXGRID" And Not IsModuleEnabled(modNineBoxGrid)) _
            And Not (sKey = "WORKFLOW" And Not IsModuleEnabled(modWorkflow)) _
            And Not (sKey = "VERSION1" And Not IsModuleEnabled(modVersionOne)) _
            And Not (sKey = "ACCORD" And Not IsModuleEnabled(modAccord)) Then
                      
            .CurrentX = giPRINT_XINDENT + (giPRINT_XCOLUMN * (miColumnNumber - 1))
  
            ' Icon
            Set objIcon = frmGroupMaint1.imgSystemPermissions.ListImages("IMG_" & sKey).Picture
            .PaintPicture objIcon, Printer.CurrentX, Printer.CurrentY - 50
  
            ' Description
            .FontSize = 8
            .CurrentX = .CurrentX + 270
            Printer.Print rsPermissionCategory!Description
            .CurrentY = .CurrentY + 30
  
            For iLoopItems = 1 To mvar_objSystemPermissions.Count
              If mvar_objSystemPermissions.Item(iLoopItems).CategoryKey = sKey Then
                mvar_objSystemPermissions.Item(iLoopItems).PrintSecurity miColumnNumber
              End If
            Next iLoopItems
           
            ' Check bottom of page
            .CurrentY = .CurrentY + 450
            If Printer.CurrentY > mlngBottom - 600 Then
              .CurrentY = intTopY
              miColumnNumber = miColumnNumber + 1
              
              ' Handle the System Permissions overflowing onto extra pages.
              If ((2 * giPRINT_XINDENT) + (giPRINT_XCOLUMN * miColumnNumber)) > Printer.ScaleWidth Then
                
                mlngBottom = CalculateBottomOfPage
                Printer.CurrentY = mlngBottom + 1
                ForceEndOfPage

                miColumnNumber = 1
                .CurrentY = intTopY
              End If
            End If
          End If
          
          rsPermissionCategory.MoveNext
          
        Loop
        rsPermissionCategory.Close
        
        ' Print page footer
        mlngBottom = CalculateBottomOfPage
        Printer.CurrentY = mlngBottom + 1
        CheckEndOfPage2 mlngBottom, False
        
      End With
    End If
    
TidyUpAndExit:
  Set rsTables = Nothing
  Set rsPermissionCategory = Nothing
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Function CheckToPrintHeader() As Boolean
  CheckToPrintHeader = False
  
  If gasPrintOptions(1).PrintLPaneTABLE And gasPrintOptions(1).PrintRPaneTABLE Then
    CheckToPrintHeader = True
    Exit Function
  End If
  
  If gasPrintOptions(1).PrintLPaneUSERS Or gasPrintOptions(1).PrintRPaneUSERS Then
    CheckToPrintHeader = True
    Exit Function
  End If
      
  If gasPrintOptions(1).PrintLPaneSYSTEM Or gasPrintOptions(1).PrintRPaneSYSTEM Then
    CheckToPrintHeader = True
    Exit Function
  End If
  
End Function

Public Property Let Users_Initialised(ByVal pfData As Boolean)
  mvar_fUsersInitialised = pfData
  
End Property

Public Property Let Initialised(ByVal pfData As Boolean)
    mvar_fInitialised = pfData
    
End Property

Public Property Get Initialised() As Boolean
  Initialised = mvar_fInitialised
  
End Property

Public Property Get Views() As SecurityTables
  ' Create the Views collection if it doesn't exist.
  If mvar_objViews Is Nothing Then
    Set mvar_objViews = New SecurityTables
  End If

  Set Views = mvar_objViews
  
End Property

Public Property Set Views(pobjData As SecurityTables)
  Set mvar_objViews = pobjData
  
End Property

Public Property Get SystemPermissions() As clsSystemPermissions
  ' Create the System permissions collection if it doesn't exist.
  If mvar_objSystemPermissions Is Nothing Then
    Set mvar_objSystemPermissions = New clsSystemPermissions
  End If

  Set SystemPermissions = mvar_objSystemPermissions
    
End Property

Public Property Set SystemPermissions(pobjData As clsSystemPermissions)
  Set mvar_objSystemPermissions = pobjData
  
End Property

Public Property Let DeleteGroup(ByVal pfData As Boolean)
  mvar_fDeleteGroup = pfData
  
End Property

Public Property Get DeleteGroup() As Boolean
  DeleteGroup = mvar_fDeleteGroup
  
End Property

Public Property Let NewGroup(ByVal pfData As Boolean)
  mvar_fNewGroup = pfData
  
End Property

Public Property Get NewGroup() As Boolean
  NewGroup = mvar_fNewGroup
  
End Property
Public Property Let Changed(ByVal pfData As Boolean)
  mvar_fChanged = pfData
  
End Property


Public Property Get Changed() As Boolean
  Changed = mvar_fChanged
    
End Property

Public Property Get Users() As SecurityUsers
  ' Create the Users collection if it doesn't exist.
  If mvar_objUsers Is Nothing Then
    Set mvar_objUsers = New SecurityUsers
  End If

  Set Users = mvar_objUsers
  
End Property

Public Property Set Users(pobjData As SecurityUsers)
  Set mvar_objUsers = pobjData
    
End Property

Public Property Get Tables() As SecurityTables
  ' Create the Tables collection if it doesn't exist.
  If mvar_objTables Is Nothing Then
    Set mvar_objTables = New SecurityTables
  End If

  Set Tables = mvar_objTables
    
End Property


Public Property Set Tables(pobjData As SecurityTables)
  Set mvar_objTables = pobjData
    
End Property

Public Property Let Name(ByVal psData As String)
  On Error GoTo NameLetErr

  mvar_sName = psData
  Exit Property

NameLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityGroup:Name Property Let")
  MsgBox "Error SecurityGroup:Name Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get Name() As String
  On Error GoTo NameGetErr

  Name = mvar_sName
  Exit Property

NameGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityGroup:Name Property Get")
  MsgBox "Error SecurityGroup:Name Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Let OriginalName(ByVal psData As String)
  On Error GoTo OriginalNameLetErr

  mvar_sOriginalName = psData
  Exit Property

OriginalNameLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityGroup:OriginalName Property Let")
  MsgBox "Error SecurityGroup:OriginalName Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let AccessCopyGroup(ByVal psData As String)

  mvar_sAccessCopyGroup = psData

End Property


Public Property Get OriginalName() As String
  On Error GoTo OriginalNameGetErr

  OriginalName = mvar_sOriginalName
  Exit Property

OriginalNameGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityGroup:OriginalName Property Get")
  MsgBox "Error SecurityGroup:OriginalName Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get AccessCopyGroup() As String
  
  AccessCopyGroup = mvar_sAccessCopyGroup
  
End Property


Private Sub Class_Initialize()
  #If DebugMode Then
    'get the next available class ID, and print out
    'that the class was created successfully
    mvar_lngClassDebugID = GetNextClassDebugID()
'    Debug.Print "'" & TypeName(Me) & "' instance " & mvar_lngClassDebugID & " created"
  #End If
  
End Sub

Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mvar_objUsers = Nothing
  Set mvar_objTables = Nothing
  Set mvar_objViews = Nothing
'  Set mvar_objSysPrivs = Nothing
Set mvar_objSystemPermissions = Nothing
  
  ' Flag that the class is being destroyed.
  #If DebugMode Then
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mvar_lngClassDebugID) & " is terminating"
  #End If
  
End Sub

#If DebugMode Then
  Public Property Get ClassDebugID()
    'if we are in debug mode, surface this property that consumers can query
    ClassDebugID = mvar_lngClassDebugID
  End Property
#End If

Public Property Get RequireLogout() As Boolean
  RequireLogout = mvar_fRequireLogout
End Property

Public Property Let RequireLogout(ByVal blnNewValue As Boolean)
  mvar_fRequireLogout = blnNewValue
End Property

Public Function Clone(ByVal pbPerfectCopy As Boolean, Optional pstrNewName As String) As SecurityGroup

  Dim objNewGroup As SecurityGroup

  Set objNewGroup = New SecurityGroup
  
  ' Copy the settings
  With objNewGroup
    .Changed = IIf(pbPerfectCopy, Changed, True)
    .DeleteGroup = IIf(pbPerfectCopy, DeleteGroup, False)
    .Initialised = Initialised
    .Name = IIf(pbPerfectCopy, Name, IIf(Len(pstrNewName) > 0, pstrNewName, "Copy_of_" & Name))
    .NewGroup = IIf(pbPerfectCopy, NewGroup, True)
    .RequireLogout = IIf(pbPerfectCopy, RequireLogout, False)
    Set .SystemPermissions = SystemPermissions.Clone
    Set .Tables = Tables.Clone
    Set .Users = Users.Clone(pbPerfectCopy)
    .Users_Initialised = Users_Initialised
    Set .Views = Views.Clone
  End With
  
  ' Return the cloned security group
  Set Clone = objNewGroup

End Function

Public Function Replace(ByVal pobjGroup As SecurityGroup) As Boolean

  ' Copy the settings
  With Me
    .Changed = pobjGroup.Changed
    .DeleteGroup = pobjGroup.DeleteGroup
    .Initialised = pobjGroup.Initialised
    .Name = pobjGroup.Name
    .NewGroup = pobjGroup.NewGroup
    .RequireLogout = pobjGroup.RequireLogout
    Set .SystemPermissions = pobjGroup.SystemPermissions
    Set .Tables = pobjGroup.Tables
    Set .Users = pobjGroup.Users
    .Users_Initialised = pobjGroup.Users_Initialised
    Set .Views = pobjGroup.Views
  End With
  
  Replace = True

End Function

Public Property Get AccessConfiguration() As Variant
  AccessConfiguration = mvar_avAccessConfiguration
End Property

Public Property Let AccessConfiguration(ByVal vNewValue As Variant)
  mvar_avAccessConfiguration = vNewValue

End Property

Public Property Get CopyGroup() As String
  CopyGroup = mvar_sCopyGroup

End Property

Public Property Let CopyGroup(ByVal psNewValue As String)
  mvar_sCopyGroup = psNewValue

End Property
