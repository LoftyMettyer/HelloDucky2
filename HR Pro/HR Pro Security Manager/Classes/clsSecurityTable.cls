VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SecurityTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SecurityColumns"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

' Set this to 0 to disable debug code in this class.
#Const DebugMode = 0

#If DebugMode Then
  ' Local variable to hold the serialized class ID that was created in Class_Initialize.
  Private mvar_lngClassDebugID As Long
#End If

' Local variables to hold property values.
Private mvar_lngTableID As Long
Private mvar_lngViewTableID As Long
Private mvar_sName As String
Private mvar_fInsertPrivilege As Boolean
Private mvar_fInsertOriginalPrivilege As Boolean
Private mvar_iSelectPrivilege As ColumnPrivilegeStates
Private mvar_iSelectOriginalPrivilege As ColumnPrivilegeStates
Private mvar_iUpdatePrivilege As ColumnPrivilegeStates
Private mvar_iUpdateOriginalPrivilege As ColumnPrivilegeStates
Private mvar_fDeletePrivilege As Boolean
Private mvar_fDeleteOriginalPrivilege As Boolean
Private mvar_fHideFromMenu As Boolean
Private mvar_fChanged As Boolean
Private mvar_objColumns As SecurityColumns
Private mvar_fColumnsInitialised As Boolean
Private mvar_iTableType As Integer
Private mvar_iParentJoinType As Integer
Private mvar_iParentCount As Integer
Private mvar_iAccessibleParentCount As Integer
Private mbPrintHeadings As Boolean

Private mlngBottom As Long



Public Property Get ParentJoinType() As Integer
  ParentJoinType = mvar_iParentJoinType

End Property
Public Function PrintSecurity(piLevel As Integer, psDescription As String) As Boolean
  ' Print the
  On Error GoTo ErrorTrap

  Dim fOK As Boolean
  Dim fPrintedHeadings As Boolean
  Dim intOldFontSize As Integer
  Dim objColumn As SecurityColumn
  Dim intCurrentX As Integer
  Dim iCurrentPrintY As Integer
  Dim strPrinterString As String
  Dim objIcon As IPictureDisp
  Dim intSameLine As Integer
  Dim strDisplayName As String
  
  fOK = True

  If (gasPrintOptions(1).PrintLPaneTABLESVIEWS = True) Or (gasPrintOptions(1).PrintRPaneTABLESVIEWS = True) Then
    'Position the printing.
    With Printer
        intOldFontSize = .FontSize

        'check where your bottom is
        mlngBottom = CalculateBottomOfPage()
        
        If .CurrentY >= mlngBottom - 400 Then
          .CurrentY = mlngBottom + 1
          CheckEndOfPage2 mlngBottom
          mbPrintHeadings = True
        End If
        
        If mbPrintHeadings Then
            ' Print Headings
            .FontSize = 8: .FontBold = True
            .CurrentX = giPRINT_XINDENT + (0 * giPRINT_XSPACE)
            .CurrentY = giPRINT_YINDENT
            
            If gasPrintOptions(1).PrintBlankVersion Then
                Printer.Print "Blank Version For Admin Purposes - GROUPNAME:  " & gstrPrintGroupName
            Else
                Printer.Print "GROUPNAME:  " + gstrPrintGroupName
            End If
            .FontBold = False
            
            intOldFontSize = .FontSize
            .FontSize = 8
            .FontUnderline = True
            '.CurrentX = giPRINT_XINDENT + ((piLevel + 1) * giPRINT_XSPACE)
            .CurrentX = ((piLevel + 1) * giPRINT_XSPACE)
            .CurrentY = .CurrentY + giPRINT_YSPACE
            Printer.Print "Table and View Details"
            .FontUnderline = False
            
            .FontSize = 6: .FontUnderline = True
  
            .CurrentY = .CurrentY + (giPRINT_YSPACE)
            intSameLine = .CurrentY + (giPRINT_YSPACE)
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 3) * giPRINT_XINDENT)
            .CurrentY = intSameLine
            Printer.Print "New Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 4) * giPRINT_XINDENT)
            .CurrentY = intSameLine
            Printer.Print "Edit Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 5) * giPRINT_XINDENT)
            .CurrentY = intSameLine
            Printer.Print "Read Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 6) * giPRINT_XINDENT)
            .CurrentY = intSameLine
            Printer.Print "Delete Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 7) * giPRINT_XINDENT)
            .CurrentY = intSameLine
            Printer.Print "Hide from Menu"
  
            mbPrintHeadings = False
            
            'This variable stops the column headings being printed twice.
            fPrintedHeadings = True
            .FontUnderline = False
        End If
        
        'NHRD30072004 Fault 8515
        If (gasPrintOptions(1).PrintLPaneTABLE Or gasPrintOptions(1).PrintRPaneTABLE) And fPrintedHeadings = False Then
            .FontSize = 6: .FontUnderline = True
            .CurrentY = .CurrentY + (giPRINT_YSPACE):         intSameLine = .CurrentY + (giPRINT_YSPACE)
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 3) * giPRINT_XINDENT):    .CurrentY = intSameLine
            Printer.Print "New Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 4) * giPRINT_XINDENT):    .CurrentY = intSameLine
            Printer.Print "Edit Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 5) * giPRINT_XINDENT):    .CurrentY = intSameLine
            Printer.Print "Read Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 6) * giPRINT_XINDENT):    .CurrentY = intSameLine
            Printer.Print "Delete Table"
  
            .CurrentX = giPRINT_XINDENT + ((piLevel + 7) * giPRINT_XINDENT):    .CurrentY = intSameLine
            Printer.Print "Hide from Menu"
        End If
        
        fPrintedHeadings = False
        
        ' Current line coordinates
        .FontSize = 8
        .FontUnderline = False
        
        If gasPrintOptions(1).PrintLPaneTABLE Then
          iCurrentPrintY = .CurrentY + (giPRINT_YSPACE + 1)
        Else
          iCurrentPrintY = .CurrentY + (giPRINT_YSPACE + 180)
        End If
        ' Table Type Icon
        Select Case mvar_iTableType
          Case tabParent
            Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("TOPLEVELTABLE").Picture
            strDisplayName = IIf(Len(mvar_sName) > 45, Left(mvar_sName, 42) & "...", mvar_sName)
          Case tabChild
            Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("CHILDTABLE").Picture
            strDisplayName = IIf(Len(mvar_sName) > 45, Left(mvar_sName, 42) & "...", mvar_sName)
          Case tabLookup
            Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("LOOKUPTABLE").Picture
            strDisplayName = IIf(Len(mvar_sName) > 45, Left(mvar_sName, 42) & "...", mvar_sName)
          Case Else
            Set objIcon = frmGroupMaint1.imlSmallIcons.ListImages.Item("VIEW").Picture
            strDisplayName = Replace(psDescription, " - VIEW", "", 1, -1, vbBinaryCompare)
            strDisplayName = IIf(Len(strDisplayName) > 45, Left(strDisplayName, 42) & "...", strDisplayName)
        End Select
        
        .CurrentX = giPRINT_XINDENT + ((piLevel - 1) * giPRINT_XSPACE)
        .CurrentY = iCurrentPrintY - 90
        .PaintPicture objIcon, Printer.CurrentX, Printer.CurrentY
        
        ' Description
        .CurrentX = .CurrentX + objIcon.Width + 5
        .CurrentY = iCurrentPrintY - 5
        
        ' Shorten name if necessary
        Printer.Print strDisplayName
                      
        'New Table Property
        .CurrentY = iCurrentPrintY - 30
        intCurrentX = giPRINT_XINDENT + ((piLevel + 3) * giPRINT_XINDENT) + 200
        Call PrintPictureBox(intCurrentX, iCurrentPrintY - 90, mvar_fInsertPrivilege)
        
        'Edit Table Property
        intCurrentX = giPRINT_XINDENT + ((piLevel + 4) * giPRINT_XINDENT) + 250
        Call PrintPictureBox(intCurrentX, iCurrentPrintY - 90, mvar_iUpdatePrivilege)

        'Read Table Property
        intCurrentX = giPRINT_XINDENT + ((piLevel + 5) * giPRINT_XINDENT) + 200
        Call PrintPictureBox(intCurrentX, iCurrentPrintY - 90, mvar_iSelectPrivilege)

        'Delete Table Property
        intCurrentX = giPRINT_XINDENT + ((piLevel + 6) * giPRINT_XINDENT) + 250
        Call PrintPictureBox(intCurrentX, iCurrentPrintY - 90, mvar_fDeletePrivilege)
        
        If mvar_iTableType = tabLookup Then
          'Hide from menu Property (Lookuip Tables Only)
          intCurrentX = giPRINT_XINDENT + ((piLevel + 7) * giPRINT_XINDENT) + 250
          Call PrintPictureBox(intCurrentX, iCurrentPrintY - 90, mvar_fHideFromMenu)
        End If
                
        If mvar_iParentCount > 1 Then
            .FontSize = 6
                'Users have access to child records related to ...
                strPrinterString = "Users have access to child records for..."
                
                Select Case ParentJoinType
                    Case 0
                        strPrinterString = strPrinterString + "'Any Permitted Parent'"
                    Case 1
                        strPrinterString = strPrinterString + "'All Permitted Parents'"
                    Case Else
                        strPrinterString = "Null"
                End Select
                
                .CurrentX = giPRINT_XINDENT + ((piLevel + 1) * giPRINT_XSPACE)
                .CurrentY = .CurrentY + giPRINT_YSPACE + 80
                Printer.Print strPrinterString
        End If
        
        'Restore font size
        .FontSize = intOldFontSize
        
    End With
    ' Print out these details if desirable
    If (gasPrintOptions(1).PrintLPaneTABLE = True) Or (gasPrintOptions(1).PrintRPaneTABLE = True) Then
        'Only print them if some of the details have been changed.
        'As it is they will be printed everytime you run this bit of code.
        'NHRD17092003 Fault 6396 If you reactivate this line the Security Permissions will not print column privaliges unless they have changed.
      'If (SelectPrivilege = giPRIVILEGES_SOMEGRANTED) Or (UpdatePrivilege = giPRIVILEGES_SOMEGRANTED) Or (gasPrintOptions(1).PrintBlankVersion = True) Then
                
        If mvar_objColumns.Count > 0 Then
          'If it is the last line there is no point in printing the headings
          'as they are separated from the data coloumns - so check for this.
          CheckEndOfPage2 (mlngBottom - (giPRINT_YSPACE * 3))
          
          PrintColumnHeadings (piLevel)
          CheckEndOfPage2 (mlngBottom)
            For Each objColumn In mvar_objColumns
              If Not objColumn.Changed Then objColumn.PrintSecurity piLevel
              'Exit For ' shouldn't be activated testing purposes only

              If Printer.CurrentX = 1000 And Printer.CurrentY = 1000 Then 'NewPage
                If mvar_objColumns(mvar_objColumns.Count).Name <> objColumn.Name Then 'NHRD15112004 Fault 8515
                  PrintColumnHeadings (piLevel)
                End If
              End If
            Next objColumn
        End If
      'End If  'NHRD17092003 Fault 6396 See above
    End If
  End If
TidyUpAndExit:
  PrintSecurity = fOK
  Exit Function
  
ErrorTrap:
  fOK = False
  Resume TidyUpAndExit
  
End Function

Public Sub PrintColumnHeadings(piLevel As Integer)
Dim iCurrentPrintY As Integer
Dim intOldFontSize As Integer

  With Printer
    intOldFontSize = .FontSize
    .FontSize = 6
    .FontUnderline = True

    iCurrentPrintY = .CurrentY + (giPRINT_YSPACE * 3)
    'CheckEndOfPage2 mlngBottom

    .CurrentX = giPRINT_XINDENT + ((piLevel + 5) * giPRINT_XINDENT)
    .CurrentY = iCurrentPrintY
    Printer.Print "Read Column"

    .CurrentX = giPRINT_XINDENT + ((piLevel + 6) * giPRINT_XINDENT)
    .CurrentY = iCurrentPrintY
    Printer.Print "Edit Column"

    .FontUnderline = False
    .FontSize = intOldFontSize
    .FontUnderline = False
  End With
End Sub

Public Property Get ParentCount() As Integer
  ParentCount = mvar_iParentCount

End Property


Public Property Let ParentJoinType(ByVal piNewValue As Integer)
  mvar_iParentJoinType = piNewValue

End Property

Public Property Let ParentCount(ByVal piNewValue As Integer)
  mvar_iParentCount = piNewValue

End Property


Public Sub GrantAll()
  ' Grant all permissions on all columns in this table/view.
  Dim iLoop As Integer
  Dim objColumn As SecurityColumns
  
  For iLoop = 1 To Columns.Count
    ' Grant permission to the column.
    Set objColumn = Columns
    With objColumn.Item(iLoop)
      .Changed = (Not .SelectPrivilege) Or (Not .UpdatePrivilege)
      .SelectPrivilege = True
      .UpdatePrivilege = True
      
      If Not mvar_fChanged Then
        mvar_fChanged = .Changed
      End If
    End With
    Set objColumn = Nothing
  Next iLoop
  
End Sub
Public Property Let Columns_Initialised(ByVal pfData As Boolean)
    mvar_fColumnsInitialised = pfData
    
End Property


Public Property Get Columns_Initialised() As Boolean
  Columns_Initialised = mvar_fColumnsInitialised
  
End Property




Public Property Let Changed(ByVal pfData As Boolean)
  On Error GoTo ChangedLetErr

  mvar_fChanged = pfData
  Exit Property

ChangedLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Changed Property Let")
  MsgBox "Error SecurityTable:Changed Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get Changed() As Boolean
  On Error GoTo ChangedGetErr

  Changed = mvar_fChanged
  Exit Property

ChangedGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Changed Property Get")
  MsgBox "Error SecurityTable:Changed Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let DeletePrivilege(ByVal pfData As Boolean)
  On Error GoTo DeletePrivilegeLetErr

  mvar_fDeletePrivilege = pfData
  Exit Property

DeletePrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:DeletePrivilege Property Let")
  MsgBox "Error SecurityTable:DeletePrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get DeletePrivilege() As Boolean
  On Error GoTo DeletePrivilegeGetErr

  DeletePrivilege = mvar_fDeletePrivilege
  Exit Property

DeletePrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:DeletePrivilege Property Get")
  MsgBox "Error SecurityTable:DeletePrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let UpdatePrivilege(ByVal piData As ColumnPrivilegeStates)
  On Error GoTo UpdatePrivilegeLetErr

  mvar_iUpdatePrivilege = piData
  Exit Property

UpdatePrivilegeLetErr:
    'Call RaiseError(MyUnhandledError, "SecurityTable:UpdatePrivilege Property Let")
    MsgBox "Error SecurityTable:UpdatePrivilege Property Let.", _
      vbExclamation + vbOKOnly, App.ProductName
    
End Property

Public Property Get UpdatePrivilege() As ColumnPrivilegeStates
  On Error GoTo UpdatePrivilegeGetErr

  Dim objColumn As SecurityColumn
  Dim fAllGranted As Boolean
  Dim fNoneGranted As Boolean
  
  fAllGranted = True
  fNoneGranted = True
  
  ' Check the select privilege of the table/view's columns.
  For Each objColumn In Columns
    If objColumn.UpdatePrivilege Then
      fNoneGranted = False
    Else
      fAllGranted = False
    End If
  Next
  
  If fAllGranted Then
    mvar_iUpdatePrivilege = giPRIVILEGES_ALLGRANTED
  ElseIf fNoneGranted Then
    mvar_iUpdatePrivilege = giPRIVILEGES_NONEGRANTED
  Else
    mvar_iUpdatePrivilege = giPRIVILEGES_SOMEGRANTED
  End If
  
  UpdatePrivilege = mvar_iUpdatePrivilege
  Exit Property

UpdatePrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:UpdatePrivilege Property Get")
  MsgBox "Error SecurityTable:UpdatePrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Let SelectPrivilege(ByVal piData As ColumnPrivilegeStates)
  On Error GoTo SelectPrivilegeLetErr

  mvar_iSelectPrivilege = piData
  Exit Property

SelectPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:SelectPrivilege Property Let")
  MsgBox "Error SecurityTable:SelectPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get SelectPrivilege() As ColumnPrivilegeStates
  On Error GoTo SelectPrivilegeGetErr
  
  Dim objColumn As SecurityColumn
  Dim fAllGranted As Boolean
  Dim fNoneGranted As Boolean
  
  fAllGranted = True
  fNoneGranted = True
  
  ' Check the select privilege of the table/view's columns.
  For Each objColumn In Columns
    If objColumn.SelectPrivilege Then
      fNoneGranted = False
    Else
      fAllGranted = False
    End If
  Next
  
  If fAllGranted Then
    mvar_iSelectPrivilege = giPRIVILEGES_ALLGRANTED
  ElseIf fNoneGranted Then
    mvar_iSelectPrivilege = giPRIVILEGES_NONEGRANTED
  Else
    mvar_iSelectPrivilege = giPRIVILEGES_SOMEGRANTED
  End If

  SelectPrivilege = mvar_iSelectPrivilege
  Exit Property

SelectPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:SelectPrivilege Property Get")
  MsgBox "Error SecurityTable:SelectPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Let InsertPrivilege(ByVal pfData As Boolean)
  On Error GoTo InsertPrivilegeLetErr
  
  mvar_fInsertPrivilege = pfData
  Exit Property

InsertPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:InsertPrivilege Property Let")
  MsgBox "Error SecurityTable:InsertPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get InsertPrivilege() As Boolean
  On Error GoTo InsertPrivilegeGetErr
  
  InsertPrivilege = mvar_fInsertPrivilege
  Exit Property

InsertPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:InsertPrivilege Property Get")
  MsgBox "Error SecurityTable:InsertPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Function SelectPrivilegeChanged() As Boolean
  On Error GoTo SelectOriginalPrivilegeGetErr
  
  Dim objColumn As SecurityColumn
  Dim fNoneGranted As Boolean
  
  fNoneGranted = True
  
  ' Check the select OriginalPrivilege of the table/view's columns.
  For Each objColumn In Columns
    If objColumn.SelectOriginalPrivilege Then
      fNoneGranted = False
      Exit For
    End If
  Next
  
  If fNoneGranted And SelectPrivilege <> giPRIVILEGES_NONEGRANTED Then
    SelectPrivilegeChanged = True
  ElseIf Not fNoneGranted And SelectPrivilege = giPRIVILEGES_NONEGRANTED Then
    SelectPrivilegeChanged = True
  Else
    SelectPrivilegeChanged = False
  End If
  
  Exit Function

SelectOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:SelectOriginalPrivilege Property Get")
  MsgBox "Error SecurityTable:SelectOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Function

Public Property Let InsertOriginalPrivilege(ByVal pfData As Boolean)
  On Error GoTo InsertOriginalPrivilegeLetErr
  
  mvar_fInsertOriginalPrivilege = pfData
  Exit Property

InsertOriginalPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:InsertOriginalPrivilege Property Let")
  MsgBox "Error SecurityTable:InsertOriginalPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get InsertOriginalPrivilege() As Boolean
  On Error GoTo InsertOriginalPrivilegeGetErr
  
  InsertOriginalPrivilege = mvar_fInsertOriginalPrivilege
  Exit Property

InsertOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:InsertOriginalPrivilege Property Get")
  MsgBox "Error SecurityTable:InsertOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Function UpdatePrivilegeChanged() As Boolean
  On Error GoTo UpdateOriginalPrivilegeGetErr
  
  Dim objColumn As SecurityColumn
  Dim fNoneGranted As Boolean
  
  fNoneGranted = True
  
  ' Check the Update OriginalPrivilege of the table/view's columns.
  For Each objColumn In Columns
    If objColumn.UpdateOriginalPrivilege Then
      fNoneGranted = False
      Exit For
    End If
  Next
  
  If fNoneGranted And UpdatePrivilege <> giPRIVILEGES_NONEGRANTED Then
    UpdatePrivilegeChanged = True
  ElseIf Not fNoneGranted And UpdatePrivilege = giPRIVILEGES_NONEGRANTED Then
    UpdatePrivilegeChanged = True
  Else
    UpdatePrivilegeChanged = False
  End If
  
  Exit Function

UpdateOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:UpdateOriginalPrivilege Property Get")
  MsgBox "Error SecurityTable:UpdateOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Function

Public Property Let DeleteOriginalPrivilege(ByVal pfData As Boolean)
  On Error GoTo DeleteOriginalPrivilegeLetErr

  mvar_fDeleteOriginalPrivilege = pfData
  Exit Property

DeleteOriginalPrivilegeLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:DeleteOriginalPrivilege Property Let")
  MsgBox "Error SecurityTable:DeleteOriginalPrivilege Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get DeleteOriginalPrivilege() As Boolean
  On Error GoTo DeleteOriginalPrivilegeGetErr

  DeleteOriginalPrivilege = mvar_fDeleteOriginalPrivilege
  Exit Property

DeleteOriginalPrivilegeGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:DeleteOriginalPrivilege Property Get")
  MsgBox "Error SecurityTable:DeleteOriginalPrivilege Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let Name(ByVal psData As String)
  On Error GoTo NameLetErr

  mvar_sName = psData
  Exit Property

NameLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Name Property Let")
  MsgBox "Error SecurityTable:Name Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get Name() As String
  On Error GoTo NameGetErr

  Name = mvar_sName
  Exit Property

NameGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Name Property Get")
  MsgBox "Error SecurityTable:Name Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Get Columns() As SecurityColumns
  On Error GoTo ColumnsGetErr
    
  If mvar_objColumns Is Nothing Then
    Set mvar_objColumns = New SecurityColumns
  End If

  Set Columns = mvar_objColumns
  Exit Property

ColumnsGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Columns Property Get")
  MsgBox "Error SecurityTable:Columns Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName
  
End Property

Public Property Set Columns(pobjData As SecurityColumns)
  On Error GoTo ColumnsSetErr
    
  Set mvar_objColumns = pobjData
  Exit Property

ColumnsSetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:Columns Property Set")
  MsgBox "Error SecurityTable:Columns Property Set.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Private Sub Class_Initialize()
  #If DebugMode Then
    ' Get the next available class ID, and print out
    ' that the class was created successfully.
    mvar_lngClassDebugID = GetNextClassDebugID()
'    Debug.Print "'" & TypeName(Me) & "' instance " & mvar_lngClassDebugID & " created"
  #End If
  
  ' Initialise the permissions to revoked.
  mvar_fInsertPrivilege = False
  mvar_iSelectPrivilege = giPRIVILEGES_NONEGRANTED
  mvar_iUpdatePrivilege = giPRIVILEGES_NONEGRANTED
  mvar_fDeletePrivilege = False
  mvar_fChanged = False
  
End Sub

Private Sub Class_Terminate()
  ' Disassociate object variables.
  Set mvar_objColumns = Nothing
  
  #If DebugMode Then
'    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mvar_lngClassDebugID) & " is terminating"
  #End If
  
End Sub


#If DebugMode Then
  Public Property Get ClassDebugID()
    ' If we are in debug mode, surface this property that consumers can query.
    ClassDebugID = mvar_lngClassDebugID
  End Property
#End If


Public Property Get TableType() As Integer
  On Error GoTo ErrorTrap

  TableType = mvar_iTableType
  Exit Property

ErrorTrap:
  'Call RaiseError(MyUnhandledError, "SecurityTable:TableType Property Get")
  MsgBox "Error SecurityTable:TableType Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Let TableType(ByVal piNewValue As Integer)
  On Error GoTo ErrorTrap

  mvar_iTableType = piNewValue
  Exit Property

ErrorTrap:
  'Call RaiseError(MyUnhandledError, "SecurityTable:TableType Property Let")
  MsgBox "Error SecurityTable:TableType Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get AccessibleParentCount() As Integer
  AccessibleParentCount = mvar_iAccessibleParentCount
End Property

Public Property Let AccessibleParentCount(ByVal piNewValue As Integer)
  mvar_iAccessibleParentCount = piNewValue
End Property

Public Property Let HideFromMenu(ByVal pfData As Boolean)
  On Error GoTo HideFromMenuLetErr

  mvar_fHideFromMenu = pfData
  Exit Property

HideFromMenuLetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:HideFromMenu Property Let")
  MsgBox "Error SecurityTable:HideFromMenu Property Let.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get HideFromMenu() As Boolean
  On Error GoTo HideFromMenuGetErr

  HideFromMenu = mvar_fHideFromMenu
  Exit Property

HideFromMenuGetErr:
  'Call RaiseError(MyUnhandledError, "SecurityTable:HideFromMenu Property Get")
  MsgBox "Error SecurityTable:HideFromMenu Property Get.", _
    vbExclamation + vbOKOnly, App.ProductName

End Property

Public Property Get TableID() As Long
  TableID = mvar_lngTableID
End Property

Public Property Get ViewTableID() As Long
  ViewTableID = mvar_lngViewTableID
End Property

Public Property Let TableID(ByVal plngNewValue As Long)
  mvar_lngTableID = plngNewValue
End Property

Public Property Let ViewTableID(ByVal plngNewValue As Long)
  mvar_lngViewTableID = plngNewValue
End Property

Public Property Let PrintHeadings(ByVal pbNewValue As Boolean)
  mbPrintHeadings = pbNewValue
End Property
