VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' Screen Types
Enum ScreenType
  screenParent = 1
  screenHistory = 2
  screenLookup = 3
  screenFind = 4
End Enum

'Edit option constants
Public Enum EditOptions
  edtCancel = 0
  edtAdd = 2 ^ 10
  edtDelete = 2 ^ 11
  edtEdit = 2 ^ 12
  edtCopy = 2 ^ 13
  edtSelect = 2 ^ 14
  edtDeselect = 2 ^ 15
  edtPrint = 2 ^ 16
  edtProperties = 2 ^ 17
End Enum

'Table type constants
Public Enum TableTypes
  tabParent = 1
  tabChild = 2
  tabLookup = 3
End Enum

'Control type constants
Public Enum ControlTypes
  ctlCheck = 1
  ctlCombo = 2
  ctlImage = 4
  ctlOLE = 8
  ctlRadio = 16
  ctlSpin = 32
  ctlText = 64
  ctlTab = 128
  ctlLabel = 256
  ctlFrame = 512
End Enum
  
'Column type constants
Public Enum ColumnTypes
  colData = 0
  colLookup = 1
  colCalc = 2
  colSystem = 3
  colLink = 4
  colWorkPattern = 5
End Enum

'CrossTabTypes for Crosstabs and Nine Box Grids
Public Enum CrossTabType
        cttNormal = 0
        cttTurnover = 1
        cttStability = 2
        cttAbsenceBreakdown = 3
        ctt9GridBox = 4
End Enum

Public Enum AccessModes
  accNone = 0
  accFull = 1
  accSupportMode = 2   'SYS MGR ONLY
  accLimited = 3       'SYS MGR ONLY
  accSystemReadOnly = 4
End Enum

'Instantiate internal classes
Private ODBC As New ODBC

'Local variables to hold property values
Private mvarLoggedIn As Boolean
Private gfChanged As Boolean

' Column SELECT and UPDATE privilege constants
Public Enum ColumnPrivilegeStates
  giPRIVILEGES_NONEGRANTED = 0
  giPRIVILEGES_ALLGRANTED = 1
  giPRIVILEGES_SOMEGRANTED = 2
End Enum

Private mintAccessMode As AccessModes

Public Enum LockTypes
  lckNone = 0
  lckSaving = 1
  lckManual = 2
  lckReadWrite = 3
  lckSaveRequest = 4
End Enum


Public Property Get Changed() As Boolean
  'MH2001 20010531 Put in an extra check to only allow save changes if full access
  'Changed = gfChanged
  Changed = ((gfChanged Or gbShiftSave) And mintAccessMode = accFull)
End Property

Public Property Let Changed(ByVal IsChanged As Boolean)
  gfChanged = IsChanged
  
End Property
Public Property Get LoggedIn() As Boolean
  LoggedIn = mvarLoggedIn
End Property

Public Property Let LoggedIn(ByVal vData As Boolean)
  mvarLoggedIn = vData
End Property

Public Property Get Name() As String
  Name = App.ProductName
End Property

Public Function Activate() As Boolean
  On Error GoTo ErrorTrap
  
  Dim fOK As Boolean
  
  If Not LoggedIn Then
    fOK = Login
  End If
  
  ' NPG20090924 Fault HRPro-328
  Dim frmStyle As New frmHiddenStyle
  Load frmStyle
  
  If IsFormLoaded("frmHiddenStyle") Then
    Unload frmStyle
  End If
  
  If LoggedIn Then
    ' Initialise the collection of user groups (roles).
    fOK = InitialiseGroupsCollection(gObjGroups)

    If fOK Then
      fOK = LoadDomainSecurityPolicy
    End If

    If fOK Then
      ' Load the main form.
      frmMain.SetBackground (True)
      frmMain.Show
    End If
    
    'fOK = True
  End If
  
ErrorHandler:
  Activate = fOK
  Exit Function

ErrorTrap:
  fOK = False
  Err = False
  
End Function



Public Function IsFormLoaded(ByVal FormName As String) As Boolean
    Dim frm As Form


    For Each frm In Forms
        If LCase(frm.Name) = LCase(FormName) Then IsFormLoaded = True
    Next frm
End Function



Public Function Login() As Boolean
  On Error GoTo ErrorTrap
  
  If LoggedIn Then
    Logout
  End If
  
  If Not LoggedIn Then
  
    Load frmLogin
    LoggedIn = frmLogin.OK
    
    If Not LoggedIn Then
      frmLogin.Show vbModal
      LoggedIn = frmLogin.OK

      If Not LoggedIn Then
        Logout
      End If

    End If

    Unload frmLogin
    Set frmLogin = Nothing
  End If
  
  Login = LoggedIn
  If Login Then Call AuditAccess("Log In", "Security")
  Exit Function
  
ErrorTrap:
  Login = False
  Err = False
  
End Function

Public Function Logout() As Boolean
  On Error GoTo ErrorTrap

  If Not gADOCon Is Nothing Then
    If Application.AccessMode = accFull Then
      UnlockDatabase lckReadWrite
    End If
  End If

  If LoggedIn Then
  
    If Not gObjGroups Is Nothing Then
      Set gObjGroups = Nothing
    End If
    
    gADOCon.Close
    Set gADOCon = Nothing   'MH20060213 Fault 10786
    
    LoggedIn = False
  End If
  
  Logout = (Not LoggedIn)

  Exit Function
  
ErrorTrap:
  Logout = False
  Err = False
  
End Function

Private Sub Class_Initialize()
  mintAccessMode = accFull
End Sub

Private Sub Class_Terminate()
  'Release internal classes
  Set ODBC = Nothing
End Sub

Public Property Get AccessMode() As AccessModes
  AccessMode = mintAccessMode
End Property

Public Property Let AccessMode(ByVal intNewValue As AccessModes)
  mintAccessMode = intNewValue
End Property

