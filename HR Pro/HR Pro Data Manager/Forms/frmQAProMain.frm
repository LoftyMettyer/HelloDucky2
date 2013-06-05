VERSION 5.00
Begin VB.Form frmQAProMain 
   Caption         =   "QuickAddress Pro"
   ClientHeight    =   5745
   ClientLeft      =   2355
   ClientTop       =   2640
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQAProMain.frx":0000
   LinkTopic       =   "frmMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   8070
   Begin VB.TextBox txtMessage 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   360
      TabIndex        =   13
      Top             =   5280
      Width           =   7335
   End
   Begin VB.ListBox lstResults 
      Height          =   2790
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   7335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select"
      Height          =   315
      Left            =   6480
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   6015
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame fraOptions 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdBack 
         Caption         =   "&Back"
         Height          =   315
         Left            =   6360
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbDatabase 
         Height          =   315
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   3165
      End
      Begin VB.OptionButton optTypedown 
         Caption         =   "&Typedown"
         Height          =   255
         Left            =   2670
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin VB.OptionButton optSingleLine 
         Caption         =   "&Single line"
         Height          =   255
         Left            =   1290
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   7320
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblDatabase 
         Alignment       =   1  'Right Justify
         Caption         =   "Database :"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label lblMode 
         Alignment       =   1  'Right Justify
         Caption         =   "Mode :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label lblHistory 
      Height          =   225
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label lblMatchCount 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblPrompt 
      Caption         =   "Enter place or postcode"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   7335
   End
End
Attribute VB_Name = "frmQAProMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const MODULE_NAME As String = "frmQAProMain"
Const DATASET_LENGTH As Long = 20
Const SEARCH_LENGTH As Long = 100
Const DEFAULT_DATASET As String = "GBR"
Const POSTCODE_DELIMITER = 9
Const ADDRESSFORM_FIELDS = 8
Const BUFFER_LENGTH As Long = 200
Const LB_SETTABSTOPS As Long = 402

' This is the handle to Pro API. A value of -1 indicates that the API is currently shutdown.
Public lngHandle As Long

' The currently selected search engine mode - qaengine_SINGLELINE or qaengine_TYPEDOWN
Public lngCurrentEngine As Long

' The currently selected data set
Public strCurrentData As String

' The picklist level. The first picklist stage is level 0. Each subsequent step-in
' results in the level being incremented. A value of -1 indicates that no results
' are available or the user has aborted the address capture process.
Public lngPicklistLevel As Long

' The selected (leaf-node) address. This will be set to -1 whilst the user is
' navigating the picklists. Once a leaf-node address has been selected this will
' be set to the offset of the picklist item.
Public lngAddress As Long

' Message to be displayed to the user. This is used to report postcode recordes,
' bordering localities etc.
Public strUserMessage As String

' Search history. This contains details of original search expression and the
' picklist items that have been stepped into. This information is displayed to
' the user to give them "context" for the current picklist.
Private arrHistory(10) As String
                                                    
'Module level var to hold the postcode we will be searching for
Private strPostCode As String
'Module level var to store whether or not the field mappings are individual or merged
Private fIndividual As Boolean
'Module level var to store which form called the Afd routine
Private frmForm As frmRecEdit4 ' Form

Private mbQuickAddressMode As Boolean
Private mobjQAPostcodes() As HRProDataMgr.PostCode
                                                    
Private Sub cmdBack_Click()

    Call HandleStepOut
    
    'txtInput.SetFocus

End Sub


Private Sub Form_Activate()

    ' Hide all unnecessary user options, then replicate the search button click.
  
    Set cmdBack.Container = frmQAProMain
    cmdBack.Top = 5280
    cmdBack.Left = 5160
    
    ' change the new button to be cancel
    cmdNew.Caption = "C&ancel"
    cmdNew.Top = 5280
    cmdNew.Left = 3840
    
    fraOptions.Visible = False
    lblPrompt.Visible = False
    txtInput.Visible = False
    ' cmdSelect.Visible = False
    txtMessage.Visible = False
    
    lstResults.Top = 585
    lstResults.Height = 4545
    
    lblHistory.Top = 240
    lblMatchCount.Top = 240
    
    txtMessage.Width = 2895
    cmdSelect.Top = 5280
    
    Dim blnReturn As Boolean
     
    blnReturn = HandleSelection(lstResults.ListIndex)
    
    If blnReturn Then
        FormatAddress ("")
    End If
    
    'txtInput.SetFocus

  ' Get rid of the icon off the form
  RemoveIcon Me

End Sub

Public Function InitialiseQA(PostCode As String, fIndiv As Boolean, frmCallingForm As Form, FieldName As String) As Boolean
    ' Quick Address or AFD mode
  mbQuickAddressMode = False

  'Store the postcode entered from the control on the users recedit form into a
  'module level variable
  strPostCode = Trim(PostCode)
  
  'Leave if the postcode is blank
  If strPostCode = "" Then
    InitialiseQA = False
    Exit Function
  End If
  
  'Store if the mapped fields are individual or not in a module level variable
  fIndividual = fIndiv
  
  'Set the calling form to a module level variable
  Set frmForm = frmCallingForm
  
  InitialiseQA = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyF1
      If ShowAirHelp(Me.HelpContextID) Then
        KeyCode = 0
      End If
  End Select
End Sub

'*******************************************************
' Purpose:  Called when the Form loads. Initialises the
'           Primary API and populates the Database list.
'*******************************************************

Private Sub Form_Load()

Dim lngOpenReturn As Long
Dim arrTabs(1 To 2) As Long


' Initialise state
lngHandle = 0
lngCurrentEngine = qaengine_TYPEDOWN
lngPicklistLevel = -1
lngAddress = -1
strUserMessage = ""

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("Form_Load", MODULE_NAME)
    
    Hook Me.hWnd, 4800, 2610
    
    lngOpenReturn = QA_Open("", "", lngHandle)

    ' If QA_Open returned no errors
    If lngOpenReturn = 0 Then
        
        Call BeginDialog
        
    ' Else QA_Open failed
    Else
    
        ' Raise Error
        Err.Raise lngOpenReturn
    
    End If
    
    'Set tab stops in lstResults
    arrTabs(1) = 235
    arrTabs(2) = 295
    Call SendMessageA(lstResults.hWnd, LB_SETTABSTOPS, 2, arrTabs(1))
    
' End Error Handler
Form_Load_Done:
    Call mProcStack.ExitProc("Form_Load")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "Form_Load"
    'Shutdown instance of the API
    Call ShutdownAPI
    Call mProcStack.ExitProc("Form_Load")
    End

End Sub

'*****************************************************
' Purpose:  Begin a new search. This function should be
'           called when starting a new search or when
'           switching data set or search engine mode.
' Inputs:
'   blnNewSearch:   True if performing new search
'                   False if switching data or engine
'*****************************************************

Public Sub BeginSearch(blnNewSearch As Boolean)

Dim lngError As Long
Dim strPrompt As String * SEARCH_LENGTH
Dim strSearchText As String * SEARCH_LENGTH

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("BeginSearch", MODULE_NAME)
    
    ' Blank the status line message
    DisplayMessage ("")
    
    ' Check if need to clear the search text
    If blnNewSearch Then
    
        SetSearchText ("")
    
    End If
    
    ' Discard any previous search results and reset search state
    lngError = QA_EndSearch(lngHandle)
    lngPicklistLevel = -1
    lngAddress = -1
    strUserMessage = ""
    
    ' Read current engine mode and data set as selected by the user
    Call GetEngine
    Call GetDataID
    
    ' Select engine and read current refinement prompt
    frmQAProMain.MousePointer = ccHourglass
    lngError = QA_SetActiveData(lngHandle, strCurrentData)
    frmQAProMain.MousePointer = ccDefault
    If lngError = 0 Then
        lngError = QA_SetEngine(lngHandle, lngCurrentEngine)
        If lngError = 0 Then
    
            lngError = QA_GetPrompt(lngHandle, 0, strPrompt, SEARCH_LENGTH, 0, "", 0)
            
        End If
    End If
   
    ' If Typedown search active then submit search to
    ' ensure informational prompts are displayed
    If lngError = 0 And lngCurrentEngine = qaengine_TYPEDOWN Then
    
        strSearchText = GetSearchText()
        frmQAProMain.MousePointer = ccHourglass
        lngError = QA_Search(lngHandle, strSearchText)
        frmQAProMain.MousePointer = ccDefault
        
        ' Set the search context to indicate that we are at
        ' the first stage of picklist handling (level 0)
        lngPicklistLevel = 0
        arrHistory(0) = ""
    
    End If
    
    ' If there were no errors then display prompt, initial picklist
    ' results and any data expiry warnings
    If lngError = 0 Then

        Call EnableCapture
        Call SetPrompt(strPrompt)

        frmQAProMain.MousePointer = ccHourglass
        Call UpdatePicklist
        frmQAProMain.MousePointer = ccDefault

    Else

        Call DisableCapture
        ' Raise Error
        Err.Raise lngError

    End If
    
' End Error Handler
BeginSearch_Done:
    Call mProcStack.ExitProc("BeginSearch")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "BeginSearch"
    Resume BeginSearch_Done

End Sub

Private Sub cmdNew_Click()

    If (QuitProgram() = False) Then
        'Cancel = 1
    End If

    Unload Me

   'Call BeginSearch(True)
   
   'txtInput.SetFocus

End Sub


Private Sub Form_Resize()
  ' Minimum width of 4800
'  If Me.Width < 4800 Then Me.Width = 4800
'  If Me.Height < 2610 Then Me.Height = 2610

  ' Keep the results box in proportion to resized form
  lstResults.Width = Me.Width - 975
  lstResults.Height = Me.Height - 1740
  
  ' Keep the buttons left aligned and anchored to bottom of form.
  cmdNew.Left = Me.Width - 4470
  cmdNew.Top = Me.Height - 1005
  cmdBack.Left = Me.Width - 3150
  cmdBack.Top = Me.Height - 1005
  cmdSelect.Left = Me.Width - 1830
  cmdSelect.Top = Me.Height - 1005
  
End Sub

Private Sub imgInfo_Click()

    Call DisplaySystemInfo

End Sub

Private Sub optSingleLine_Click()

    Call BeginSearch(False)
    
    txtInput.SetFocus

End Sub

Private Sub optTypedown_Click()

    Call BeginSearch(False)
    
    txtInput.SetFocus

End Sub

Private Sub cmbDatabase_Click()

    Call BeginSearch(False)
    
    Call CheckForDataExpiry

End Sub



Private Sub cmdSelect_Click()

    Dim blnReturn As Boolean
     
    blnReturn = HandleSelection(lstResults.ListIndex)
    
    If blnReturn Then
        FormatAddress ("")
        
        Call AcceptPostcode
                
    End If
    
    'txtInput.SetFocus

End Sub

Private Sub lstResults_DblClick()

    Dim blnReturn As Boolean
     
    blnReturn = HandleSelection(lstResults.ListIndex)
    
    If blnReturn Then
        FormatAddress ("")
        Call AcceptPostcode
        Unload Me
    End If

    'txtInput.SetFocus
End Sub

' This is called by VB if the Form is closed by the control bar or menu
Private Sub Form_Unload(Cancel As Integer)

    If (QuitProgram() = False) Then
        Cancel = 1
    End If

  Unhook Me.hWnd
End Sub

'*****************************************************
' Purpose:  Display a non-serious message to the user.
'           Typically, such messages are displayed in a
'           status bar, often along with some form of
'           additional highlighting.
' Inputs:
'   strMessage:   The message text
'*****************************************************

Private Sub DisplayMessage(ByVal strMessage As String)

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("DisplayMessage", MODULE_NAME)

    txtMessage.Text = strMessage
    
' End Error Handler
DisplayMessage_Done:
    Call mProcStack.ExitProc("DisplayMessage")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "DisplayMessage"
    Resume DisplayMessage_Done

End Sub

'*****************************************************
' Purpose:  Display a serious error to the user. Typically,
'           such messages should be displayed as a separate
'           alert dialog that cannot be easily overlooked
'           by the user.
' Inputs:
'   strContext:   Message describing location of error
'   lngErrorCode: QA Pro error code
'*****************************************************

Private Sub ReportError(ByVal strContext As String, lngErrorCode As Long)

Dim strError As String * 100
Dim strMessage As String * 200

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("ReportError", MODULE_NAME)

    Call QA_ErrorMessage(lngErrorCode, strError, 100)
    
    strMessage = strContext & " - " & strError
    
    COAMsgBox strMessage
    
' End Error Handler
ReportError_Done:
    Call mProcStack.ExitProc("ReportError")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "ReportError"
    Resume ReportError_Done

End Sub

'*****************************************************
' Purpose:  Read the current search text from the UI.
' Returns:  The search text
'*****************************************************

Private Function GetSearchText()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("GetSearchText", MODULE_NAME)

    GetSearchText = txtInput.Text
    
' End Error Handler
GetSearchText_Done:
    Call mProcStack.ExitProc("GetSearchText")
    Exit Function

HandleError:
    Process_Error MODULE_NAME, Err, "GetSearchText"
    Resume GetSearchText_Done

End Function

'*****************************************************
' Purpose:  Set the current search text on the UI.
' Inputs:
'   strText:    The new search text
'*****************************************************

Private Sub SetSearchText(ByVal strText As String)

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SetSearchText", MODULE_NAME)
    
    txtInput.Text = strText
    
' End Error Handler
SetSearchText_Done:
    Call mProcStack.ExitProc("SetSearchText")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "SetSearchText"
    Resume SetSearchText_Done

End Sub

'*****************************************************
' Purpose:  Read the current engine mode from the UI
'           and set the global engine mode variable.
'*****************************************************

Private Sub GetEngine()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("GetEngine", MODULE_NAME)

    If optSingleLine.Value = True Then
    
        lngCurrentEngine = qaengine_SINGLELINE
        
    Else
    
        lngCurrentEngine = qaengine_TYPEDOWN
        
    End If
    
' End Error Handler
GetEngine_Done:
    Call mProcStack.ExitProc("GetEngine")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "GetEngine"
    Resume GetEngine_Done

End Sub

'*****************************************************
' Purpose:  Read the current data set from the UI
'           and set the global data set variable.
'*****************************************************

Private Sub GetDataID()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("GetDataID", MODULE_NAME)

    Dim lngGetDataInfoReturn As Long
    Dim strDataID As String * DATASET_LENGTH
    Dim strName As String * SEARCH_LENGTH
    
    lngGetDataInfoReturn = QA_GetData(lngHandle, cmbDatabase.ListIndex, strDataID, DATASET_LENGTH, strName, SEARCH_LENGTH)
                
    ' If QA_GetDataInfo returned no errors
    If lngGetDataInfoReturn = 0 Then
    
        strCurrentData = strDataID
    
    Else
    
        ' Raise Error
        Err.Raise lngGetDataInfoReturn
    
    End If
    
' End Error Handler
GetDataID_Done:
    Call mProcStack.ExitProc("GetDataID")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "GetDataID"
    Resume GetDataID_Done

End Sub

'*****************************************************
' Purpose:  Enable all fields on address capture dialog.
'           This function is called to restore the dialog
'           after having previously been disabled.
'*****************************************************

Private Sub EnableCapture()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("EnableCapture", MODULE_NAME)

    txtInput.Enabled = True
    cmdSelect.Enabled = True
    lstResults.Enabled = True
    
' End Error Handler
EnableCapture_Done:
    Call mProcStack.ExitProc("EnableCapture")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "EnableCapture"
    Resume EnableCapture_Done

End Sub

'*****************************************************
' Purpose:  Disable address capture. This function is
'           called when address capture is unavailable
'           in order to prevent further interaction with
'           the user. This can happen when an invalid data
'           set has been selected or the client/server
'           connection has been broken.
'*****************************************************

Private Sub DisableCapture()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("DisableCapture", MODULE_NAME)

    txtInput.Enabled = False
    cmdSelect.Enabled = False
    lstResults.Enabled = False
    txtMessage.Text = "Address capture unavailable"

    
' End Error Handler
DisableCapture_Done:
    Call mProcStack.ExitProc("DisableCapture")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "DisableCapture"
    Resume DisableCapture_Done

End Sub

'*****************************************************
' Purpose:  Display the current search prompt. This
'           function is also responsible for changing
'           the state of any associated buttons
'           (search/select) and for moving the focus
'           to the search field.
' Inputs:
'   strPrompt
'*****************************************************

Private Sub SetPrompt(ByVal strPrompt As String)

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SetPrompt", MODULE_NAME)
    
    lblPrompt.Caption = strPrompt
    
    ' Update the select button with the appropriate text
    If lngCurrentEngine = qaengine_SINGLELINE And lngPicklistLevel < 0 Then
    
        cmdSelect.Caption = "Search"
    
    Else
    
        cmdSelect.Caption = "Select"
    
    End If
    
    ' Set focus to the end of the search text
    'txtInput.SetFocus
    
' End Error Handler
SetPrompt_Done:
    Call mProcStack.ExitProc("SetPrompt")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "SetPrompt"
    Resume SetPrompt_Done

End Sub

'*****************************************************
' Purpose:  Shutdown the Pro API, releasing all resources.
'*****************************************************

Private Sub ShutdownAPI()

Dim lngCloseReturn As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("ShutdownAPI", MODULE_NAME)
    
    ' Shutdown instance of the API
    lngCloseReturn = QA_Close(lngHandle)
    
    ' Free all resources associated with API
    Call QA_Shutdown
    
' End Error Handler
ShutdownAPI_Done:
    Call mProcStack.ExitProc("ShutdownAPI")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "ShutdownAPI"
    Resume ShutdownAPI_Done

End Sub

'*****************************************************
' Purpose:  Prompts user before exiting.
'*****************************************************

Private Function QuitProgram() As Boolean

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("QuitProgram", MODULE_NAME)

    If True Then 'COAMsgBox("Do you really want to exit?", vbYesNo + vbQuestion, "Exit?") = vbYes Then
    
        Call ShutdownAPI
        QuitProgram = True
        
    Else
        
        QuitProgram = False
    
    End If
    
' End Error Handler
QuitProgram_Done:
    Call mProcStack.ExitProc("QuitProgram")
    Exit Function

HandleError:
    Process_Error MODULE_NAME, Err, "QuitProgram"
    Resume QuitProgram_Done
    
End Function

'*****************************************************
' Purpose:  Display the current search prompt. This
'           function is also responsible for changing
'           the state of any associated buttons
'           (search/select) and for moving the focus
'           to the search field.
' Inputs:
'   strPrompt
'*****************************************************

Private Sub UpdatePicklist()

Dim lngError As Long                            ' Pro error
Dim lngAvailable As Long                        ' Match counts
Dim lngPotential As Long                        ' Used to iterate through picklist items
Dim lngIndex As Long                            ' Picklist information
Dim strPicklistItem As String * SEARCH_LENGTH
Dim strPicklistAddr As String * 55
Dim lngScore As Long
Dim lngFlags As Long
Dim blnCanStep As Boolean                       ' True if the picklist can be stepped into
Dim strPostCode As String * 16                  ' Used to locate the tab character preceeding the postcode
Dim strScore As String * 4                      ' Formatted score
Dim lngChar As Long
Dim lngGetResultDescriptionReturn As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("UpdatePicklist", MODULE_NAME)
    
    ' Call the display function that formats and outputs the
    ' picklist header (search history and number of matches).
    ' The main reason for checking for errors throughout this
    ' function is to detect and act upon a lost connection.
    
     If lngCurrentEngine = qaengine_SINGLELINE And lngPicklistLevel < 0 Then
    
        lngAvailable = 0
        lngPotential = 0
        lngError = 0
        
    Else
    
        lngError = QA_GetSearchStatus(lngHandle, lngAvailable, lngPotential, 0)
    
    End If
    
    If lngError = 0 Then
    
        Call StartPickListDisplay(lngPotential)
        
        ' Iterate for every picklist item
        For lngIndex = 0 To lngAvailable - 1
        
            lngGetResultDescriptionReturn = QA_GetResult(lngHandle, lngIndex, strPicklistItem, SEARCH_LENGTH, lngScore, lngFlags)
            
            strPicklistItem = UnMakeCString(strPicklistItem)
            
            ' If QA_GetResult returned an error
            If lngGetResultDescriptionReturn <> 0 Then
            
                Err.Raise lngGetResultDescriptionReturn
            
            End If
            
            If lngFlags And qaresult_CANSTEP Then
            
                blnCanStep = True
                
            Else
            
                blnCanStep = False
            
            End If
            
            If lngFlags And qaresult_INFORMATION Then
            
                ' Display an informational prompt
                Call SetPickListInformational(lngIndex, blnCanStep, strPicklistItem)
                
            Else
                
                ' Display a picklist item containing address data
                ' The postcode is split out and the score formatted
                lngChar = InStr(strPicklistItem, vbTab)
                
                If lngChar > 0 Then
                    strPostCode = Mid(strPicklistItem, lngChar + 1)
                    strPicklistAddr = Left(strPicklistItem, lngChar - 1)
                Else
                    strPostCode = Space(16)
                    strPicklistAddr = strPicklistItem
                End If
                
                If lngScore > 0 Then
                
                    strScore = lngScore & "%"
                
                Else
                
                    strScore = ""
                
                End If
                
                Call SetPickListAddress(lngIndex, blnCanStep, strPicklistAddr, strPostCode, strScore)
            
            End If
        
        Next
        
        If lngError = 0 Then
        
            ' Determine if the first picklist item should be selected.
            ' We do not want to select (highlight) informational prompts
            ' so as to not distract the user whilst they are typing.
            
            lngIndex = -1
            
            If lngAvailable > 0 Then
            
                lngGetResultDescriptionReturn = QA_GetResult(lngHandle, 0, "", 0, 0, lngFlags)
                                                                       
                ' If QA_GetResult returned no errors
                If lngGetResultDescriptionReturn = 0 Then
                
                    If Not CBool(lngFlags And qaresult_INFORMATION) Then
                    
                        lngIndex = 0
                    
                    End If
                    
                    ' Update the display with the contents of the picklist
                    Call EndPickListDisplay(lngIndex)
                    
                Else
                
                    ' Raise Error
                    Err.Raise lngGetResultDescriptionReturn
                
                End If
                
            End If
            
            ' Check if there are any user messages pending
            If strUserMessage <> "" Then
            
                Call DisplayMessage(strUserMessage)
                strUserMessage = ""
            
            End If
        
        End If
    
    End If
    
    ' If there was an error retrieving the results then
    ' inform the user that the search is being aborted
    If lngError <> 0 Then
    
        Call DisableCapture
        Call ReportError("Address capture cancelled", lngError)
    
    End If

' End Error Handler
UpdatePicklist_Done:
    Call mProcStack.ExitProc("UpdatePicklist")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "UpdatePicklist"
    Resume UpdatePicklist_Done

End Sub

'*****************************************************
' Purpose:  Display the picklist header. This generally
'           consists of the search history (providing
'           the user with "context" for the displayed
'           picklist) and the total number of matches
'           available. This function also resets any
'           previous picklist of results.
' Inputs:
'   lngPotential:   Potential number of matches to report
'*****************************************************

Private Sub StartPickListDisplay(lngPotential As Long)

Dim strBuffer As String * 100

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("StartPickListDisplay", MODULE_NAME)
    
    If lngPicklistLevel >= 0 Then
    
        lblHistory.Caption = arrHistory(lngPicklistLevel)
        
    Else
    
        lblHistory.Caption = ""
        
    End If
    
    If lngPotential > 9999 Then
    
        lblMatchCount.Caption = "Too many matches"
    
    Else
    
        lblMatchCount.Caption = lngPotential & " matches"
        
    End If
    
    lstResults.Clear

' End Error Handler
StartPickListDisplay_Done:
    Call mProcStack.ExitProc("StartPickListDisplay")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "StartPickListDisplay"
    Resume StartPickListDisplay_Done

End Sub

'*****************************************************
' Purpose:  Display an informational picklist entry.
' Inputs:
'   lngIndex:           Picklist offset
'   blnCanStep:         True if item can be stepped into by user
'   strPicklistItem:    Informational prompt
'*****************************************************

Private Sub SetPickListInformational(lngIndex As Long, blnCanStep As Boolean, ByVal strPicklistItem As String)

Dim strMessage As String * SEARCH_LENGTH

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SetPickListInformational", MODULE_NAME)
    
    If blnCanStep Then
    
        strMessage = "+ " & strPicklistItem
                
    Else
    
        strMessage = "  " & strPicklistItem
    
    End If
    
    lstResults.AddItem strMessage

' End Error Handler
SetPickListInformational_Done:
    Call mProcStack.ExitProc("SetPickListInformational")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "SetPickListInformational"
    Resume SetPickListInformational_Done

End Sub

'*****************************************************
' Purpose:  Display a picklist entry (non-informational).
' Inputs:
'   lngIndex:       Picklist offset
'   blnCanStep:     True if item can be stepped into by user
'   strAddress:     Address information (truncated if too long)
'   strPostcode:    Postcode if available
'   strScore:       Formatted score (or "" if beyond level 0)
'*****************************************************

Private Sub SetPickListAddress(lngIndex As Long, blnCanStep As Boolean, ByVal strAddress As String, ByVal strPostCode As String, ByVal strScore As String)

Dim strMessage As String
Dim strCanStep As String

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SetPickListAddress", MODULE_NAME)
    
    'Format picklist items so that postcodes and score appear in columns
    If blnCanStep Then
        
        strMessage = "+ " & strAddress & vbTab & strPostCode & vbTab & strScore
        
    Else
        
        strMessage = "   " & strAddress & vbTab & strPostCode & vbTab & strScore
        
    End If
    
    lstResults.AddItem strMessage
    

' End Error Handler
SetPickListAddress_Done:
    Call mProcStack.ExitProc("SetPickListAddress")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "SetPickListAddress"
    Resume SetPickListAddress_Done

End Sub

'*****************************************************
' Purpose:  Refresh the picklist with the current refinement text.
'*****************************************************

Private Sub HandleRefinement()

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("HandleRefinement", MODULE_NAME)
    
    ' Blank the status line message
    Call DisplayMessage("")
    
    ' Refine and update picklist
    Call PerformRefinement
    
    frmQAProMain.MousePointer = ccHourglass
    Call UpdatePicklist
    frmQAProMain.MousePointer = ccDefault

' End Error Handler
HandleRefinement_Done:
    Call mProcStack.ExitProc("HandleRefinement")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "HandleRefinement"
    Resume HandleRefinement_Done

End Sub

'*****************************************************
' Purpose:  Apply refinement text to the current picklist.
'           If the single line search engine is active and
'           a search is yet to be submitted then no action
'           is taken.
'*****************************************************

Private Sub PerformRefinement()

Dim strSearchText As String * SEARCH_LENGTH
Dim strStatus As String * 100
Dim lngStatus As Long
Dim lngError As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("PerformRefinement", MODULE_NAME)
    
    If lngPicklistLevel >= 0 Then
    
        ' Check whether search text should cleared
        lngError = QA_GetPromptStatus(lngHandle, qapromptint_DYNAMIC, lngStatus, _
            strStatus, 100)
    
        ' Get refinement text as entered by user and then apply to current picklist
        frmQAProMain.MousePointer = ccHourglass
        Call QA_Search(lngHandle, GetSearchText())
        frmQAProMain.MousePointer = ccDefault
        
        If lngStatus = qavalue_FALSE Then
            ' Prompt is non-dynamic so clear search text
            Call SetSearchText("")
        End If
            
    End If

' End Error Handler
PerformRefinement_Done:
    Call mProcStack.ExitProc("PerformRefinement")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "PerformRefinement"
    Resume PerformRefinement_Done

End Sub

'*****************************************************
' Purpose:  Refresh UI with picklist contents. This
'           function is called once the picklists of
'           results has been generated and needs to be
'           displayed to the user. The offset identifies
'           the picklist item to select (highlight) by
'           default. This will be -1 if an informational
'           prompt is being displayed which should not be
'           selected in order to avoid distracting the user.
' Inputs:
'   lngListIndex:   Picklist offset (-1 if none)
'*****************************************************

Private Sub EndPickListDisplay(lngListIndex As Long)

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("EndPickListDisplay", MODULE_NAME)
    
    lstResults.ListIndex = lngListIndex

' End Error Handler
EndPickListDisplay_Done:
    Call mProcStack.ExitProc("EndPickListDisplay")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "EndPickListDisplay"
    Resume EndPickListDisplay_Done

End Sub

'*****************************************************
' Purpose:  Respond to a request to activate the current
'           selection. If a single line search has not
'           yet been performed then this is executed.
'           Otherwise an attempt is made to select the
'           specified picklist item.
' Inputs:
'   lngListIndex:   Currently selected picklist item
'                   (-1 if none
' Returns:
'   0 = User still navigating picklists
'   1 = Leaf-node address selected
'*****************************************************

Private Function HandleSelection(lngListIndex As Long)

Dim lngError As Long                        ' QA Pro error code
Dim blnDisplay As Boolean                   ' True if picklist results to display
Dim strPrompt As String * SEARCH_LENGTH     ' Refinement prompt

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("HandleSelection", MODULE_NAME)
    
    ' Blank the status line message
    Call DisplayMessage("")
    
    If lngPicklistLevel < 0 Then
    
        ' Execute a single line search
        blnDisplay = SingleLineSearch()
        
    Else
    
        ' If no picklist item is currently selected then default to the first
        ' item. This saves the user having to highlight informational prompts
        ' in order to activate them.
        If lngListIndex < 0 Then
        
            lngListIndex = 0
        
        End If
        
        ' Attempt to select the picklist item
        blnDisplay = SelectPickListItem(lngListIndex)
        
    End If
    
    If lngAddress >= 0 Then
    
        ' Reached a final (leaf-node) address
        HandleSelection = True
        
    Else
    
        ' If necessary then update refinement prompt
        ' and refresh picklist contents
        lngError = QA_GetPrompt(lngHandle, 0, strPrompt, SEARCH_LENGTH, 0, "", 0)
        Call SetPrompt(strPrompt)
        
        ' If they were no errors
        If lngError = 0 Then
            If blnDisplay Then
                frmQAProMain.MousePointer = ccHourglass
                Call UpdatePicklist
                frmQAProMain.MousePointer = ccDefault
            End If
            HandleSelection = False
        Else
            ' Raise Error
            Err.Raise lngError
        End If
    End If

' End Error Handler
HandleSelection_Done:
    Call mProcStack.ExitProc("HandleSelection")
    Exit Function

HandleError:
    Process_Error MODULE_NAME, Err, "HandleSelection"
    Resume HandleSelection_Done

End Function

'*****************************************************
' Purpose:  Execute a single line search. This function
'           will submit the specified search expression
'           and then perform an automatic step-in if
'           appropriate to do so.
' Returns:
'   0 = Address capture unavailable
'   1 = Search successfully submitted (but not necessarily
'       matches available)
'*****************************************************

Private Function SingleLineSearch()

Dim lngError As Long                        ' QA Pro error code
Dim strSearch As String * SEARCH_LENGTH     ' Search string
Dim strStatus As String * 100
Dim lngStatus As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SingleLineSearch", MODULE_NAME)
    
    ' Check whether search text should cleared after search
    lngError = QA_GetPromptStatus(lngHandle, qapromptint_DYNAMIC, lngStatus, _
            strStatus, 100)
    
    ' Get the search text and submit to the single line engine
    strSearch = GetSearchText()
    frmQAProMain.MousePointer = ccHourglass
    lngError = QA_Search(lngHandle, strSearch)
    frmQAProMain.MousePointer = ccDefault
    
    If lngError <> 0 Then
        
        'Handle errors
        Call DisableCapture
        Call ReportError("Address capture unavailable", lngError)
        SingleLineSearch = False
    
    Else
        
        If lngStatus = qavalue_FALSE Then
            ' Prompt is non-dynamic so clear search text
            Call SetSearchText("")
        End If
    
        ' Set the search context to indicate that we are at the
        ' first stage of picklist handling (level 0). The original
        ' search string is copied into the search history.
        lngPicklistLevel = 0
        arrHistory(0) = "Searching on... '" & RTrim(strSearch) & "'"
        
        ' Perform any suggested step-ins on the resulting pick-lists
        Call AutoStepInAndFormat
        
        SingleLineSearch = True
        
    End If

' End Error Handler
SingleLineSearch_Done:
    Call mProcStack.ExitProc("SingleLineSearch")
    Exit Function

HandleError:
    Process_Error MODULE_NAME, Err, "SingleLineSearch"
    Resume SingleLineSearch_Done

End Function

'*****************************************************
' Purpose:  Step into a picklist item. The picklist
'           level is incremented and the search
'           history updated. This function should only
'           be called where it is already known that
'           the item is a leaf-node address or
'           unresolvable.
' Inputs:
'   lngListIndex:   Offset of picklist item to be
'                   stepped Into
'*****************************************************

Private Sub PerformStepIn(lngListIndex As Long)

Dim strLine As String * SEARCH_LENGTH   ' Picklist information
Dim lngFlags As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("PerformStepIn", MODULE_NAME)
    
    ' Store the partial address information in the search history
    Call QA_GetResultDetail(lngHandle, lngListIndex, qaresultstr_PARTIALADDRESS, _
                    0, strLine, SEARCH_LENGTH)
    arrHistory(lngPicklistLevel + 1) = strLine
    
    
    Call QA_GetResult(lngHandle, lngListIndex, "", 0, 0, lngFlags)
    
    ' Check if need to report on GBR postcode recodes or AUS bordering locality match
    If strUserMessage = "" Then
    
        If CBool(lngFlags And qaresult_POSTCODERECODED) Then
        
            strUserMessage = "Postcode Updated (recoded)"
            
        ElseIf CBool(lngFlags And qaresult_CROSSBORDERMATCH) Then
            
            strUserMessage = "Bordering Locality match"
        
        End If
    
    End If
    
    ' Increment the level and perform the step-in
    lngPicklistLevel = lngPicklistLevel + 1
    
    frmQAProMain.MousePointer = ccHourglass
    Call QA_StepIn(lngHandle, lngListIndex)
    frmQAProMain.MousePointer = ccDefault
    
    ' Reset the refinement text. Note that this is not cleared if stepping
    ' into an informational prompt as this needs to be carried through.
    If Not CBool(lngFlags And qaresult_INFORMATION) Then
    
        Call SetSearchText("")
    
    End If

' End Error Handler
PerformStepIn_Done:
    Call mProcStack.ExitProc("PerformStepIn")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "PerformStepIn"
    Resume PerformStepIn_Done

End Sub

'*****************************************************
' Purpose:  Attempt to select a picklist item. If the
'           item is invalid or cannot be resolved then
'           a warning is displayed to the user. If the
'           item is not a leaf-node address then it is
'           stepped into.
' Inputs:
'   lngListIndex: Offset of picklist item being selected
' Returns:
'   0 = Error occurred (invalid picklist item specified)
'   1 = Picklist item selected
'*****************************************************

Private Function SelectPickListItem(lngListIndex As Long)

Dim blnSelected As Boolean          ' True if picklist item selected
Dim lngFlags As Long                ' Picklist flag

blnSelected = False

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("SelectPickListItem", MODULE_NAME)
    
    ' Get details of picklist entry (if defined)
    If QA_GetResult(lngHandle, lngListIndex, "", 0, 0, lngFlags) <> 0 Then
    
        Call DisplayMessage("Invalid selection")
    
    ElseIf CBool(lngFlags And qaresult_FULLADDRESS) Then
    
        ' Reached a final (leaf-node) address
        lngAddress = lngListIndex
        blnSelected = True
        
    ElseIf CBool(lngFlags And qaresult_CANSTEP) Then
    
        ' Step into a (resolvable) picklist item and check for
        ' any subsequent automatic step-ins
        Call PerformStepIn(lngListIndex)
        Call AutoStepInAndFormat
        blnSelected = True
        
    ElseIf CBool(lngFlags And qaresult_UNRESOLVABLERANGE) Then
    
        ' Unexpandable range such as in the USA
        Call DisplayMessage("Enter value within displayed range")
    
    ElseIf CBool(lngFlags And qaresult_INCOMPLETEADDR) Then
    
        ' Dummy item (enter building details)
        Call DisplayMessage("Enter value")
    
    ElseIf CBool(lngFlags And qaresult_WARNINFORMATION) Then
    
        ' Handle errors such as No matches, timeout etc.
        Call DisplayMessage("Use another search")
    
    ElseIf CBool(lngFlags And qaresult_INFORMATION) Then
    
        ' Over threshold
        Call DisplayMessage("Type to generate picklist")
    
    End If
    
    SelectPickListItem = blnSelected

' End Error Handler
SelectPickListItem_Done:
    Call mProcStack.ExitProc("SelectPickListItem")
    Exit Function

HandleError:
    Process_Error MODULE_NAME, Err, "SelectPickListItem"
    Resume SelectPickListItem_Done

End Function

'********************************************************************
' Purpose:   Attempt to step-out of the current picklist (step back).
'
'********************************************************************

Private Sub HandleStepOut()

Dim lngError As Long
Dim strPrompt As String * SEARCH_LENGTH
    
' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("HandleStepOut", MODULE_NAME)
    
' Blank the status line message
    Call DisplayMessage("")
    
' Perform a step-out
    If lngCurrentEngine = qaengine_SINGLELINE Then
    
        If lngPicklistLevel >= 0 Then
            
            ' If at the top level picklist then reset the search
            ' otherwise perform a step-out
            lngPicklistLevel = lngPicklistLevel - 1
            If lngPicklistLevel < 0 Then
                lngError = QA_EndSearch(lngHandle)
            Else
                lngError = QA_StepOut(lngHandle)
            End If
        
        End If
        
    Else
    
        If lngPicklistLevel > 0 Then
            lngPicklistLevel = lngPicklistLevel - 1
            lngError = QA_StepOut(lngHandle)
        End If
    
    End If
    
' Clear refinement text, update search prompt and then refresh
' picklist contents
    Call SetSearchText("")
    lngError = QA_GetPrompt(lngHandle, 0, strPrompt, SEARCH_LENGTH, 0, "", 0)
    Call SetPrompt(strPrompt)
    Call PerformRefinement
    
    frmQAProMain.MousePointer = ccHourglass
    Call UpdatePicklist
    frmQAProMain.MousePointer = ccDefault
    
' End Error Handler
HandleStepOut_Done:
    Call mProcStack.ExitProc("HandleStepOut")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "HandleStepOut"
    Resume HandleStepOut_Done
    
End Sub

'******************************************************************
' Purpose:  Format and display the selected (leaf-node) address to
'           the user. This stage should be modified according to
'           the requirements of the host application. Typically,
'           this will result in the selected address being written
'           back to the application's data entry screen.
' Input:    strLayout - Address layout (as defined in QAWSERVE.INI)
'
'******************************************************************

Private Sub FormatAddress(ByVal strLayout As String)

Dim lngLayoutError As Long
Dim lngFormatError As Long
Dim lngCount As Long
Dim lngInfo As Long
Dim strWarning As String

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("FormatAddress", MODULE_NAME)
    
' Attempt to select the requested layout. This is only likely to fail
' if the layout does not actually exist in the configuration file.
    
    lngLayoutError = QA_SetActiveLayout(lngHandle, strLayout)
    If lngLayoutError = 0 Then
        lngFormatError = QA_FormatResult(lngHandle, lngAddress, "", lngCount, lngInfo)
    End If
    
    If lngLayoutError = 0 And lngFormatError = 0 Then
    
        ' Check if need to warn the user about information being lost due to the
        ' layout being too small in size. Otherwise display any user messages
        ' that are pending.
        
        If CBool(lngInfo And qaformat_OVERFLOW) Then
            strWarning = "Address elements lost"
        ElseIf CBool(lngInfo And qaformat_TRUNCATED) Then
            strWarning = "Address elements truncated"
        Else
            strWarning = strUserMessage
        End If
        
        ' Output the formatted address
        Call DisplayAddress(lngCount, strWarning)
    End If

    If lngLayoutError <> 0 Then
        Call ReportError("Cannot format address", lngLayoutError)
    ElseIf lngFormatError <> 0 Then
        Call ReportError("Cannot format address", lngFormatError)
    End If
    
' End Error Handler
FormatAddress_Done:
    Call mProcStack.ExitProc("FormatAddress")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "FormatAddress"
    Resume FormatAddress_Done
        
End Sub

'****************************************************************************
' Purpose:  Display the formatted leaf-node address.
' Input:    lngCount    - number of address lines
'           strWarning  - warning message
'****************************************************************************

Private Sub DisplayAddress(lngCount As Long, ByVal strWarning As String)

Dim lngIndex As Long
Dim lngTYPE As Long
Dim lngError As Long
Dim strLine As String * SEARCH_LENGTH
Dim strLabel As String * SEARCH_LENGTH
Dim strLabelDisplay As String
Dim lngChar As Long

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("DisplayAddress", MODULE_NAME)

'Retrieve and display formatted address
    For lngIndex = 0 To (lngCount - 1)
        
        lngError = QA_GetFormattedLine(lngHandle, lngIndex, strLine, SEARCH_LENGTH, strLabel, SEARCH_LENGTH, lngTYPE)
        
        strLabel = UnMakeCString(strLabel)
        strLine = UnMakeCString(strLine)
        
        If lngIndex = 0 And Len(Trim(strLabel)) = 0 Then
            strLabel = "Address"
        End If
        
        strLabelDisplay = RTrim(strLabel) & " :"
        
        ' frmAddress.lblAddress(lngIndex).Caption = strLabelDisplay
        ' frmAddress.txtAddress(lngIndex).Text = strLine
        
        ' ReDim Preserve aFullAddress(1, lngIndex)
        aFullAddress(0, lngIndex) = strLabelDisplay
        aFullAddress(1, lngIndex) = strLine
        
    Next
    
    For lngIndex = lngCount To (ADDRESSFORM_FIELDS - 1)
        ' frmAddress.lblAddress(lngIndex).Visible = False
        ' frmAddress.txtAddress(lngIndex).Visible = False
    Next

    ' txtMessage = strWarning
    
    frmQAProMain.Hide
    
    
    ' frmAddress.Show 1

' End Error Handler
DisplayAddress_Done:
    Call mProcStack.ExitProc("DisplayAddress")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "DisplayAddress"
    Resume DisplayAddress_Done

End Sub

'****************************************************************************
' Purpose:  Check for data expiry. A warning message is displayed to the user
'           if the current data set is about to expire.
'****************************************************************************'

Private Sub CheckForDataExpiry()

Dim strActive As String * DATASET_LENGTH
Dim strDataID As String * DATASET_LENGTH
Dim lngDaysLeft As Long
Dim lngCount As Long
Dim lngIndex As Long
Dim lngWarning As Long
Dim strMessage As String

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("CheckForDataExpiry", MODULE_NAME)
    
    Call QA_GetActiveData(lngHandle, strActive, DATASET_LENGTH)
    
    Call QA_GetLicensingCount(lngHandle, lngCount, lngWarning)
    
    ' Find the index of the active data set
    lngIndex = 0
    Call QA_GetLicensingDetail(lngHandle, lngIndex, qalicencestr_ID, _
        0, strDataID, DATASET_LENGTH)

    While UnMakeCString(strDataID) <> UnMakeCString(strActive) And lngIndex < lngCount
        
        lngIndex = lngIndex + 1
        Call QA_GetLicensingDetail(lngHandle, lngIndex, qalicencestr_ID, _
            0, strDataID, DATASET_LENGTH)
            
    Wend
    
    ' Put up a warning message if the current data set is going to expire
    ' in the next month
    If QA_GetLicensingDetail(lngHandle, lngIndex, qalicenceint_DATADAYSLEFT, _
                lngDaysLeft, "", 0) = 0 And lngDaysLeft < 31 Then
                
        
        strMessage = "Warning -- " & RTrim(UnMakeCString(strActive)) & _
            " data will expire in " & lngDaysLeft & " day(s)"
        DisplayMessage (strMessage)
        
    End If

' End Error Handler
CheckForDataExpiry_Done:
    Call mProcStack.ExitProc("CheckForDataExpiry")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "CheckForDataExpiry"
    Resume CheckForDataExpiry_Done

End Sub

'*************************************************************
' Purpose:  Check for and act upon any auto step-in or format
'           flags. These flags signify where the first item
'           in a picklist can be automatically selected in
'           order to speed up the address capture process.
'*************************************************************
Private Sub AutoStepInAndFormat()

    Dim lngState As Long            ' Search state
    
' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("AutoStepInAndFormat", MODULE_NAME)
    
    
    ' Loop whilst an auto step-in is recommended. We allow
    ' "past-close" step-ins to occur as the user can navigate
    ' back to close matches from any resulting picklist
    Do While QA_GetSearchStatus(lngHandle, 0, 0, lngState) = 0 And _
            (CBool(lngState And qastate_AUTOSTEPINSAFE) Or _
            CBool(lngState And qastate_AUTOSTEPINPASTCLOSE))
        
        'Warn the user if we have stepped past close matches
        If CBool(lngState And qastate_AUTOSTEPINPASTCLOSE) And _
            strUserMessage = "" Then
            
            strUserMessage = "Step 'Back' for Close Matches"
            
        End If
        
        'Perform the step-in on the first picklist item
        Call PerformStepIn(0)
    Loop
    
    ' Check if we can also select a final (leaf-node) address.
    ' We do not check for the "past close" flag at this stage as the
    ' user is unable to navigate back to see any close matches once a
    ' final address has been selected
    If CBool(lngState And qastate_AUTOFORMATSAFE) Then
        lngAddress = 0
    End If
    
' End Error Handler
AutoStepInAndFormat_Done:
    Call mProcStack.ExitProc("AutoStepInAndFormat")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "AutoStepInAndFormat"
    Resume AutoStepInAndFormat_Done
    
End Sub





'*************************************************************
' Purpose:  Initialise address capture dialog. Typically, this
'           will read in the list of available data sets as
'           this can be a relatively lengthy operation and
'           does not need to be performed everytime the
'           address capture dialog is invoked.
'*************************************************************
Private Sub BeginDialog()

Dim lngDataCountReturn As Long
Dim lngGetDataInfoReturn As Long
Dim lngDefaultDataIndex As Long

Dim lngIndex As Long
Dim lngDataCount As Long
Dim strDataID As String * DATASET_LENGTH
Dim strName As String * SEARCH_LENGTH

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("BeginDialog", MODULE_NAME)
    
    lngDataCountReturn = QA_GetDataCount(lngHandle, lngDataCount)
        
    ' If QA_GetDataCount returned no errors
    If lngDataCountReturn = 0 Then
        
        ' Populate Database combo box
        For lngIndex = 0 To lngDataCount - 1
            
            lngGetDataInfoReturn = QA_GetData(lngHandle, lngIndex, strDataID, DATASET_LENGTH, strName, SEARCH_LENGTH)
                
            ' If QA_GetDataInfo returned no errors
            If lngGetDataInfoReturn = 0 Then
                
                cmbDatabase.AddItem UnMakeCString(strName)
                    
                If UnMakeCString(strDataID) = DEFAULT_DATASET Then
                    
                    lngDefaultDataIndex = lngIndex
                    
                End If
            
            Else
                
                ' Raise Error
                Err.Raise lngGetDataInfoReturn
                
            End If
            
        Next
            
        cmbDatabase.ListIndex = lngDefaultDataIndex
            
        Call CheckForDataExpiry
            
    ' Else QA_GetDataCount failed
    Else
        
        ' Raise Error
        Err.Raise lngDataCountReturn
        
    End If
    

' End Error Handler
BeginDialog_Done:
    Call mProcStack.ExitProc("BeginDialog")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "BeginDialog"
    Resume BeginDialog_Done

End Sub

'**************************************************************
' Purpose:  Display the system information
'**************************************************************
Private Sub DisplaySystemInfo()

Dim lngSysInfoReturn As Long
Dim lngLineCount As Long
Dim lngIndex As Long
Dim strLine As String * BUFFER_LENGTH
Dim strSystemInfo As String

' Begin Error Handler
On Error GoTo HandleError
    Call mProcStack.EnterProc("DisplaySystemInfo", MODULE_NAME)

    lngSysInfoReturn = QA_GenerateSystemInfo(lngHandle, qasysinfo_SYSTEM, lngLineCount)
    
    If lngSysInfoReturn = 0 Then
        For lngIndex = 0 To lngLineCount - 1
        
            Call QA_GetSystemInfo(lngHandle, lngIndex, strLine, BUFFER_LENGTH)
            
            strSystemInfo = strSystemInfo & RTrim(UnMakeCString(strLine)) & vbCrLf
    
        Next
        
        'frmSysInfo.txtSysInfo = strSystemInfo
        
        'frmSysInfo.Show 1
    Else
        ' Raise Error
        Err.Raise lngSysInfoReturn
    End If

' End Error Handler
DisplaySystemInfo_Done:
    Call mProcStack.ExitProc("DisplaySystemInfo")
    Exit Sub

HandleError:
    Process_Error MODULE_NAME, Err, "DisplaySystemInfo"
    Resume DisplaySystemInfo_Done

End Sub



Private Sub txtInput_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
    
        If lngCurrentEngine <> qaengine_SINGLELINE Or GetSearchText <> "" _
            Or lngPicklistLevel <> -1 Then
            
            Call cmdSelect_Click
    
        End If
        
    ElseIf KeyCode = vbKeyUp And lstResults.ListIndex > 0 Then
    
        lstResults.ListIndex = lstResults.ListIndex - 1
        
    ElseIf KeyCode = vbKeyDown And lstResults.ListIndex < lstResults.ListCount - 1 Then
    
        lstResults.ListIndex = lstResults.ListIndex + 1
        
    ElseIf KeyCode = vbKeyPageUp Then
    
        If lstResults.ListIndex > 13 Then
            lstResults.ListIndex = lstResults.ListIndex - 13
        Else
            lstResults.ListIndex = 0
        End If
        
    ElseIf KeyCode = vbKeyPageDown Then
    
        If lstResults.ListIndex < lstResults.ListCount - 13 Then
            lstResults.ListIndex = lstResults.ListIndex + 13
        Else
            lstResults.ListIndex = lstResults.ListCount - 1
        End If
        
    Else
        
        Call HandleRefinement
        
    End If
    
End Sub



Private Sub AcceptPostcode()

  Dim objControl As Control
  
  'Let user know somethings happening
  Screen.MousePointer = vbHourglass
  
  'setup the temp column name variables...these store the field names to put the
  'Afd data in.
  Dim tempforename As String
  Dim tempinitials As String
  Dim tempsurname As String
  Dim tempaddress As String
  Dim tempproperty As String
  Dim tempstreet As String
  Dim templocality As String
  Dim temptown As String
  Dim tempcounty As String
  Dim temptelephone As String
        
        'For all fields that are mapped correctly, store the field names
'        If txtForename.Tag <> 0 Then tempforename = datGeneral.GetColumnName(txtForename.Tag)
'        If txtInitials.Tag <> 0 Then tempinitials = datGeneral.GetColumnName(txtInitials.Tag)
'        If txtSurname.Tag <> 0 Then tempsurname = datGeneral.GetColumnName(txtSurname.Tag)
        If Val(aFullAddress(2, 0)) <> 0 Then tempproperty = datGeneral.GetColumnName(Val(aFullAddress(2, 0)))
        If Val(aFullAddress(2, 1)) <> 0 Then tempstreet = datGeneral.GetColumnName(Val(aFullAddress(2, 1)))
        'If Val(aFullAddress(2, 2)) <> 0 Then templocality = datGeneral.GetColumnName(Val(aFullAddress(2, 2)))
        If Val(aFullAddress(2, 2)) <> 0 Then temptown = datGeneral.GetColumnName(Val(aFullAddress(2, 2)))
        If Val(aFullAddress(2, 3)) <> 0 Then tempcounty = datGeneral.GetColumnName(Val(aFullAddress(2, 3)))
        'If txtTelephone.Tag <> 0 Then temptelephone = datGeneral.GetColumnName(txtTelephone.Tag)
'
        For Each objControl In frmForm.Controls

          'Loop through controls on the user form and copy the text accross if
          'the checkbox is checked.
          If objControl.Tag > 0 And IsNumeric(objControl.Tag) Then
'            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempforename Then objControl.Text = txtForename.Text
'            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempinitials And chkInitials.Value Then objControl.Text = txtInitials.Text
'            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempsurname And chkSurname.Value Then objControl.Text = txtSurname.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempproperty Then objControl.Text = aFullAddress(1, 0)
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempstreet Then objControl.Text = aFullAddress(1, 1)
            'If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = templocality And chkLocality.Value Then objControl.Text = txtLocality.Text
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = temptown Then objControl.Text = aFullAddress(1, 2)
            If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = tempcounty Then objControl.Text = aFullAddress(1, 3)
            'If frmForm.mobjScreenControls.Item(objControl.Tag).ColumnName = temptelephone And chkTelephone.Value Then objControl.Text = txtTelephone.Text
          End If

        Next objControl

      
  
  'Unload the Afd screen
  Unload Me
  
  'Return mousepointer to normal
  Screen.MousePointer = vbDefault
  
End Sub

