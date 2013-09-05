VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

Private mbPrinterPrompt As Boolean
Private mbPrinterConfirm As Boolean
Private mlngPageNum As Long
Private mlngBottom As Long
Private msUserName As String

Public Enum PrintFontFormat
  pffNormal = 0
  pffBold = 1
  pffNonBold = 2
End Enum

Const giPRINT_XINDENT = 1000
Const giPRINT_YINDENT = 1000
Const giPRINT_XSPACE = 500
Const giPRINT_YSPACE = 100

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    pdwSupportedOptions As Long, _
                                                    pdwEnabledOptions As Long)

    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                pdwEnabledOptions = IIf(m_fSafeForScripting, _
              INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, _
               IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, _
              INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As Long, _
                                                    ByVal dwOptionsSetMask As Long, _
                                                    ByVal dwEnabledOptions As Long)
    Dim Rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
             INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, _
          IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> _
              INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Public Property Get IsOK() As Boolean
  Dim pstrError As String
  
  On Error GoTo Init_ERROR
  
  If Printer.CurrentX = 0 Then
  End If
  
  IsOK = True
  
  Exit Sub
  
Init_ERROR:
  Select Case Err.Number
    Case 482: pstrError = "Printer Error : Please check your printer connection."
    Case Else: pstrError = Err.Description
  End Select
  
  'MsgBox pstrError, vbExclamation + vbOKOnly, App.Title
  MsgBox pstrError, vbExclamation + vbOKOnly, "OpenHR Intranet"
  
  IsOK = False
  
End Property

Public Function PrintStart(ByVal pbForceOptions As Boolean, psUserName As String) As Boolean
  Dim objPrinter As Printer
  
  mlngPageNum = 0
  msUserName = psUserName
  
  mbPrinterPrompt = GetSetting("HR Pro", "Printer", "Prompt", True)
  mbPrinterConfirm = GetSetting("HR Pro", "Printer", "Confirm", False)
  
  If mbPrinterPrompt Or pbForceOptions Then
    With frmPrintOptions
      .PrintDefinition
      .Show vbModal
      If .Cancelled Then
        PrintStart = False
      Else
        Printer.Copies = .PrintCopies
        Printer.Orientation = IIf(.PrintPortrait, vbPRORPortrait, vbPRORLandscape)
        PrintStart = True
      End If
    End With
  
    'Tidy up
    Unload frmPrintOptions
    Set frmPrintOptions = Nothing
  Else
    PrintStart = True
    Printer.Copies = 1
    Printer.Orientation = vbPRORPortrait
  End If


End Function


Public Sub ClipboardStart()
  Clipboard.Clear
  
End Sub

Public Function ClipboardGetText() As String
  ClipboardGetText = Clipboard.GetText
  
End Function

Public Sub ClipboardSetText(psText As String)
  Clipboard.SetText psText
  
End Sub


Public Sub PrintHeader(strTitle As String)

  Dim fOk As Boolean

  With Printer
    .Font = "Verdana"
    
    ' Print the header.
    .FontSize = 12
    .FontBold = True
    .FontUnderline = True
    
    .CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(strTitle) / 2)
    .CurrentY = giPRINT_YINDENT
    Printer.Print strTitle & vbCrLf

    'This line will work out the Y difference before the title was printed and
    'after the title was printed.  It will be times by 3 so that a blank space
    'is left just above the footer text
    mlngBottom = .ScaleHeight - ((.CurrentY - giPRINT_YINDENT) * 3)
    
    
    .FontSize = 10
    .FontBold = False
    .FontUnderline = False

  End With

End Sub

Public Sub PrintNormal(Optional strText As String)
  PrintText strText
End Sub

Public Sub PrintStraight(Optional strText As String)
  ' No pre-processing of the text, using the current prointer obbject settings.
  Printer.Print strText
  
End Sub


Public Property Let CurrentX(pdblX As Double)
  Printer.CurrentX = pdblX
End Property
Public Property Get CurrentX() As Double
  CurrentX = Printer.CurrentX
End Property

Public Property Let PrinterBold(pfValue As Boolean)
  Printer.FontBold = pfValue
End Property

Public Property Get PrinterBold() As Boolean
  PrinterBold = Printer.FontBold
End Property


Public Property Get CurrentY() As Double
  CurrentY = Printer.CurrentY
End Property


Public Property Let CurrentY(pdblY As Double)
  Printer.CurrentY = pdblY
  
End Property


Public Sub KillDoc()
  Printer.KillDoc
  
End Sub


Public Sub PrintTitle(strText As String)

  Printer.Print
  Printer.FontSize = 12
  Printer.FontUnderline = True
  PrintText strText, pffBold
  Printer.FontUnderline = False
  Printer.FontSize = 10
  Printer.Print

End Sub

Public Sub PrintEnd()

  With Printer
    Do While .CurrentY < mlngBottom
      Printer.Print vbCr
    Loop
    Call FooterText
    .EndDoc
  End With

End Sub


Public Sub PrintConfirm(pstrMessage As String, pstrHeading As String)

  Dim strMBText As String
  Dim msb As frmMessageBox
  Dim iShowMeAgain As Integer
  
  If mbPrinterConfirm Then
    strMBText = pstrMessage & " printing complete." _
      & vbCrLf & vbCrLf & "(" & Printer.DeviceName & ")"
  
    iShowMeAgain = IIf(mbPrinterConfirm, 1, 0)
    If iShowMeAgain = 1 Then
      Set msb = New frmMessageBox
      msb.MessageBox strMBText, vbInformation, pstrHeading, iShowMeAgain, "Don't show me this confirmation again." & vbCrLf
      mbPrinterConfirm = IIf(iShowMeAgain = 1, True, False)
      SaveSetting "HR Pro", "Printer", "Confirm", mbPrinterConfirm
      Set msb = Nothing
    End If
  End If

End Sub

Private Sub CheckEndOfLine(lngLeft As Long, lngRight As Long)
  'Carriage Return if line too long
  If Printer.CurrentX > lngRight Then
    Printer.Print " "
    Printer.CurrentX = lngLeft
  Else
    Printer.Print " ";
  End If
End Sub
Private Sub CheckEndOfPage()
  If Printer.CurrentY > mlngBottom Then
    Call FooterText
    Printer.NewPage
    Printer.CurrentY = giPRINT_YINDENT
    Printer.CurrentX = giPRINT_XINDENT
  End If
End Sub

Private Sub FooterText()
  
  Dim strPageNum As String
  Dim objGeneralFunctions As New clsGeneral
  
  mlngPageNum = mlngPageNum + 1
  strPageNum = "Page " & CStr(mlngPageNum)

  Printer.FontSize = 8
  Printer.Print " "
  Printer.FontBold = False
  
  Printer.CurrentX = giPRINT_XINDENT
  Printer.Print "Printed on " & Format(Now, objGeneralFunctions.LocaleDateFormat) & _
                " at " & Format(Now, "hh:nn") & " by " & msUserName
  
  Printer.CurrentX = (Printer.ScaleWidth - giPRINT_XINDENT) - Printer.TextWidth(strPageNum)
  Printer.Print strPageNum

  Printer.FontSize = 10

  Set objGeneralFunctions = Nothing
  
End Sub


Public Sub PrintBold(Optional strText As String)
  PrintText strText, pffBold
End Sub

Public Sub PrintNonBold(Optional strText As String)
  PrintText strText, pffNonBold
End Sub



Private Sub PrintText(strText As String, Optional pffStyle As PrintFontFormat = pffNormal)

  Dim strChar As String
  Dim intCount As Integer
  Dim lngLeft As Long
  Dim lngCentre As Long
  Dim lngRight As Long
  
  With Printer
    
    'Set up margins
    lngLeft = giPRINT_XINDENT
    lngCentre = .ScaleWidth / 2
    lngRight = .ScaleWidth - (giPRINT_XINDENT * 1.5)

    .FontBold = (pffStyle <> pffNonBold)
    .CurrentX = lngLeft

    intCount = 1
    Do While intCount <= Len(strText)
      strChar = Mid$(strText, intCount, 1)


      Select Case strChar
      Case " "
        'Check if gone past right margin
        Call CheckEndOfLine(lngLeft, lngRight)
        
      Case vbTab
        'Jump to the centre of the page
        'If gone past centre then print a space
        If .CurrentX <= lngCentre Then
          .CurrentX = lngCentre
        Else
          Call CheckEndOfLine(lngLeft, lngRight)
        End If
        
      Case ":"
        'Only bold after ':' if force bold
        Printer.Print strChar;
        .FontBold = (pffStyle = pffBold)

      Case Else
        Printer.Print strChar;
      
      End Select

      'Make sure that we are not less than left margin !
      If .CurrentX < lngLeft Then
        .CurrentX = lngLeft
      End If
      
      Call CheckEndOfPage
      intCount = intCount + 1
    Loop

    Printer.Print
    Call CheckEndOfPage
    
  End With

End Sub

