VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.1#0"; "CODEJO~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.1#0"; "CODEJO~1.OCX"
Begin VB.UserControl COA_OLE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1350
   ScaleWidth      =   1875
   ToolboxBitmap   =   "COA_OLE.ctx":0000
   Begin XtremeSuiteControls.PushButton ctlCommand 
      Height          =   1200
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1200
      _Version        =   851969
      _ExtentX        =   2117
      _ExtentY        =   2117
      _StockProps     =   79
      Caption         =   "Server"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   10
      DrawFocusRect   =   0   'False
      TextImageRelation=   1
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   1395
      Top             =   135
      _Version        =   851969
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "COA_OLE.ctx":0312
   End
End
Attribute VB_Name = "COA_OLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectSafetyTLB.IObjectSafety

' Event Declarations:
Event Click()

Private msFileName As String

Private miOLEType As OLEType
Private mobjEmbeddedStream As ADODB.Stream
Private mlngColumnID As Long
Private mstrEncryptionKey As String
Private mbEncrypted As Boolean

Public Enum OLEType
  OLE_LOCAL = 0
  OLE_SERVER = 1
  OLE_EMBEDDED = 2
  OLE_UNC = 3
End Enum

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

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

' Left here for backward compatibility
Public Property Get OleOnServer() As Boolean
End Property
  
' Left here for backward compatibility
Public Property Let OleOnServer(ByVal pfNewValue As Boolean)
End Property

Public Property Get OLEType() As OLEType
  OLEType = miOLEType
End Property

Public Property Let OLEType(ByVal piNewValue As OLEType)
  
  Dim strCaption As String
  Dim strKey As String
  
  miOLEType = piNewValue
  
  strKey = IIf(Len(msFileName) > 0, "FULL", "EMPTY") & IIf(Not Enabled, "_DISABLED", "")
  
  Select Case miOLEType
    Case OLE_LOCAL
      ctlCommand.Caption = "(Local)"
'      ctlCommand.Picture = imlIcons.Overlay("LOCAL", strKey)
    Case OLE_SERVER
      ctlCommand.Caption = "(Server)"
'      ctlCommand.Picture = imlIcons.Overlay("SERVER", strKey)
    Case OLE_EMBEDDED
      ctlCommand.Caption = "(Embedded)"
'      ctlCommand.Picture = imlIcons.Overlay("EMBED", strKey)
    Case OLE_UNC
      ctlCommand.Caption = IIf(Len(msFileName) > 0, "(Linked)", "(Link)")
'      ctlCommand.Picture = imlIcons.Overlay("LINK", strKey)
  
  End Select
  
  FileName = msFileName
  
  PropertyChanged "OLEType"
  
End Property

Public Property Get FileName() As String
  FileName = msFileName
End Property

Public Property Let FileName(ByVal psNewValue As String)
  
  Dim Icon As ImageManagerIcon
  Dim fFull As Boolean
  Dim strOLEType As String
  Dim strOLEKey As String
  Dim strDisplayName As String
  Dim strKey As String
  
  msFileName = psNewValue
  fFull = (Len(msFileName) > 0)
 
  ' OLE Type
  Select Case miOLEType
    Case OLE_LOCAL
      strDisplayName = GetFileNameOnly(msFileName)
      strOLEType = "Local"
      strOLEKey = "LOCAL"
    Case OLE_SERVER
      strDisplayName = GetFileNameOnly(msFileName)
      strOLEType = "Server"
      strOLEKey = "SERVER"
    Case OLE_EMBEDDED
      strDisplayName = GetFileNameOnly(msFileName)
      strOLEType = "Embedded"
      strOLEKey = "EMBED"
    Case OLE_UNC
      strDisplayName = msFileName
      strOLEType = IIf(fFull, "Linked", "Link")
      strOLEKey = "LINK"
      
  End Select
     
  strKey = IIf(fFull, "FULL", "EMPTY") '& IIf(Not Enabled, "_DISABLED", "")
     
  ' Display the type on the button - no filename
  ctlCommand.Caption = strOLEType
  ctlCommand.ToolTipText = IIf(Len(strDisplayName) > 0, strDisplayName & " " & strOLEType, "Empty")
  
  If fFull Then
    Set Icon = ImageManager1.Icons.GetImage(2, 32)
  Else
    Set Icon = ImageManager1.Icons.GetImage(1, 32)
  End If
    
  ' Read only or not?
  If UserControl.Enabled Then
    ctlCommand.Picture = Icon.CreatePicture(xtpImageNormal)
  Else
    ctlCommand.Picture = Icon.CreatePicture(xtpImageDisabled)
  End If

  PropertyChanged "FileName"
  
End Property

Private Sub ctlCommand_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    
  ctlCommand.Left = 0
  ctlCommand.Top = 0
  Set mobjEmbeddedStream = New ADODB.Stream

End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal pbEnabled As Boolean)
  
  UserControl.Enabled = pbEnabled
  ctlCommand.Enabled = pbEnabled
'  ctlCommand.MaskColor = IIf(pbEnabled, vbMagenta, vbYellow)

  PropertyChanged "Enabled"

  ' Force refresh display
  OLEType = OLEType

End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Enabled = PropBag.ReadProperty("Enabled", True)
  FileName = PropBag.ReadProperty("FileName", "")
  OleOnServer = PropBag.ReadProperty("OLEOnServer", True)
  OLEType = PropBag.ReadProperty("OLEType", OLE_SERVER)
End Sub

' Resize the constituent controls.
Private Sub UserControl_Resize()
  
 ctlCommand.Height = UserControl.Height
 ctlCommand.Width = UserControl.Width

End Sub

' Write property values to the property bag.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("FileName", msFileName, "")
  Call PropBag.WriteProperty("OLEType", miOLEType, OLE_SERVER)

End Sub

Public Property Let EmbeddedStream(ByRef pobjStream As ADODB.Stream)
  Set mobjEmbeddedStream = pobjStream
End Property

Public Property Get EmbeddedStream() As ADODB.Stream
  Set EmbeddedStream = mobjEmbeddedStream
End Property

Public Property Get ColumnID() As Long
  ColumnID = mlngColumnID
End Property

Public Property Let ColumnID(plngNewValue As Long)
  mlngColumnID = plngNewValue
End Property

' Extracts just the filename from a path
Private Function GetFileNameOnly(pstrFilePath As String) As String
  Dim astrPath() As String
  
  If Len(pstrFilePath) > 0 Then
    astrPath = Split(pstrFilePath, "\")
    GetFileNameOnly = astrPath(UBound(astrPath))
  End If
  
End Function
