VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mblnSuppressPrompt As Boolean
Private mstrHeading As String
Private mintNumberOfCopies As Integer
Private mblnCollate As Boolean
Private mblnPortrait As Boolean
Private mblnGridlines As Boolean
Private mblnShading As Boolean
Private mblnHeadingsOnEveryPage As Boolean
Private mintMarginTop As Integer
Private mintMarginLeft As Integer
Private mintMarginRight As Integer
Private mintMarginBottom As Integer
Private mblnPrintSelectedRowsOnly As Boolean
Private mblnCancelled As Boolean

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const c_DTDefFmt = DT_NOPREFIX 'Or DT_SINGLELINE Or DT_VCENTER

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum EAlignment
  eRight = DT_RIGHT
  eLeft = DT_LEFT
  eCenter = DT_CENTER
End Enum

Private Const TWIPSPERMM = 56.7
Private Const OFFSET_X = 20
Private Const OFFSET_Y = 5
Private Const MARGINCORRECTION = 5 * TWIPSPERMM

Private Const PRINTERFACTOR = 5
Private Const THINLINE = 1 * PRINTERFACTOR
Private Const THICKLINE = 2 * PRINTERFACTOR
  
Public Enum GridCellType
  cellGROUPHEADING = 0
  cellCOLUMNHEADING = 1
  cellDATA = 2
  cellSEPARATOR = 3
End Enum

Private Const COLUMN_ISHEADING = "IsHeading"
Private Const COLUMN_ISPHOTO = "IsPhoto"
Private Const PHOTOSTYLESET = "PhotoSS_"
Private Const RECPROFFOLLOWONCORRECTION = 10

Public Property Get Cancelled() As Boolean
  Cancelled = mblnCancelled
End Property

Public Property Let Cancelled(ByVal blnNewValue As Boolean)
  mblnCancelled = blnNewValue
End Property

Public Property Get PrintSelectedRowsOnly() As Boolean
  PrintSelectedRowsOnly = mblnPrintSelectedRowsOnly
End Property

Public Property Let PrintSelectedRowsOnly(ByVal blnNewValue As Boolean)
  mblnPrintSelectedRowsOnly = blnNewValue
End Property

Public Property Get SuppressPrompt() As Boolean
  SuppressPrompt = mblnSuppressPrompt
End Property

Public Property Let SuppressPrompt(ByVal blnNewValue As Boolean)
  mblnSuppressPrompt = blnNewValue
End Property

Public Property Get Heading() As String
  Heading = mstrHeading
End Property

Public Property Let Heading(ByVal strNewValue As String)
  mstrHeading = strNewValue
End Property

Public Property Get Gridlines() As Boolean
  Gridlines = mblnGridlines
End Property

Public Property Let Gridlines(ByVal intNewValue As Boolean)
  mblnGridlines = intNewValue
End Property

Public Property Get NumberOfCopies() As Integer
  NumberOfCopies = mintNumberOfCopies
End Property

Public Property Let NumberOfCopies(ByVal intNewValue As Integer)
  mintNumberOfCopies = intNewValue
End Property

Public Property Get CollateCopies() As Boolean
  CollateCopies = mblnCollate
End Property

Public Property Let CollateCopies(ByVal intNewValue As Boolean)
  mblnCollate = intNewValue
End Property

Public Property Get Portrait() As Boolean
  Portrait = mblnPortrait
End Property

Public Property Let Portrait(ByVal blnNewValue As Boolean)
  mblnPortrait = blnNewValue
End Property

Public Property Get MarginTop() As Integer
  MarginTop = mintMarginTop
End Property

Public Property Let MarginTop(ByVal intNewValue As Integer)
  mintMarginTop = intNewValue
End Property

Public Property Get MarginBottom() As Integer
  MarginBottom = mintMarginBottom
End Property

Public Property Let MarginBottom(ByVal intNewValue As Integer)
  mintMarginBottom = intNewValue
End Property

Public Property Get MarginLeft() As Integer
  MarginLeft = mintMarginLeft
End Property

Public Property Let MarginLeft(ByVal intNewValue As Integer)
  mintMarginLeft = intNewValue
End Property

Public Property Get MarginRight() As Integer
  MarginRight = mintMarginRight
End Property

Public Property Let MarginRight(ByVal intNewValue As Integer)
  mintMarginRight = intNewValue
End Property

Private Sub Class_Initialize()
  mblnSuppressPrompt = False
  mintNumberOfCopies = 1
  mblnCollate = True
  mblnPortrait = False
  mblnGridlines = True
  mblnShading = True
  mblnHeadingsOnEveryPage = True
  
  If GetSystemMeasurement = "us" Then ' Inches
    mintMarginTop = 1
    mintMarginBottom = 1
    mintMarginLeft = 1
    mintMarginRight = 1
  Else
    mintMarginTop = 10 ' MM
    mintMarginBottom = 10
    mintMarginLeft = 10
    mintMarginRight = 10
  End If
  
  mblnPrintSelectedRowsOnly = False
   
End Sub

Public Sub PrintConfirm()
End Sub



