VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutputOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlngOutputFormat As Long
Private mblnOutputScreen As Boolean
Private mblnOutputPrinter As Boolean
Private mstrOutputPrinterName As String
Private mblnOutputSave As Boolean
Private mlngOutputSaveAs As Long
Private mlngOutputSaveExisting As Long
Private mblnOutputEmail As Boolean
Private mlngOutputEmailAddr As Long
Private mstrOutputEmailSubject As String
Private mstrOutputEmailAttachAs As String
Private mlngOutputEmailFileFormat As Long
Private mstrOutputFilename As String

Public Property Get Format() As Long
  Format = mlngOutputFormat
End Property
Public Property Let Format(lngValue As Long)
  mlngOutputFormat = lngValue
End Property

Public Property Get Screen() As Boolean
  Screen = mblnOutputScreen
End Property
Public Property Let Screen(blnValue As Boolean)
  mblnOutputScreen = blnValue
End Property

Public Property Get Printer() As Boolean
  Printer = mblnOutputPrinter
End Property
Public Property Let Printer(blnValue As Boolean)
  mblnOutputPrinter = blnValue
End Property

Public Property Get PrinterName() As String
  PrinterName = mstrOutputPrinterName
End Property
Public Property Let PrinterName(strValue As String)
  mstrOutputPrinterName = strValue
End Property

Public Property Get Save() As Boolean
  Save = mblnOutputSave
End Property
Public Property Let Save(blnValue As Boolean)
  mblnOutputSave = blnValue
End Property

Public Property Get SaveAs() As Long
  SaveAs = mlngOutputSaveAs
End Property
Public Property Let SaveAs(lngValue As Long)
  mlngOutputSaveAs = lngValue
End Property

Public Property Get SaveExisting() As Long
  SaveExisting = mlngOutputSaveExisting
End Property
Public Property Let SaveExisting(lngValue As Long)
  mlngOutputSaveExisting = lngValue
End Property

Public Property Get Email() As Boolean
  Email = mblnOutputEmail
End Property
Public Property Let Email(blnValue As Boolean)
  mblnOutputEmail = blnValue
End Property

Public Property Get EmailAddr() As Long
  EmailAddr = mlngOutputEmailAddr
End Property
Public Property Let EmailAddr(lngValue As Long)
  mlngOutputEmailAddr = lngValue
End Property

Public Property Get EmailSubject() As String
  EmailSubject = mstrOutputEmailSubject
End Property
Public Property Let EmailSubject(strValue As String)
  mstrOutputEmailSubject = strValue
End Property

Public Property Get EmailAttachAs() As String
  EmailAttachAs = mstrOutputEmailAttachAs
End Property
Public Property Let EmailAttachAs(strValue As String)
  mstrOutputEmailAttachAs = strValue
End Property

Public Property Get FileFormat() As Long
  FileFormat = mlngOutputFileFormat
End Property
Public Property Let FileFormat(lngValue As Long)
  mlngOutputFileFormat = lngValue
End Property

