Option Strict Off
Option Explicit On

Imports HR.Intranet.Server.BaseClasses

Friend Class clsOutputGrid
	Inherits BaseOutputFormat


	''Private WithEvents mgrdPrintGrid As SSDBGrid
	'	Private mobjPrintGrid As clsPrintGrid
	Private mobjParent As clsOutputRun

	Private mstrDefTitle As String
	Private mstrErrorMessage As String
	Private mlngPageCount As Integer

	Private mblnScreen As Boolean
	Private mblnPrinter As Boolean
	Private mstrPrinterName As String
	Private mblnSave As Boolean
	Private mlngSaveExisting As Integer
	Private mblnEmail As Boolean
	Private mstrFileName As String

	Public Sub ClearUp()
		''  Set mgrdPrintGrid = Nothing
		'UPGRADE_NOTE: Object mobjPrintGrid may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		'mobjPrintGrid = Nothing
	End Sub

	Public WriteOnly Property Screen() As Boolean
		Set(ByVal Value As Boolean)
			mblnScreen = Value
		End Set
	End Property

	Public WriteOnly Property DestPrinter() As Boolean
		Set(ByVal Value As Boolean)
			mblnPrinter = Value
		End Set
	End Property

	Public WriteOnly Property PrinterName() As String
		Set(ByVal Value As String)
			mstrPrinterName = Value
		End Set
	End Property

	Public WriteOnly Property Save() As Boolean
		Set(ByVal Value As Boolean)
			mblnSave = Value
		End Set
	End Property


	Public Property SaveExisting() As Integer
		Get
			SaveExisting = mlngSaveExisting
		End Get
		Set(ByVal Value As Integer)
			mlngSaveExisting = Value
		End Set
	End Property

	Public WriteOnly Property Email() As Boolean
		Set(ByVal Value As Boolean)
			mblnEmail = Value
		End Set
	End Property

	Public WriteOnly Property FileName() As String
		Set(ByVal Value As String)
			mstrFileName = Value
		End Set
	End Property

	''Public Sub DataGrid(objNewValue As SSDBGrid)
	''
	''  Dim strDefaultPrinter As String
	''
	''  Set mgrdPrintGrid = objNewValue
	''
	''  If mstrErrorMessage <> vbNullString Then
	''    Exit Sub
	''  End If
	''
	''
	''  If mlngPageCount = 1 Then
	''    Set mobjPrintGrid = New clsPrintGrid
	''  End If
	''
	''  mobjParent.SetPrinter
	''
	''  mobjPrintGrid.Heading = mstrDefTitle
	''  mobjPrintGrid.Grid = mgrdPrintGrid
	''  mobjPrintGrid.SuppressPrompt = (mlngPageCount > 1)
	''  mobjPrintGrid.PrintGrid False
	''
	''  If mobjPrintGrid.Cancelled Then
	''    mstrErrorMessage = "Cancelled by User."
	''  End If
	''
	''  mobjParent.ResetDefaultPrinter
	''
	''End Sub

	'Public Function RecordProfilePage(pfrmRecProfile As Form, _
	''  piPageNumber As Integer, _
	''  pcolStyles As Collection)
	'
	'  On Error GoTo ErrorTrap
	'
	'  Dim fOK As Boolean
	'  Dim strDefaultPrinter As String
	'
	'  fOK = True
	'
	'  If piPageNumber = 1 Then
	'    Set mobjPrintGrid = New clsPrintGrid
	'
	'    mobjParent.SetPrinter
	'
	'    mobjPrintGrid.Heading = mstrDefTitle
	'  End If
	'  mobjPrintGrid.SuppressPrompt = (piPageNumber > 1)
	'
	'  mobjPrintGrid.PrintRecordProfilePage pfrmRecProfile, piPageNumber
	'
	'  fOK = Not mobjPrintGrid.Cancelled
	'
	'  mobjParent.ResetDefaultPrinter
	'
	'TidyUpAndExit:
	'  'Set mobjPrintGrid = Nothing
	'  RecordProfilePage = fOK
	'  Exit Function
	'
	'ErrorTrap:
	'  fOK = False
	'  Resume TidyUpAndExit
	'
	'End Function

	Public WriteOnly Property Parent() As clsOutputRun
		Set(ByVal Value As clsOutputRun)
			mobjParent = Value
		End Set
	End Property

	Public ReadOnly Property ErrorMessage() As String
		Get
			ErrorMessage = mstrErrorMessage
		End Get
	End Property

	'''Private Sub mgrdPrintGrid_PrintInitialize(ByVal ssPrintInfo As SSDataWidgets_B.ssPrintInfo)
	'''  Call mobjPrintGrid.PrintInitialise(ssPrintInfo)
	'''End Sub

	Public Sub AddPage(ByRef strDefTitle As String, ByRef mstrSheetName As String, ByRef colStyles As Collection)
		mstrDefTitle = strDefTitle
		mlngPageCount = mlngPageCount + 1
	End Sub
End Class