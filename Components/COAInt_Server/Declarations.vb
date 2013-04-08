Option Strict Off
Option Explicit On
Module Declarations
	
	Public gADOCon As ADODB.Connection
	
  Public datGeneral As New clsGeneral
	
	Public gsUsername As String
	Public gsActualLogin As String
	Public gsUserGroup As String
	
	Public gcoTablePrivileges As CTablePrivileges
	Public gcolColumnPrivilegesCollection As Collection
	Public gcolLinks As Collection
	Public gcolNavigationLinks As Collection
	
	Public gbEnableUDFFunctions As Boolean
	
	'SQL DatType
	Public Enum SQLDataType
		sqlUnknown = 0 ' ?
		sqlOle = -4 ' OLE columns
		sqlBoolean = -7 ' Logic columns
		sqlNumeric = 2 ' Numeric columns
		sqlInteger = 4 ' Integer columns
		sqlDate = 11 ' Date columns
		sqlVarChar = 12 ' Character columns
		sqlVarBinary = -3 ' Photo columns
		sqlLongVarChar = -1 ' Working Pattern columns
	End Enum
	
	Public Enum GlobalType
		glAdd = 1
		glUpdate = 2
		glDelete = 3
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
		tabTopLevel = 1
		tabChild = 2
		tabLookup = 3
	End Enum
	
	'Control type constants
	Public Enum ControlTypes
		ctlCheck = 1
		ctlCombo = 2
		ctlImage = 4
		ctlOle = 8
		ctlRadio = 16
		ctlSpin = 32
		ctlText = 64
		ctlTab = 128
		ctlLabel = 256
		ctlFrame = 512
		ctlPhoto = 1024
		ctlCommand = 2048
		ctlworkingpattern = 4096
		ctlline = 2 ^ 13
	End Enum
	
	'Column type constants
	Public Enum ColumnTypes
		ColData = 0
		colLookup = 1
		colCalc = 2
		colSystem = 3
		colLink = 4
		'  colWorkingPattern = 5
	End Enum
	
	'Case Conversion Types
	Public Enum CaseConvert
		convNone = 0
		convUpper = 1
		convLower = 2
		convProper = 3
	End Enum
	
	Public UI As clsUI
	
	Public ASRDEVELOPMENT As Boolean
	
	Public Enum CrossTabType
		cttNormal = 0
		cttTurnover = 1
		cttStability = 2
		cttAbsenceBreakdown = 3
	End Enum
	
	Public Enum OutputFormats
		fmtDataOnly = 0
		fmtCSV = 1
		fmtHTML = 2
		fmtWordDoc = 3
		fmtExcelWorksheet = 4
		fmtExcelGraph = 5
		fmtExcelPivotTable = 6
		fmtFixedLengthFile = 7
		fmtCMGFile = 8
		fmtSQLTable = 99
	End Enum
	
	Public Enum OutputDestinations
		desScreen = 0
		desPrinter = 1
		desSave = 2
		desEmail = 3
	End Enum
End Module