Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text.RegularExpressions.Regex
Imports HR.Intranet.Server.Expressions

Namespace BaseClasses
	Public Class BaseReport
		Inherits BaseForDMI

		Public OutputScreen As Boolean
		Public Property Name() As String

		Private _outputFormat As OutputFormats
		Private _outputFilename As String
		Private _outputPreview As Boolean

		Public Property OutputFormat() As OutputFormats
			Get

				Select Case _outputFormat
					Case OutputFormats.ExcelGraph, OutputFormats.ExcelPivotTable, OutputFormats.ExcelWorksheet
						Return _outputFormat
					Case Else
						Return OutputFormats.ExcelWorksheet
				End Select

			End Get

			<DebuggerStepThrough()> _
			Set(value As OutputFormats)
				_outputFormat = value
			End Set

		End Property

		Public Property OutputPreview() As Boolean
			Get
				Return _outputPreview Or _outputFormat = OutputFormats.DataOnly Or _outputFormat = OutputFormats.WordDoc Or _outputFormat = OutputFormats.HTML Or _outputFormat = OutputFormats.CSV
			End Get

			<DebuggerStepThrough()> _
			Set(value As Boolean)
				_outputPreview = value
			End Set
		End Property

		Public Property OutputFilename As String

			<DebuggerStepThrough()> _
			Get
				Return DownloadFileName
			End Get

			<DebuggerStepThrough()> _
			Set(value As String)
				_outputFilename = value
			End Set

		End Property

		Public ReadOnly Property DownloadFileName As String
			Get
				Dim cleanFileName As String = _outputFilename
				Dim fileFromDefinitionName As String = ""
				Dim regexSearch As String
				Dim r As Regex

				If _outputFilename = "" Then
					regexSearch = New String(Path.GetInvalidFileNameChars()) + New String(Path.GetInvalidPathChars())
					r = New Regex(String.Format("[{0}]", Regex.Escape(regexSearch)))
					fileFromDefinitionName = Name
					cleanFileName = r.Replace(fileFromDefinitionName, "")
				End If

				Select Case _outputFormat
					Case OutputFormats.ExcelGraph, OutputFormats.ExcelPivotTable, OutputFormats.ExcelWorksheet
						cleanFileName = Path.GetFileNameWithoutExtension(cleanFileName) & DefaultFileExtension(_outputFormat)

					Case Else
						cleanFileName = Path.GetFileNameWithoutExtension(cleanFileName) & DefaultFileExtension(OutputFormat)

				End Select

				Return cleanFileName

			End Get
		End Property

		Private Function DefaultFileExtension(OutputType As OutputFormats) As String

			Select Case OutputType
				Case OutputFormats.ExcelGraph, OutputFormats.ExcelPivotTable, OutputFormats.ExcelWorksheet
					Return ".xlsx"
				Case OutputFormats.WordDoc
					Return ".docx"
				Case Else
					Return ".txt"

			End Select

		End Function

#Region "FROM clsGeneral"

		Protected Function IsDateColumn(strType As String, lngTableID As Integer, lngColumnID As Integer) As Boolean

			Select Case strType
				Case "C" 'Column
					Return (Columns.GetById(lngColumnID).DataType = ColumnDataType.sqlDate)

				Case Else	'Calculation
					Dim objCalcExpr = New clsExprExpression(SessionInfo)
					objCalcExpr.Initialise(lngTableID, lngColumnID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
					objCalcExpr.ConstructExpression()
					objCalcExpr.ValidateExpression(True)

					Return (objCalcExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_DATE)

			End Select

		End Function

		Protected Function IsBitColumn(strType As String, lngTableID As Integer, lngColumnID As Integer) As Boolean

			Select Case strType
				Case "C" 'Column
					Return (Columns.GetById(lngColumnID).DataType = ColumnDataType.sqlBoolean)

				Case Else	'Calculation
					Dim objCalcExpr = New clsExprExpression(SessionInfo)
					objCalcExpr.Initialise(lngTableID, lngColumnID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
					objCalcExpr.ConstructExpression()
					objCalcExpr.ValidateExpression(True)

					Return (objCalcExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC)

			End Select

		End Function

		Public Function DoesColumnUseSeparators(plngColumnID As Integer) As Boolean
			Return Columns.GetById(plngColumnID).Use1000Separator
		End Function

		Public Function GetDecimalsSize(plngColumnID As Integer) As Integer
			Return Columns.GetById(plngColumnID).Decimals
		End Function

#End Region

#Region "FROM modIntranet"

		Public Function GetEmailGroupName(lngGroupID As Integer) As String

			Dim rsTemp As DataTable
			Dim strSQL As String

			Try

				strSQL = "SELECT Name FROM ASRSysEmailGroupName " & "WHERE EmailGroupID = " & CStr(lngGroupID)
				rsTemp = DB.GetDataTable(strSQL, CommandType.Text)

				For Each objRow As DataRow In rsTemp.Rows
					'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
					If Not IsDBNull(objRow("Name")) Then
						Return objRow("Name").ToString
					End If
				Next

			Catch ex As Exception
				Throw

			End Try

			Return ""

		End Function

#End Region



	End Class
End Namespace