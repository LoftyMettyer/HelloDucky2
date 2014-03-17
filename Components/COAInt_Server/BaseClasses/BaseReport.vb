Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports System.IO

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
					Case OutputFormats.fmtExcelGraph, OutputFormats.fmtExcelPivotTable, OutputFormats.fmtExcelWorksheet
						Return _outputFormat
					Case Else
						Return OutputFormats.fmtExcelWorksheet
				End Select

			End Get

			<DebuggerStepThrough()> _
			Set(value As OutputFormats)
				_outputFormat = value
			End Set

		End Property

		Public Property OutputPreview() As Boolean
			Get
				Return _outputPreview Or _outputFormat = OutputFormats.fmtDataOnly Or _outputFormat = OutputFormats.fmtWordDoc Or _outputFormat = OutputFormats.fmtHTML Or _outputFormat = OutputFormats.fmtCSV
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
				Dim sName As String = _outputFilename

				If _outputFilename = "" Then
					sName = Name
				End If

				Select Case _outputFormat
					Case OutputFormats.fmtExcelGraph, OutputFormats.fmtExcelPivotTable, OutputFormats.fmtExcelWorksheet
						Return Path.GetFileNameWithoutExtension(sName) & DefaultFileExtension(_outputFormat)

					Case Else
						Return Path.GetFileNameWithoutExtension(sName) & DefaultFileExtension(OutputFormat)

				End Select


			End Get
		End Property

		Private Function DefaultFileExtension(OutputType As OutputFormats) As String

			Select Case OutputType
				Case OutputFormats.fmtExcelGraph, OutputFormats.fmtExcelPivotTable, OutputFormats.fmtExcelWorksheet
					Return ".xlsx"
				Case OutputFormats.fmtWordDoc
					Return ".docx"
				Case Else
					Return ".txt"

			End Select

		End Function

#Region "FROM clsGeneral"

		Protected Function IsDateColumn(strType As String, lngTableID As Integer, lngColumnID As Integer) As Boolean

			Select Case strType
				Case "C" 'Column
					Return (Columns.GetById(lngColumnID).DataType = SQLDataType.sqlDate)

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
					Return (Columns.GetById(lngColumnID).DataType = SQLDataType.sqlBoolean)

				Case Else	'Calculation
					Dim objCalcExpr = New clsExprExpression(SessionInfo)
					objCalcExpr.Initialise(lngTableID, lngColumnID, ExpressionTypes.giEXPR_RUNTIMECALCULATION, ExpressionValueTypes.giEXPRVALUE_UNDEFINED)
					objCalcExpr.ConstructExpression()
					objCalcExpr.ValidateExpression(True)

					Return (objCalcExpr.ReturnType = ExpressionValueTypes.giEXPRVALUE_LOGIC)

			End Select

		End Function

#End Region

	End Class
End Namespace