Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports System.IO

Namespace BaseClasses
	Public Class BaseReport
		Inherits BaseForDMI

		Public OutputScreen As Boolean

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
				Return _outputPreview Or (_outputFormat = OutputFormats.fmtDataOnly Or OutputScreen)
			End Get

			<DebuggerStepThrough()> _
			Set(value As Boolean)
				_outputPreview = value
			End Set
		End Property

		Public Property OutputFilename As String

			<DebuggerStepThrough()> _
			Get
				Return _outputFilename
			End Get

			<DebuggerStepThrough()> _
			Set(value As String)
				_outputFilename = value
			End Set

		End Property

		Public ReadOnly Property DownloadFileName As String
			Get
				If _outputFilename = "" Then
					Return "ReportOutput.xlsx"
				Else

					Select Case _outputFormat
						Case OutputFormats.fmtExcelGraph, OutputFormats.fmtExcelPivotTable, OutputFormats.fmtExcelWorksheet
							Return Path.GetFileName(_outputFilename)
						Case Else
							Return Path.GetFileNameWithoutExtension(_outputFilename) + ".xlsx"
					End Select

				End If

			End Get
		End Property


	End Class
End Namespace