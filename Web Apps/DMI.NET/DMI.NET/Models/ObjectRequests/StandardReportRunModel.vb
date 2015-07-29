Option Strict On
Option Explicit On

Imports DMI.NET.Code.Attributes

Namespace Models.ObjectRequests

	Public Class StandardReportRunModel
		Inherits PromptedValuesModel

		Public Property txtRecordSelectionType As String
		Public Property txtFromDate As String
		Public Property txtToDate As String
		Public Property txtBasePicklistID As Integer
		Public Property txtBasePicklist As String
		Public Property txtBaseFilterID As Integer
		Public Property txtBaseFilter As String
		Public Property txtAbsenceTypes As String
		Public Property txtSRV As String
		Public Property txtShowDurations As String
		Public Property txtShowInstances As String
		Public Property txtShowFormula As String
		Public Property txtOmitBeforeStart As Boolean
		Public Property txtOmitAfterEnd As Boolean
		Public Property txtOrderBy1 As String
		Public Property txtOrderBy1ID As Integer
		Public Property txtOrderBy1Asc As Boolean
		Public Property txtOrderBy2 As String
		Public Property txtOrderBy2ID As Integer
		Public Property txtOrderBy2Asc As Boolean
		Public Property txtMinimumBradfordFactor As Boolean
		Public Property txtMinimumBradfordFactorAmount As Integer
		Public Property txtDisplayBradfordDetail As Boolean
		Public Property txtPrintFPinReportHeader As Boolean
		Public Property txtRecSelCurrentID As Integer
		Public Property action As String
		Public Property txtSend_OutputPreview As String
		Public Property txtSend_OutputFormat As String
		Public Property txtSend_OutputScreen As String
		Public Property txtSend_OutputPrinter As String
		Public Property txtSend_OutputPrinterName As String
		Public Property txtSend_OutputSave As String
		Public Property txtSend_OutputSaveExisting As String
		Public Property txtSend_OutputEmail As String
		Public Property txtSend_OutputEmailAddr As String

		<AllowHtml>
		Public Property txtSend_OutputEmailSubject As String

		<ExcludeChar("/?""<>|*@~[]{{}}#+'¬")>
		<AllowHtml>
		Public Property txtSend_OutputEmailAttachAs As String

		<ExcludeChar("/?""<>|*@~[]{{}}#+'¬")>
		<AllowHtml>
		Public Property txtSend_OutputFilename As String

		Public Property txtFilterName As String
		Public Property txtPicklistName As String
		Public Property txtPersonnelTableID As Integer

	End Class
End Namespace