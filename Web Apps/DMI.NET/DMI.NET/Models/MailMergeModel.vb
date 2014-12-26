Option Explicit On
Option Strict On

Imports System.Collections.ObjectModel
Imports DMI.NET.Code.Attributes
Imports DMI.NET.Classes
Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports HR.Intranet.Server.Extensions
Imports HR.Intranet.Server.Enums
Imports DMI.NET.ViewModels.Reports

Namespace Models

	Public Class MailMergeModel
		Inherits ReportBaseModel

		Public Overrides ReadOnly Property ReportType As UtilityType
			Get
				Return UtilityType.utlMailMerge
			End Get
		End Property

		<MinLength(3, ErrorMessage:="You must select at least one column for your mail merge.")> _
		Public Overrides Property ColumnsAsString As String

		<DisplayName("Template :"), DisplayFormat(ConvertEmptyStringToNull:=False)>
		<Required(ErrorMessage:="No template name entered.")>
		Public Property TemplateFileName As String

		<DisplayName("Pause before merge")>
		Public Property PauseBeforeMerge As Boolean

		<DisplayName("Suppress blank lines")>
		Public Property SuppressBlankLines As Boolean
		Public Property OutputFormat As MailMergeOutputTypes

		<DisplayName("Display on screen")>
		Public Property DisplayOutputOnScreen As Boolean

		<DisplayName("Send to printer")>
		Public Property SendToPrinter As Boolean

		<DisplayName("Engine :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property PrinterName As String

		<DisplayName("Save to file")>
		Public Property SaveToFile As Boolean

		<RequiredIf("SaveToFile", True, ErrorMessage:="No filename entered.")>
		<MaxLength(255, ErrorMessage:="File Name cannot be longer than 255 characters.")>
		<DisplayName("File Name :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		<ExcludeChar("/*?""<>|")>
		Public Property Filename As String

		<DisplayName("Email Address :")>
		<NonZeroIf("OutputFormat", MailMergeOutputTypes.IndividualEmail, ErrorMessage:="No email group selected.")>
		Public Property EmailGroupID As Integer

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		<DisplayName("Email Subject :")>
		Public Property EmailSubject As String

		<DisplayName("Send as attachment")>
		Public Property EmailAsAttachment As Boolean

		<RequiredIf("EmailAsAttachment", True, ErrorMessage:="Email attachment name is required.")>
		<MaxLength(255, ErrorMessage:="Email attachment cannot be longer than 255 characters.")> _
		<DisplayName("Attach As :")>
		<DisplayFormat(ConvertEmptyStringToNull:=False)>
		Public Property EmailAttachmentName As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property WordDocumentPrinter As String

		<DisplayFormat(ConvertEmptyStringToNull:=False)> _
		Public Property DocumentManagementPrinter As String

		Public Overrides Sub SetBaseTable(TableID As Integer)
		End Sub

		Public Property AvailableEmails As Collection(Of ReportTableItem)

		<RegularExpression("True", ErrorMessage:="No output destination selected.")>
		Public ReadOnly Property IsDestinationOK As Boolean
			Get

				If OutputFormat = MailMergeOutputTypes.WordDocument Then
					Return (SaveToFile OrElse SendToPrinter OrElse DisplayOutputOnScreen)
				Else
					Return True
				End If

			End Get
		End Property


		Public Overrides Function GetAvailableTables() As IEnumerable(Of ReportTableItem)

			Dim objItems As New Collection(Of ReportTableItem)

			' Add base table
			Dim objTable = SessionInfo.Tables.Where(Function(m) m.ID = BaseTableID).FirstOrDefault
			objItems.Add(New ReportTableItem With {.id = objTable.ID, .Name = objTable.Name, .Relation = ReportRelationType.Base})

			For Each objParent In SessionInfo.Relations.Where(Function(m) m.ChildID = objTable.ID)
				Dim sParentName = SessionInfo.Tables.GetById(objParent.ParentID).Name
				objItems.Add(New ReportTableItem With {.id = objParent.ParentID, .Name = sParentName, .Relation = ReportRelationType.Parent1})
			Next

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

		Public Overrides Function GetAvailableSortColumns(Self As SortOrderViewModel) As IEnumerable(Of ReportColumnItem)

			Dim objItems As New Collection(Of ReportColumnItem)

			' Add all columns that aren't already included in the sort collection
			For Each objColumn In Columns.Where(Function(m) m.IsExpression = False)
				If SortOrders.Where(Function(m) m.ColumnID = objColumn.ID).Count = 0 Then
					objItems.Add(objColumn)
				End If
			Next

			' Add self to collection if not already there
			If Self.ColumnID > 0 AndAlso objItems.Where(Function(m) m.ID = Self.ColumnID).Count = 0 Then
				Dim objItem = Columns.Where(Function(m) m.ID = Self.ColumnID).FirstOrDefault
				objItems.Add(objItem)
			End If

			Return objItems.OrderBy(Function(m) m.Name)

		End Function

	End Class

End Namespace