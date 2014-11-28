Option Explicit On
Option Strict On

Imports DMI.NET.Code.Attributes
Imports DMI.NET.Classes
Imports HR.Intranet.Server.Enums

Namespace Models
	''' <summary>
	''' This class contains the data members and member functions used for standard report configuations
	''' </summary>
	''' <remarks></remarks>
	Public Class StandardReportsConfigurationModel

		''' <summary>
		''' Gets OR Sets the report type
		''' </summary>
		''' <remarks></remarks>
		Public ReportType As UtilityType

		''' <summary>
		''' Gets OR Sets the record selection type
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property SelectionType As RecordSelectionType

		''' <summary>
		''' Gets OR Sets the table id
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property TableId As Int32

		''' <summary>
		''' Gets OR Sets the filter Id
		''' </summary>
		''' <value></value>
		''' <returns>The Filter Id</returns>
		''' <remarks></remarks>
		<NonZeroIf("SelectionType", RecordSelectionType.Filter, ErrorMessage:="No filter selected for base table.")> _
		Public Property FilterId As Integer

		''' <summary>
		''' Gets OR Sets the picklist Id
		''' </summary>
		''' <value></value>
		''' <returns>The Picklist Id</returns>
		''' <remarks></remarks>
		<NonZeroIf("SelectionType", RecordSelectionType.Picklist, ErrorMessage:="No picklist selected for base table.")>
		Public Property PicklistId As Integer

		''' <summary>
		''' Gets OR Sets the filter name
		''' </summary>
		''' <value></value>
		''' <returns>The filter name</returns>
		''' <remarks></remarks>
		Public Property FilterName As String

		''' <summary>
		''' Gets OR Sets the picklist name
		''' </summary>
		''' <value></value>
		''' <returns>The picklist name</returns>
		''' <remarks></remarks>
		Public Property PicklistName As String

		''' <summary>
		''' Gets OR Sets the value of Custom date
		''' </summary>
		''' <value></value>
		''' <returns>True if selected, false otherwise</returns>
		''' <remarks></remarks>
		Public Property IsCustomDate As Boolean

		''' <summary>
		''' Gets OR Sets the value of start date
		''' </summary>
		''' <value></value>
		''' <returns>True if selected, false otherwise</returns>
		''' <remarks></remarks>
		Public Property StartDate As String

		''' <summary>
		''' Gets OR Sets the custom start date id
		''' </summary>
		Public Property StartDateId As Integer

		''' <summary>
		''' Gets OR Sets the custom end date id
		''' </summary>
		Public Property EndDateId As Integer

		''' <summary>
		''' Gets OR Sets the value of end date
		''' </summary>
		''' <value></value>
		''' <returns>True if selected, false otherwise</returns>
		''' <remarks></remarks>
		Public Property EndDate As String

		''' <summary>
		''' Gets OR Sets the value of display the picklist or filter title in report header
		''' </summary>
		''' <value></value>
		''' <remarks></remarks>
		Public Property DisplayTitleInReportHeader As Boolean

		''' <summary>
		''' Gets OR Sets the value of default date
		''' </summary>
		''' <value></value>
		''' <returns>True if selected, false otherwise</returns>
		''' <remarks></remarks>
		Public Property IsDefaultDate As Boolean

		''' <summary>
		''' Gets OR Sets the value for display title in report header
		''' </summary>
		''' <value></value>
		''' <returns>True if selected, false otherwise</returns>
		''' <remarks></remarks>
		Public Property IsDisplayTitleInHeader As Boolean

		''' <summary>
		''' Gets OR Sets the output for the report configuration
		''' </summary>
		''' <value></value>
		''' <returns>The report output model</returns>
		''' <remarks></remarks>
		Public Property OutputTab As ReportOutputModel

		''' <summary>
		''' Gets OR Sets the list of columns
		''' </summary>
		''' <value></value>
		''' <returns>The list of columns</returns>
		''' <remarks></remarks>
		Public Property Columns As New List(Of ReportColumnItem)

		''' <summary>
		''' Gets OR Sets the list of absence type
		''' </summary>
		''' <value></value>
		''' <returns>The list of absence type</returns>
		''' <remarks></remarks>
		Public Property AbsenceTypes As New List(Of AbsenceType)

		''' <summary>
		''' Gets OR Sets the absence types as string
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property AbsenceTypesAsString As String

		''' <summary>
		''' Constructor
		''' </summary>
		''' <param name="reportType">The report type</param>
		''' <remarks></remarks>
		Public Sub New(reportType As UtilityType)
			Me.ReportType = reportType
		End Sub

		''' <summary>
		''' Constructor
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
		End Sub


	End Class

	Public Class AbsenceType
		Implements IJsonSerialize

		''' <summary>
		''' Gets OR Sets the Absence type
		''' </summary>
		''' <value></value>
		''' <returns>The Absence type</returns>
		''' <remarks></remarks>
		Public Property Type As String

		''' <summary>
		''' Gets OR Sets the value for if type selected
		''' </summary>
		''' <value></value>
		''' <returns>True id type is selected, False otherwise</returns>
		''' <remarks></remarks>
		Public Property IsSelected As Boolean

		''' <summary>
		''' Gets OR Sets the type concatenated with the section key
		''' </summary>
		''' <value></value>
		''' <returns>The type name with section key</returns>
		''' <remarks></remarks>
		Public Property TypeWithSectionKey As String

		Public Property ID() As Integer Implements IJsonSerialize.ID

	End Class

End Namespace