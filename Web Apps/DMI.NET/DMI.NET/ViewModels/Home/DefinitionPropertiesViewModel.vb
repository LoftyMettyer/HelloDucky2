Option Strict On
Option Explicit On

Imports HR.Intranet.Server.Enums
Imports System.Collections.ObjectModel
Imports System.ComponentModel

Namespace ViewModels
	Public Class DefinitionPropertiesViewModel
		Implements IJsonSerialize

		Public Property ID As Integer Implements IJsonSerialize.ID
		Public Property Type As UtilityType

		<DisplayName("Name :")>
		Public Property Name As String

		<DisplayName("Created Date :")>
		Public Property CreatedDate As String

		<DisplayName("Last Saved Date :")>
		Public Property LastSaveDate As String

		<DisplayName("Last Run Date :")>
		Public Property LastRunDate As String

		<DisplayName("Current Usage :")>
		Public Property Usage As Collection(Of DefinitionPropertiesViewModel)

		Public ReadOnly Property LastRunHidden As String
			Get

				If Type = UtilityType.utlCalculation Or Type = UtilityType.utlFilter Or Type = UtilityType.utlPicklist Then
					Return "visibility: hidden"
				End If

				Return ""

			End Get
		End Property

		Public ReadOnly Property NameUrlDecoded As String
			Get
				Return HttpUtility.UrlDecode(Name)
			End Get
		End Property
	End Class

End Namespace
