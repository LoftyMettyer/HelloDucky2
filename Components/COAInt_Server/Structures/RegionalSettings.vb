Imports System.Globalization

Namespace Structures
	Public Structure RegionalSettings
		Public Culture As CultureInfo

		Public ReadOnly Property DateFormat As DateTimeFormatInfo
			Get
				Return Culture.DateTimeFormat
			End Get
		End Property

		Public ReadOnly Property DateSeparator As String
			Get
				Return "/"
			End Get
		End Property

	End Structure
End Namespace