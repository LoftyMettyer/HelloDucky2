Option Strict On
Option Explicit On

Imports System.ComponentModel

Namespace Enums

	Public Enum ExistingFile
		Overwrite = 0

		<Description("Do not overwrite")>
		DoNotOverwrite = 1

		<Description("Add sequential number to name")>
		AddSequentialToName = 2

		<Description("Append to file")>
		AppendToFile = 3

		<Description("Create new sheet in workbook")>
		CreateNewSheet = 4

	End Enum

End Namespace