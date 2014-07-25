Option Strict On
Option Explicit On

Namespace Enums
	Public Enum ExistingFile
		Overwrite = 0
		DoNotOverwrite = 1
		AddSequentialToName = 2
		AppendToFile = 3
	End Enum

End Namespace