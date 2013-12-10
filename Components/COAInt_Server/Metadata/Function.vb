Option Strict On
Option Explicit On

Imports System.Collections.Generic

Namespace Metadata
	Public Class [Function]
		Inherits Base
			Public ReturnType As Integer
			Public Parameters As ICollection(Of FunctionParameter)
	End Class
End Namespace