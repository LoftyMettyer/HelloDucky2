Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Enums

Namespace Metadata
	Public Class [Function]
		Inherits Base
		Public ReturnType As ExpressionValueTypes
		Public Parameters As ICollection(Of FunctionParameter)
	End Class
End Namespace