Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata

Module Declarations

	Public gsUsername As String
  Public gsActualLogin As String
  Public gsUserGroup As String

	Public gcoTablePrivileges As ICollection(Of TablePrivilege)

	Public gcolColumnPrivilegesCollection As Collection
	Public gcolLinks As List(Of Link)
	Public gcolNavigationLinks As List(Of Link)

	Public Functions As ICollection(Of Metadata.Function)
	Public Operators As ICollection(Of Metadata.Operator)

	Public Modules As List(Of ModuleSetting)


End Module