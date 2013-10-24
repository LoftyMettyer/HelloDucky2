Option Strict Off
Option Explicit On

Imports System.Collections.Generic

Module Declarations

  Public gADOCon As ADODB.Connection

  Public datGeneral As New clsGeneral
	Public dataAccess As New clsDataAccess

  Public gsUsername As String
  Public gsActualLogin As String
  Public gsUserGroup As String

	Public gcoTablePrivileges As ICollection(Of CTablePrivilege)

	Public gcolColumnPrivilegesCollection As Collection
  Public gcolLinks As Collection
  Public gcolNavigationLinks As Collection

	Public Tables As ICollection(Of Metadata.Table)
	Public Columns As ICollection(Of Metadata.Column)
	Public Relations As ICollection(Of Metadata.Relation)
	Public ModuleSettings As ICollection(Of Metadata.ModuleSetting)
	Public UserSettings As ICollection(Of Metadata.UserSetting)
	Public Functions As ICollection(Of Metadata.Function)
	Public Operators As ICollection(Of Metadata.Operator)

End Module