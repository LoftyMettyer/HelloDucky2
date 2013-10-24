Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports HR.Intranet.Server.Metadata

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

	Public Tables As ICollection(Of Table)
	Public Columns As ICollection(Of Column)
	Public Relations As ICollection(Of Relation)
	Public ModuleSettings As ICollection(Of ModuleSetting)
	Public UserSettings As ICollection(Of UserSetting)
	Public Functions As ICollection(Of Metadata.Function)
	Public Operators As ICollection(Of Metadata.Operator)

	Public Permissions As ICollection(Of Permission)

End Module