﻿Namespace Things
  <Serializable()> _
  Public Class Relation
    Inherits Things.Base

    Public RelationshipType As ScriptDB.RelationshipType
    Public ParentID As HCMGuid
    Public ChildID As HCMGuid

    Public DependantColumns As Things.Collections.Generic
    '    Public DependantOnParent As Boolean = False

    Public Sub New()
      DependantColumns = New Things.Collections.Generic
    End Sub

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return ScriptDB.Consts.UserTable & MyBase.Name
      End Get
    End Property

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Relation
      End Get
    End Property

  End Class

End Namespace
