Imports System.IO
Imports System.Xml

Namespace Things

  Public Class Table
    Inherits Things.Base

    Public TableType As TableType
    Public ManualSummaryColumnBreaks As Boolean
    Public AuditInsert As Boolean
    Public AuditDelete As Boolean
    Public DefaultOrderID As HCMGuid
    Public DefaultEmailID As HCMGuid
    Public IsRemoteView As Boolean

    Public DependsOnColumns As New Things.Collection

    Public ReadOnly Property Indexes As Things.Collection
      Get
        Return Me.Objects(Things.Enums.Type.Index)
      End Get
    End Property

    Public UpdateStatements As New ArrayList

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return ScriptDB.Consts.UserTable & MyBase.Name
      End Get
    End Property

    Public Overrides Property Name As String
      Get
        Return MyBase.Name
      End Get
      Set(ByVal value As String)
        MyBase.Name = value
      End Set
    End Property

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.Table
      End Get
    End Property

    ' Returns all objects
    <System.ComponentModel.Browsable(False), System.Xml.Serialization.XmlIgnore()> _
    Public ReadOnly Property GetRelation(ByVal ID As HCMGuid) As Things.Relation
      Get

        Dim objRelation As Things.Relation
        Dim bFound As Boolean

        For Each objRelation In Objects(Things.Type.Relation)
          If objRelation.RelationshipType = ScriptDB.RelationshipType.Child Then
            If objRelation.ChildID = ID Then
              bFound = True
              Exit For
            End If
          Else
            If objRelation.ParentID = ID Then
              bFound = True
              Exit For
            End If
          End If
        Next

        Return objRelation

      End Get

    End Property

#Region "Child Objects"

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property Columns()
      Get
        Return Me.Objects(Things.Type.Column)
      End Get
    End Property

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property Validations()
      Get
        Return Me.Objects(Things.Type.Validation)
      End Get
    End Property

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property Views()
      Get
        Return Me.Objects(Things.Type.View)
      End Get
    End Property

#End Region

#Region "Individual objects"

    Public Function Column(ByRef [ColumnID] As HCMGuid) As Things.Column

      Dim objChild As Things.Base

      For Each objChild In Objects(Things.Type.Column)
        If objChild.Type = Type.Column And objChild.ID = ColumnID Then
          Return CType(objChild, Things.Column)
        End If
      Next

      Return Nothing

    End Function

    Public Function Expression(ByRef [ExpressionID] As HCMGuid) As Things.Expression

      Dim objChild As Things.Base

      For Each objChild In Objects(Things.Type.Expression)
        If objChild.Type = Type.Column And objChild.ID = [ExpressionID] Then
          Return CType(objChild, Things.Expression)
        End If
      Next

      Return Nothing

    End Function

    Public Function RecordDescription() As Things.Expression

      Dim objChild As Things.Base

      For Each objChild In Objects(Things.Type.RecordDescription)
        If objChild.Type = Type.RecordDescription Then
          Return CType(objChild, Things.RecordDescription)
        End If
      Next

      Return Nothing

    End Function

    Public Function TableOrderFilter(ByRef RowDetails As Things.ChildRowDetails) As Things.TableOrderFilter

      'ByRef Order As Things.TableOrder, ByRef Filter As Things.Expression _
      '            , ByRef Relation As Things.Relation) As Things.TableOrderFilter

      Dim objChild As Things.Base
      Dim objOFilter As Things.TableOrderFilter

      For Each objChild In Objects(Things.Type.TableOrderFilter)
        objOFilter = CType(objChild, Things.TableOrderFilter)

        If objOFilter.RowDetails.Order Is RowDetails.Order _
            And objOFilter.RowDetails.Filter Is RowDetails.Filter() _
            And objOFilter.RowDetails.Relation Is RowDetails.Relation _
            And objOFilter.RowDetails.RowNumber = RowDetails.RowNumber _
            And objOFilter.RowDetails.RowSelection = RowDetails.RowSelection Then
          Return objOFilter
        End If
      Next

      ' New table filter. Add to the stack and return
      objOFilter = New Things.TableOrderFilter
      objOFilter.RowDetails.Order = RowDetails.Order
      objOFilter.RowDetails.Filter = RowDetails.Filter()
      objOFilter.RowDetails.Relation = RowDetails.Relation
      objOFilter.RowDetails.RowNumber = RowDetails.RowNumber
      objOFilter.RowDetails.RowSelection = RowDetails.RowSelection
      objOFilter.ComponentNumber = Objects(Things.Type.TableOrderFilter).Count + 1
      objOFilter.Parent = Me
      Me.Objects.Add(objOFilter)

      Return objOFilter

    End Function

#End Region

  End Class
End Namespace