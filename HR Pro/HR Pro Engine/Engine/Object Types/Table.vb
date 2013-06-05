﻿Imports System.IO
Imports System.Xml
Imports System.Runtime.InteropServices

Namespace Things

  <ClassInterface(ClassInterfaceType.None), ComVisible(True), Serializable()>
  Public Class Table
    Inherits Base
    Implements COMInterfaces.ITable
    Implements COMInterfaces.IObject

    Public Property TableType As TableType
    Public Property ManualSummaryColumnBreaks As Boolean
    Public Property AuditInsert As Boolean
    Public Property AuditDelete As Boolean
    Public Property DefaultOrderID As Integer
    Public Property DefaultEmailID As Integer
    Public Property IsRemoteView As Boolean
    Public Property RecordDescription() As RecordDescription

    Public Property Indexes As ICollection(Of Index)
    Public Property Columns As ICollection(Of Column)
    Public Property Validations As ICollection(Of Validation)
    Public Property Views As ICollection(Of View)
    Public Property TableOrders As ICollection(Of TableOrder)
    Public Property TableOrderFilters As ICollection(Of TableOrderFilter)
    Public Property Relations As ICollection(Of Relation)
    Public Property Expressions As ICollection(Of Expression)
    Public Property Masks As ICollection(Of Mask)
    Public Property Workflows As ICollection(Of Workflow)
    Public Property Screens As ICollection(Of Screen)
    Public Property DependsOnChildColumns As ICollection(Of Column)
    Public Property DependsOnParentColumns As ICollection(Of Column)

    Public Sub New()
      Indexes = New Collection(Of Index)
      Columns = New Collection(Of Column)
      Validations = New Collection(Of Validation)
      Views = New Collection(Of View)
      TableOrders = New Collection(Of TableOrder)
      TableOrderFilters = New Collection(Of TableOrderFilter)
      Relations = New Collection(Of Relation)
      Expressions = New Collection(Of Expression)
      Masks = New Collection(Of Mask)
      Workflows = New Collection(Of Workflow)
      Screens = New Collection(Of Screen)
      DependsOnChildColumns = New Collection(Of Column)
      DependsOnParentColumns = New Collection(Of Column)
    End Sub

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return ScriptDB.Consts.UserTable & MyBase.Name
      End Get
    End Property

    Public Function GetRelation(ByVal toTableID As Integer) As Relation

      Dim relation As New Relation

      For Each relation In Me.Relations
        If relation.RelationshipType = RelationshipType.Child Then
          If relation.ChildID = toTableID Then
            Return relation
          End If
        Else
          If relation.ParentID = toTableID Then
            Return relation
          End If
        End If
      Next

      'TODO: supposed to be returning blank one if not found?
      Return relation

    End Function

#Region "TableOrderFilter"

    Public Function TableOrderFilter(ByVal RowDetails As ChildRowDetails) As TableOrderFilter

      'ByVal Order As TableOrder, ByVal Filter As Expression _
      '            , ByVal Relation As Relation) As TableOrderFilter

      For Each filer As TableOrderFilter In Me.TableOrderFilters

        If filer.RowDetails.Order Is RowDetails.Order _
            And filer.RowDetails.Filter Is RowDetails.Filter _
            And filer.RowDetails.Relation Is RowDetails.Relation _
            And filer.RowDetails.RowNumber = RowDetails.RowNumber _
            And filer.RowDetails.RowSelection = RowDetails.RowSelection Then
          Return filer
        End If
      Next

      ' New table filter. Add to the stack and return
      Dim filter As New TableOrderFilter
      filter.RowDetails.Order = RowDetails.Order
      filter.RowDetails.Filter = RowDetails.Filter()
      filter.RowDetails.Relation = RowDetails.Relation
      filter.RowDetails.RowNumber = RowDetails.RowNumber
      filter.RowDetails.RowSelection = RowDetails.RowSelection
      filter.ComponentNumber = Me.TableOrderFilters.Count + 1
      filter.Table = Me
      Me.TableOrderFilters.Add(filter)

      Return filter

    End Function

#End Region

#Region "Triggers that are still generated in the system manager need appending to the ones generated in this module. Eventually get rid of as and when port work continues"

    Private msSysMgrInsertTrigger As String
    Private msSysMgrUpdateTrigger As String
    Private msSysMgrDeleteTrigger As String

    Public Property SysMgrDeleteTrigger As String Implements COMInterfaces.ITable.SysMgrDeleteTrigger
      Get
        Return String.Format("---------------------------------------------" & vbNewLine & _
            "-- Script generated by the System Manager" & vbNewLine & _
            "---------------------------------------------" & vbNewLine & _
            "{0}" & vbNewLine, msSysMgrDeleteTrigger)
      End Get
      Set(ByVal value As String)
        msSysMgrDeleteTrigger = value
      End Set
    End Property

    Public Property SysMgrInsertTrigger As String Implements COMInterfaces.ITable.SysMgrInsertTrigger
      Get
        Return String.Format("---------------------------------------------" & vbNewLine & _
            "-- Script generated by the System Manager" & vbNewLine & _
            "---------------------------------------------" & vbNewLine & _
            "{0}" & vbNewLine, msSysMgrInsertTrigger)
      End Get
      Set(ByVal value As String)
        msSysMgrInsertTrigger = value
      End Set
    End Property

    Public Property SysMgrUpdateTrigger As String Implements COMInterfaces.ITable.SysMgrUpdateTrigger
      Get
        Return String.Format("---------------------------------------------" & vbNewLine & _
            "-- Script generated by the System Manager" & vbNewLine & _
            "---------------------------------------------" & vbNewLine & _
            "{0}" & vbNewLine, msSysMgrUpdateTrigger)

      End Get
      Set(ByVal value As String)
        msSysMgrUpdateTrigger = value
      End Set
    End Property

#End Region

  End Class
End Namespace