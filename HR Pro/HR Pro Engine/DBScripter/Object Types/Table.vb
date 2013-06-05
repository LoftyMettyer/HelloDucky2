Imports System.IO
Imports System.Xml

Namespace Things

  Public Class Table
    Inherits Things.Base

    Public TableType As Integer
    Public ManualSummaryColumnBreaks As Boolean
    Public AuditInsert As Boolean
    Public AuditDelete As Boolean
    Public DefaultOrderID As HCMGuid
    Public RecordDescription As Things.RecordDescription
    Public DefaultEmailID As HCMGuid

    Private mbIsRemoteView As Boolean

    Public Overrides ReadOnly Property PhysicalName As String
      Get
        Return Consts.UserTable & MyBase.Name
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

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property Columns()
      Get
        Return Me.Objects(Things.Type.Column)
      End Get
    End Property

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property Views()
      Get
        Return Me.Objects(Things.Type.View)
      End Get
    End Property

    'Public ReadOnly Property ToXML As String
    '  Get

    '    Dim objPropertyVal As Object

    '    Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument
    '    Dim typObj As System.Type = Me.GetType

    '    Dim declaration As System.Xml.XmlNode = xmlDoc.CreateNode(System.Xml.XmlNodeType.XmlDeclaration, Nothing, Nothing)
    '    xmlDoc.AppendChild(declaration)

    '    Dim xmlRoot As System.Xml.XmlElement = xmlDoc.CreateElement(typObj.Name)
    '    xmlDoc.AppendChild(xmlRoot)

    '    Dim piObjs As System.Reflection.PropertyInfo() = typObj.GetProperties()
    '    For Each piObj As System.Reflection.PropertyInfo In piObjs

    '      Debug.Print(piObj.Name)

    '      If piObj.GetIndexParameters.Length = 0 And piObj.CanWrite Then
    '        objPropertyVal = piObj.GetValue(Me, Nothing)

    '        If Not objPropertyVal Is Nothing Then

    '          Dim xmlSubElement As System.Xml.XmlElement = xmlDoc.CreateElement(piObj.Name)
    '          xmlSubElement.InnerText = objPropertyVal.ToString
    '          xmlRoot.AppendChild(xmlSubElement)
    '        End If

    '      End If

    '    Next

    '    ToXML = xmlDoc.InnerXml

    '  End Get
    'End Property

    Public Property IsRemoteView As Boolean
      Get
        Return mbIsRemoteView
      End Get

      Set(ByVal value As Boolean)
        mbIsRemoteView = value
      End Set
    End Property

  End Class
End Namespace