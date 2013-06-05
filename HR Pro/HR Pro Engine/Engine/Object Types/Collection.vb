Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Things

  <DataObject(True), ClassInterface(ClassInterfaceType.None), Serializable()> _
  Public Class Collection
    Inherits System.ComponentModel.BindingList(Of Things.Base)
    Implements System.Xml.Serialization.IXmlSerializable

    '    Implements System.ComponentModel.ITypedList

    '    Private Shared _ListType As System.Type = GetType(Things.iSystemObject)
    'Private Shared _SortOrder() As String = {"Name, ID"}

    '<NonSerialized()> _
    'Private Properties As PropertyDescriptorCollection

    Public Parent As Things.Base ' iSystemObject
    Public Root As Things.Base 'iSystemObject

    'Public Root As iSystemObject

    'Public Sub New()
    '  MyBase.New()

    '  ' Get the 'shape' of the list. 
    '  ' Only get the public properties marked with Browsable = true.
    '  Dim pdc As PropertyDescriptorCollection = TypeDescriptor.GetProperties(GetType(iSystemObject), New Attribute() {New BrowsableAttribute(True)})
    '  Properties = pdc.Sort()

    'End Sub

    '#Region "ITypedList Implementation"

    '    Public Function GetItemProperties(ByVal listAccessors() As System.ComponentModel.PropertyDescriptor) As System.ComponentModel.PropertyDescriptorCollection Implements System.ComponentModel.ITypedList.GetItemProperties

    '      'If listAccessors Is Nothing Then
    '      '  ' Return the property descriptors for top-level rows
    '      '  Return New PropertyDescriptorCollection(New PropertyDescriptor() {New ObjectPropertyDescriptor("Name"), New ObjectPropertyDescriptor("Description")})
    '      'Else
    '      '  ' Return the property descriptors for second-level and third-level rows
    '      '  Dim parentDescriptorName As String = listAccessors(listAccessors.Length - 1).Name
    '      '  Select Case parentDescriptorName
    '      '    Case "Table"
    '      '      Return New PropertyDescriptorCollection(New PropertyDescriptor() {New ObjectPropertyDescriptor("Column")})

    '      '    Case "Validation"
    '      '      Return New PropertyDescriptorCollection(New PropertyDescriptor() {New ObjectPropertyDescriptor("Validation")})
    '      '    Case Else

    '      '      Throw New Exception("Not implemented: " & parentDescriptorName)
    '      '  End Select
    '      'End If

    '      'Dim pd As System.ComponentModel.PropertyDescriptorCollection
    '      Dim BrowsableAttribute(0) As Attribute
    '      Dim pdc As System.ComponentModel.PropertyDescriptorCollection = Nothing

    '      BrowsableAttribute(0) = New System.ComponentModel.BrowsableAttribute(True)

    '      If listAccessors Is Nothing Then
    '        ' Return properties in sort order
    '        '      pdc = System.ComponentModel.TypeDescriptor.GetProperties(_ListType, BrowsableAttribute)
    '        'Return New PropertyDescriptorCollection(New PropertyDescriptor() {New TablePropertyDescriptor("table")})

    '      Else
    '        'Dim parentDescriptorName As String = listAccessors(listAccessors.Length - 1).Name
    '        'Select Case parentDescriptorName
    '        '  Case "Objects"
    '        '    Return New PropertyDescriptorCollection(New PropertyDescriptor() {New ObjectPropertyDescriptor("Column")})
    '        '    'Case "Column"
    '        '    '  Return New PropertyDescriptorCollection(New PropertyDescriptor() {New ObjectPropertyDescriptor("Column")})


    '        'End Select


    '        ' Return child list shape
    '        '      pdc = ListBindingHelper.GetListItemProperties(listAccessors(0).PropertyType)
    '      End If

    '      Return pdc


    '    End Function

    '    'Public Function GetListName(ByVal listAccessors() As System.ComponentModel.PropertyDescriptor) As String Implements System.ComponentModel.ITypedList.GetListName
    '    '  Return _ListType.Name
    '    'End Function

    '    ' This method is only used in the design-time framework 
    '    ' and by the obsolete DataGrid control.
    '    Public Function GetListName(ByVal listAccessors() As PropertyDescriptor) As String Implements System.ComponentModel.ITypedList.GetListName
    '      Return GetType(iSystemObject).Name
    '    End Function

    '#End Region

    '<System.ComponentModel.Browsable(False)> _
    'Public Shadows Sub Add(ByRef [Object] As Things.Base) '  iSystemObject)

    '  [Object].Parent = Parent
    '  MyBase.Add([Object])

    'End Sub



    Public Function GetObject(ByVal [Type] As Things.Type, ByVal [ID] As HCMGuid) As Things.Base ' Things.iSystemObject

      Dim objChild As Things.Base ' iSystemObject

      For Each objChild In MyBase.Items
        If objChild.ID = ID And objChild.Type = [Type] Then
          Return objChild
        End If
      Next

      Return Nothing

    End Function

#Region "Get specific objects from within the collection"

    Public Function GetSetting(ByVal [Module] As String, ByVal [Parameter] As String) As Things.Setting

      Dim objChild As Things.Base
      Dim objSetting As New Things.Setting

      For Each objChild In MyBase.Items
        If objChild.Type = Type.Setting Then
          objSetting = CType(objChild, Things.Setting)
          If objSetting.Module = [Module] And objSetting.Parameter = Parameter Then
            Return objSetting
          End If
        End If
      Next

      Return objSetting

    End Function

    Public ReadOnly Property Table(ByVal ID As HCMGuid) As Things.Table
      Get

        Dim objChild As Things.Base

        For Each objChild In MyBase.Items
          If objChild.ID = ID And objChild.Type = Type.Table Then
            Return objChild
          End If
        Next

        Return Nothing

      End Get
    End Property

#End Region

#Region "System.Xml.Serialization.IXmlSerializable"

    'Public ReadOnly Property ToXML As String
    '  Get

    '    Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
    '    Dim xmlSerializer As System.Xml.Serialization.XmlSerializer

    '    Try

    '      xmlSerializer = New System.Xml.Serialization.XmlSerializer(Me.GetType)

    '      'Using xmlStream As System.IO.MemoryStream = New System.IO.MemoryStream()
    '      '  xmlSerializer.Serialize(xmlStream, Me)
    '      '  xmlStream.Position = 0

    '      '  xmlDoc.Load(xmlStream)
    '      '  Return xmlDoc.InnerXml
    '      'End Using

    '      Using objStringWriter As System.IO.StringWriter = New System.IO.StringWriter
    '        xmlSerializer.Serialize(objStringWriter, Me)
    '        'Debug.Print(objStringWriter.ToString)
    '        Return objStringWriter.ToString
    '      End Using


    '    Catch ex As Exception
    '      Debug.Print(ex.Message)

    '    End Try

    '  End Get
    'End Property

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml

      Dim s As System.Xml.Serialization.XmlSerializer

      If Not Me.Items Is Nothing Then
        For Each objObject As Things.Base In Me.Items
          s = New System.Xml.Serialization.XmlSerializer(objObject.GetType)
          s.Serialize(writer, objObject)
        Next
      End If

    End Sub


    Public Sub ToXML(ByVal fileName As String)
      Dim serializer As New XmlSerializer(Me.GetType())
      Dim namespaces As New XmlSerializerNamespaces()
      namespaces.Add("", "")
      Using stream As New FileStream(fileName, FileMode.Create)
        serializer.Serialize(stream, Me, namespaces)
      End Using
    End Sub


#End Region

  End Class

End Namespace
