Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Things.Collections

  <ClassInterface(ClassInterfaceType.None), Serializable()> _
  Public Class BaseCollection
    Inherits System.ComponentModel.BindingList(Of Things.Base)

#Region "System.Xml.Serialization.IXmlSerializable"

    ' Implements System.Xml.Serialization.IXmlSerializable

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

    'Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    '  Return Nothing
    'End Function

    'Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    'End Sub

    'Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml

    '  Dim s As System.Xml.Serialization.XmlSerializer

    '  If Not Me.Items Is Nothing Then
    '    For Each objObject As Things.Base In Me.Items
    '      s = New System.Xml.Serialization.XmlSerializer(objObject.GetType)
    '      s.Serialize(writer, objObject)
    '    Next
    '  End If

    'End Sub

    'Public Sub ToXML(ByVal fileName As String)
    '  Dim serializer As New XmlSerializer(Me.GetType())
    '  Dim namespaces As New XmlSerializerNamespaces()
    '  namespaces.Add("", "")
    '  Using stream As New FileStream(fileName, FileMode.Create)
    '    serializer.Serialize(stream, Me, namespaces)
    '  End Using
    'End Sub

#End Region

  End Class

End Namespace
