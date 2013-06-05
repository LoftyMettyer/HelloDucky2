Imports System.Xml.Serialization
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Things

  <Serializable(), ClassInterface(ClassInterfaceType.None)> _
  Public MustInherit Class Base
    Implements COMInterfaces.iObject
    '    Implements ICloneable

    '    Implements Things.iSystemObject
    '    Implements System.Xml.Serialization.IXmlSerializable


    '#Region "Base UITypeEditor"

    '  'Public Overloads Overrides Function GetEditStyle(ByVal context As ITypeDescriptorContext) As UITypeEditorEditStyle
    '  '  If (Not context Is Nothing And Not context.Instance Is Nothing) Then
    '  '    Return UITypeEditorEditStyle.Modal
    '  '  End If
    '  '  Return MyBase.GetEditStyle(context)
    '  'End Function


    '#End Region

    '  Public Parents As ObjectCollection
    '    Public Status As System.Data.DataRowState
    Public SchemaName As String

    <System.Xml.Serialization.XmlIgnore()> _
    Public Encrypted As Boolean = False

    <System.Xml.Serialization.XmlIgnore()> _
    Public Tuning As ScriptDB.Tuning

    Private msDescription As String
    Private mID As HCMGuid
    Private msName As String
    Public NameInDB As String
    Private miSubType As Things.Type

    <System.Xml.Serialization.XmlIgnore()> _
    Public State As System.Data.DataRowState

    Private mobjChildObjects As Things.Collections.Generic
    Private mobjParent As Things.Base
    Private mobjRoot As Things.Base
    Private mbIsSelected As Boolean
    Private msPhysicalName As String

    <System.Xml.Serialization.XmlIgnore()> _
    Public Overridable ReadOnly Property PhysicalName As String Implements iObject.PhysicalName
      Get
        Return msName
      End Get
    End Property

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public Overridable Property Parent() As Things.Base
      Get
        Return mobjParent
      End Get
      Set(ByVal value As Things.Base)
        mobjParent = value
      End Set
    End Property

    Public Property Description() As String
      Get
        Return msDescription
      End Get
      Set(ByVal value As String)
        msDescription = value
      End Set
    End Property

    <System.ComponentModel.Browsable(False)> _
    Public Property ID() As HCMGuid
      Get
        Return mID
      End Get
      Set(ByVal value As HCMGuid)
        mID = value
      End Set
    End Property

    <System.ComponentModel.DisplayName("Name")> _
    Public Overridable Property Name() As String Implements COMInterfaces.iObject.Name
      Get
        Return msName
      End Get
      Set(ByVal value As String)
        msName = value
      End Set
    End Property

    'Public Overridable Sub Edit() Implements Things.iSystemObject.Edit

    '  'Dim objForm As EditObject
    '  'objForm = New EditObject
    '  'objForm.ShowDialog()

    '  ' Creates a new component.
    '  '    Dim myNewImage As New MyImage()

    '  ' Gets the attributes for the component.
    '  'Dim attributes As System.ComponentModel.AttributeCollection = System.ComponentModel.TypeDescriptor.GetAttributes(Me)

    '  '' Prints the name of the editor by retrieving the EditorAttribute from the AttributeCollection. 
    '  'Dim myAttribute As System.ComponentModel.EditorAttribute = CType(attributes(GetType(System.ComponentModel.EditorAttribute)), System.ComponentModel.EditorAttribute)
    '  'Console.WriteLine(("The editor for this class is: " & myAttribute.EditorTypeName))

    '  '    Me.EditValue(Me.GetEdit, Me)

    'End Sub

    '<System.ComponentModel.ImmutableObject(True)> _
    Public MustOverride ReadOnly Property Type() As Things.Type

    Public Overridable Property SubType() As Things.Type
      Get
        Return miSubType
      End Get
      Set(ByVal value As Things.Type)
        miSubType = value
      End Set
    End Property

    <System.Xml.Serialization.XmlElement()> _
    Public Property Objects() As Things.Collections.Generic
      Get
        Return mobjChildObjects
      End Get
      Set(ByVal value As Things.Collections.Generic)
        mobjChildObjects = value
      End Set
    End Property

    Public ReadOnly Property Objects(ByVal Index As Integer) As Things.Base
      Get
        Return mobjChildObjects.Item(Index)
      End Get
    End Property

    <System.Xml.Serialization.XmlIgnore()> _
    Public Property Objects(ByVal Type As Things.Type) As Things.Collections.Generic
      Get

        Dim objCollection As Things.Collections.Generic
        Dim objObject As Object

        objCollection = New Things.Collections.Generic
        For Each objObject In mobjChildObjects
          If objObject.Type = [Type] Then
            objCollection.Add(objObject)
          End If
        Next

        Return objCollection
      End Get
      Set(ByVal value As Things.Collections.Generic)
        mobjChildObjects = value
      End Set
    End Property

    Public Sub New()
      Tuning = New ScriptDB.Tuning
      Objects = New Things.Collections.Generic
      Objects.Parent = Me
    End Sub

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public Property Root As Things.Base
      Get
        Return mobjRoot
      End Get
      Set(ByVal value As Things.Base)
        mobjRoot = value
      End Set
    End Property

    '#Region "XML"

    '    Public ReadOnly Property ToXML As String Implements Interfaces.iSystemObject.ToXML 'Xml.XmlDocument 
    '      Get

    '        '     Dim sXML As String

    '        Dim sb As New System.Text.StringBuilder
    '        Dim writer As Xml.XmlTextWriter = New Xml.XmlTextWriter(New System.IO.StringWriter(sb))
    '        'Dim returnXML As New Xml.Serialization.XmlSerializer(Me.GetType)

    '        Dim returnXML As New Xml.Serialization.XmlSerializer(Me.GetType)



    '        'dtExport.WriteXml(writer)
    '        'writer.Close()

    '        'sXML = Replace(sb.ToString, "<DocumentElement>", "")
    '        'sXML = Replace(sXML, "</DocumentElement>", "")

    '        'GetXMLFromDataTable = sXML



    '        returnXML.Serialize(writer, Me)
    '        writer.Close()
    '        Return sb.ToString

    '        '(returnXML, Me)



    '        'returnXML.Serialize(objStreamWriter, Me)
    '        'objStreamWriter.Close()
    '        ' Return
    '      End Get
    '    End Property

    '#End Region

    <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
    Public Property IsSelected As Boolean
      Get
        Return mbIsSelected
      End Get
      Set(ByVal value As Boolean)
        mbIsSelected = value
      End Set
    End Property

    Public Sub ToXML(ByVal fileName As String)

      Dim serializer As New XmlSerializer(Me.GetType())
      Dim namespaces As New XmlSerializerNamespaces()
      namespaces.Add("", "")
      Using stream As New System.IO.FileStream(fileName, System.IO.FileMode.Create)
        serializer.Serialize(stream, Me, namespaces)
      End Using

      'If Not Me.Objects Is Nothing Then
      '  For Each objObject As Things.Base In Me.Objects
      '    objObject.WriteXml(writer)
      '  Next
      'End If

    End Sub

    ' Returns an object from its children
    Public Overridable Function GetObject(ByRef [Type] As Things.Type, ByRef [ID] As HCMGuid) As Things.Base
      GetObject = Me.Objects.GetObject(Type, ID)
    End Function


    '#Region "System.Xml.Serialization.IXmlSerializable"

    '    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
    '      Return Nothing
    '    End Function

    '    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
    '    End Sub

    '    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml

    '      Try

    '        'Dim serializer As XmlSerializer = New XmlSerializer(Me.GetType)
    '        'serializer.Serialize(writer, Me)

    '        'For Each oObject As Object In Me.XmlAttributes

    '        'Next


    '        'Dim pc() As PropertyInfo = GetType().GetProperties()
    '        'Dim ti As T = Nothing
    '        'For i As Int32 = 0 To Me.Items.Count - 1
    '        '  ti = Me.Item(i)
    '        '  writer.WriteStartElement(GetType(T).Name)
    '        '  For j As Int32 = 0 To pc.Length - 1
    '        '    If pc(j).CanRead And pc(j).CanWrite Then
    '        '      writer.WriteStartElement(pc(j).Name)
    '        '      Dim st As SerilalizeType = GetSerilalizeType(pc(j).PropertyType)
    '        '      If st = SerilalizeType.Complex Or _
    '        '         st = SerilalizeType.Array Or _
    '        '         st = SerilalizeType.ICollection Then
    '        '        writer.WriteRaw(SerializeObject(pc(j).GetValue(ti, Nothing)))
    '        '      Else
    '        '        writer.WriteString(pc(j).GetValue(ti, Nothing).ToString())
    '        '      End If
    '        '      writer.WriteEndElement()
    '        '    End If
    '        '  Next
    '        '  writer.WriteEndElement()
    '        'Next


    '        'Dim value_serializer As XmlSerializer = New XmlSerializer(GetType(TValue))
    '        'For Each key As TKey In Me.Keys
    '        '  writer.WriteStartElement("item")
    '        '  writer.WriteStartElement("key")
    '        '  key_serializer.Serialize(writer, key)
    '        '  writer.WriteEndElement()
    '        '  writer.WriteStartElement("value")
    '        '  Dim value As TValue = Me.Item(key)
    '        '  value_serializer.Serialize(writer, value)
    '        '  writer.WriteEndElement()
    '        '  writer.WriteEndElement()
    '        'Next key



    '        'writer.WriteAttributeString("Name", Name)
    '        'writer.WriteAttributeString("Type", Type)

    '        'If Not Me.Objects Is Nothing Then
    '        '  For Each objObject As Things.Base In Me.Objects
    '        '    objObject.WriteXml(writer)
    '        '  Next
    '        'End If

    '      Catch ex As Exception
    '        Debug.Print(ex.InnerException.ToString)
    '      End Try

    '    End Sub

    '    '<System.Xml.Serialization.XmlIgnore()> _
    '    'Public ReadOnly Property ToXML As String
    '    '  Get

    '    '    Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
    '    '    Dim xmlSerializer As System.Xml.Serialization.XmlSerializer

    '    '    'xmlSerializer = New System.Xml.Serialization.XmlSerializer(Me.GetType())

    '    '    ', System.Xml.Serialization.XmlIgnore()
    '    '    Try

    '    '      xmlSerializer = New System.Xml.Serialization.XmlSerializer(Me.GetType)

    '    '      'Using xmlStream As System.IO.MemoryStream = New System.IO.MemoryStream()
    '    '      '  xmlSerializer.Serialize(xmlStream, Me)
    '    '      '  xmlStream.Position = 0

    '    '      '  xmlDoc.Load(xmlStream)
    '    '      '  Return xmlDoc.InnerXml
    '    '      'End Using

    '    '      Using objStringWriter As System.IO.StringWriter = New System.IO.StringWriter


    '    '        'xmlDoc.Document()
    '    '        'xmlSerializer.standalone()
    '    '        '            xmlSerializer.
    '    '        xmlSerializer.Serialize(objStringWriter, Me)
    '    '        'Debug.Print(objStringWriter.ToString)
    '    '        Return objStringWriter.ToString
    '    '      End Using



    '    '    Catch ex As Exception
    '    '      Debug.Print(ex.Message)


    '    '    End Try



    '    '  End Get
    '    'End Property


    '#End Region


    ' Returns a collection of errors if necessary
    Public Overridable Function Validate() As Things.Collections.Generic

      Return New Things.Collections.Generic

    End Function



#Region "iClonable"

    'Private Function Clone(ByVal vObj As Object)

    '  Dim iClone As ICloneable

    '  If Not vObj Is Nothing Then
    '    If vObj.GetType.IsValueType OrElse vObj.GetType Is System.Type.GetType("System.String") Then
    '      Return vObj
    '    Else
    '      Dim newObject As Object = Activator.CreateInstance(vObj.GetType)
    '      If Not newObject.GetType.GetInterface("IEnumerable", True) Is Nothing AndAlso Not newObject.GetType.GetInterface("ICloneable", True) Is Nothing Then
    '        'This is a cloneable enumeration object so just clone it
    '        newObject = CType(vObj, ICloneable).Clone
    '        Return newObject
    '      Else
    '        For Each Item As PropertyInfo In newObject.GetType.GetProperties
    '          'If a property has the ICloneable interface, then call the interface clone method
    '          If Not (Item.PropertyType.GetInterface("ICloneable") Is Nothing) Then

    '            '           Item.PropertyType.Name
    '            ' Item.GetValue(

    '            If Item.CanWrite Then
    '              iClone = CType(Item.GetValue(vObj, Nothing), ICloneable)
    '              If Not iClone Is Nothing Then
    '                Item.SetValue(newObject, iClone.Clone, Nothing)
    '              End If
    '            End If
    '          Else
    '            'Otherwise just set the value
    '            If Item.CanWrite Then
    '              Item.SetValue(newObject, Clone(Item.GetValue(vObj, Nothing)), Nothing)
    '            End If
    '          End If
    '        Next
    '        Return newObject
    '      End If
    '    End If
    '  Else
    '    Return Nothing
    '  End If
    'End Function

    'Public Function Clone() As Object Implements System.ICloneable.Clone
    '  Return Clone(Me)
    'End Function

    'Private Function GetReflectedProperty(ByVal PropertyName As String, ByVal PropertyIndex As Object()) As Object
    '  Dim retVal As Object = Me.GetType().GetProperty(PropertyName).GetValue(Me, Nothing)
    '  If PropertyIndex IsNot Nothing Then
    '    retVal = retVal.[GetType]().GetProperty("Item").GetValue(retVal, PropertyIndex)
    '  End If
    '  Return retVal
    'End Function

    Private Function DeepClone(ByVal Obj As Object) As Object

      Dim objResult As Object = Nothing
      Using ms As New System.IO.MemoryStream()
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
        bf.Serialize(ms, obj)

        ms.Position = 0
        objResult = bf.Deserialize(ms)
      End Using
      Return objResult

    End Function

    Public Function DeepClone() As Object
      Return DeepClone(Me)
    End Function

#End Region

    '#Region "Dependency Stuff"

    '    Public Overridable ReadOnly Property Dependencies As Things.Collection
    '      Get
    '        Return New Things.Collection
    '      End Get
    '    End Property

    '#End Region


  End Class
End Namespace
