Imports System.Xml.Serialization
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Things

  <Serializable(), ClassInterface(ClassInterfaceType.None)>
  Public MustInherit Class Base
    Implements COMInterfaces.IObject

    Property ID As Integer
    Public Overridable Property Name As String Implements IObject.Name
    Public Property Description As String
    Public Property SchemaName As String
    Public Property Encrypted As Boolean
    Public Property State As System.Data.DataRowState
    Public Overridable Property SubType As Type
    Public Property Tuning As New ScriptDB.Tuning

    Public Overridable ReadOnly Property PhysicalName As String Implements IObject.PhysicalName
      Get
        Return Name
      End Get
    End Property

#Region "IClonable & XML"

    Public Sub ToXML(ByVal fileName As String)

      Dim serializer As New XmlSerializer(Me.GetType())
      Dim namespaces As New XmlSerializerNamespaces()
      namespaces.Add("", "")
      Using stream As New System.IO.FileStream(fileName, System.IO.FileMode.Create)
        serializer.Serialize(stream, Me, namespaces)
      End Using

    End Sub

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

    'Private Function DeepClone(ByVal Obj As Object) As Object

    '  Dim objResult As Object = Nothing
    '  Using ms As New System.IO.MemoryStream()
    '    Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
    '    bf.Serialize(ms, obj)

    '    ms.Position = 0
    '    objResult = bf.Deserialize(ms)
    '  End Using
    '  Return objResult

    'End Function

    'Public Function DeepClone() As Object
    '  Return DeepClone(Me)
    'End Function

#End Region

  End Class
End Namespace
