Imports System.Xml.Serialization
Imports System.Runtime.InteropServices

<Serializable(), ClassInterface(ClassInterfaceType.None)>
Public MustInherit Class Base
  Implements IObject

  Public Property Id As Integer
  Public Overridable Property Name As String Implements IObject.Name
  Public Property Description As String
  Public Property SchemaName As String
  Public Property State As DataRowState
  Public Property Tuning As New ScriptDB.Tuning

  Public Overridable ReadOnly Property PhysicalName As String Implements IObject.PhysicalName
    Get
      Return Name
    End Get
  End Property

End Class
