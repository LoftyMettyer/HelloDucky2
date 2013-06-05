Namespace Things
  Public Class TableOrder
    Inherits Things.Base

    Public Overrides ReadOnly Property Type As Enums.Type
      Get
        Return Enums.Type.TableOrder
      End Get
    End Property
  
  <System.Xml.Serialization.XmlIgnore(), System.ComponentModel.Browsable(False)> _
  Public ReadOnly Property FindWindowsColumns()
    Get
        Return Me.Objects(Things.Type.TableOrderItem)
    End Get
  End Property

  End Class

End Namespace
