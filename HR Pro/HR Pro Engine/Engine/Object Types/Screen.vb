Namespace Things
  Public Class Screen
    Inherits Things.Base

    <System.Xml.Serialization.XmlIgnore()> _
        Public Table As Things.Table

    'Public Overrides Function Commit() As Boolean
    'End Function

        Public Overrides ReadOnly Property Type As Things.Enums.Type
            Get
                Return Things.Type.Screen
            End Get
        End Property
  End Class

End Namespace