Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class HCMOptions
  Implements iOptions

  Private mbRefreshObjects As Boolean

  Public Property RefreshObjects As Boolean Implements Interfaces.iOptions.RefreshObjects
    Get
      Return mbRefreshObjects
    End Get
    Set(ByVal value As Boolean)
      mbRefreshObjects = value
    End Set
  End Property
End Class

