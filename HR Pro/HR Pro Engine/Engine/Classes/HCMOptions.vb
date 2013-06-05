Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class HCMOptions
  Implements COMInterfaces.iOptions

  Private mbRefreshObjects As Boolean
  Private mbDevelopmentMode As Boolean = False
  Private mbOverflowSafety As Boolean = True

  Public Property RefreshObjects As Boolean Implements COMInterfaces.iOptions.RefreshObjects
    Get
      Return mbRefreshObjects
    End Get
    Set(ByVal value As Boolean)
      mbRefreshObjects = value
    End Set
  End Property

  Public Property DevelopmentMode As Boolean Implements COMInterfaces.iOptions.DevelopmentMode
    Get
      Return mbDevelopmentMode
    End Get
    Set(ByVal value As Boolean)
      mbDevelopmentMode = value
    End Set
  End Property

  Public Property OverflowSafety As Boolean Implements COMInterfaces.iOptions.OverflowSafety
    Get
      Return mbOverflowSafety
    End Get
    Set(ByVal value As Boolean)
      mbOverflowSafety = value
    End Set
  End Property

End Class

