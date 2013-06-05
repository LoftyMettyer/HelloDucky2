Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class HCMOptions
  Implements COMInterfaces.iOptions

  Private mbOverflowSafety As Boolean = True

  Public Property OptimiseSaveProcess As Boolean Implements COMInterfaces.iOptions.OptimiseSaveProcess
  Public Property RefreshObjects As Boolean Implements COMInterfaces.iOptions.RefreshObjects
  Public Property DevelopmentMode As Boolean Implements COMInterfaces.iOptions.DevelopmentMode

  Public Property OverflowSafety As Boolean Implements COMInterfaces.iOptions.OverflowSafety
    Get
      Return mbOverflowSafety
    End Get
    Set(ByVal value As Boolean)
      mbOverflowSafety = value
    End Set
  End Property

End Class

