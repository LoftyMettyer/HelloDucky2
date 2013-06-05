Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class HCMOptions
  Implements COMInterfaces.IOptions

  Private mbOverflowSafety As Boolean = True

  Public Property OptimiseSaveProcess As Boolean Implements COMInterfaces.IOptions.OptimiseSaveProcess
  Public Property RefreshObjects As Boolean Implements COMInterfaces.IOptions.RefreshObjects
  Public Property DevelopmentMode As Boolean Implements COMInterfaces.IOptions.DevelopmentMode

  Public Property OverflowSafety As Boolean Implements COMInterfaces.IOptions.OverflowSafety
    Get
      Return mbOverflowSafety
    End Get
    Set(ByVal value As Boolean)
      mbOverflowSafety = value
    End Set
  End Property

End Class

