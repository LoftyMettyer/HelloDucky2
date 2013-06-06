Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)>
Public Class HCMOptions
  Implements COMInterfaces.IOptions

  Public Sub New()
    OverflowSafety = True
  End Sub

  Public Property OptimiseSaveProcess As Boolean Implements COMInterfaces.IOptions.OptimiseSaveProcess
  Public Property RefreshObjects As Boolean Implements COMInterfaces.IOptions.RefreshObjects
  Public Property DevelopmentMode As Boolean Implements COMInterfaces.IOptions.DevelopmentMode
  Public Property OverflowSafety As Boolean Implements COMInterfaces.IOptions.OverflowSafety

End Class

