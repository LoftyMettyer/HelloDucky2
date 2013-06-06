Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)>
Public Class HCMOptions
  Implements COMInterfaces.IOptions

  Public Sub New()
    OverflowSafety = True
  End Sub

  Public Property OptimiseSaveProcess As Boolean Implements IOptions.OptimiseSaveProcess
  Public Property RefreshObjects As Boolean Implements IOptions.RefreshObjects
  Public Property DevelopmentMode As Boolean Implements IOptions.DevelopmentMode
  Public Property OverflowSafety As Boolean Implements IOptions.OverflowSafety

End Class

