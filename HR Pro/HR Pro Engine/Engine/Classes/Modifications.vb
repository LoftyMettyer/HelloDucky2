Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class Modifications
  Implements IModifications

  Dim bStructureChanged As Boolean

  Public Property ExpressionChanged As Boolean Implements COMInterfaces.IModifications.ExpressionChanged
  Public Property ModuleSetupChanged As Boolean Implements COMInterfaces.IModifications.ModuleSetupChanged
  Public Property ScreenChanged As Boolean Implements COMInterfaces.IModifications.ScreenChanged
  Public Property WorkflowChanged As Boolean Implements COMInterfaces.IModifications.WorkflowChanged
  Public Property PlatformChanged As Boolean Implements COMInterfaces.IModifications.PlatformChanged

  Public Property StructureChanged As Boolean Implements COMInterfaces.IModifications.StructureChanged
    Get
      If Globals.Options.OptimiseSaveProcess Then
        Return bStructureChanged
      Else
        Return True
      End If

    End Get
    Set(ByVal value As Boolean)
      bStructureChanged = value
    End Set
  End Property

End Class
