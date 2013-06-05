Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class Modifications
  Implements iModifications

  Dim bStructureChanged As Boolean

  Public Property ExpressionChanged As Boolean Implements COMInterfaces.iModifications.ExpressionChanged
  Public Property ModuleSetupChanged As Boolean Implements COMInterfaces.iModifications.ModuleSetupChanged
  Public Property ScreenChanged As Boolean Implements COMInterfaces.iModifications.ScreenChanged
  Public Property WorkflowChanged As Boolean Implements COMInterfaces.iModifications.WorkflowChanged
  Public Property PlatformChanged As Boolean Implements COMInterfaces.iModifications.PlatformChanged

  Public Property StructureChanged As Boolean Implements COMInterfaces.iModifications.StructureChanged
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
