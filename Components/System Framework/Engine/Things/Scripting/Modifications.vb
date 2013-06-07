Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)>
Public Class Modifications
  Implements IModifications

  Dim _structureChanged As Boolean

  Public Property ExpressionChanged As Boolean Implements IModifications.ExpressionChanged
  Public Property ModuleSetupChanged As Boolean Implements IModifications.ModuleSetupChanged
  Public Property ScreenChanged As Boolean Implements IModifications.ScreenChanged
  Public Property WorkflowChanged As Boolean Implements IModifications.WorkflowChanged
  Public Property PlatformChanged As Boolean Implements IModifications.PlatformChanged

  Public Property StructureChanged As Boolean Implements IModifications.StructureChanged
    Get
      If Options.OptimiseSaveProcess Then
        Return _structureChanged
      Else
        Return True
      End If

    End Get
    Set(ByVal value As Boolean)
      _structureChanged = value
    End Set
  End Property

End Class
