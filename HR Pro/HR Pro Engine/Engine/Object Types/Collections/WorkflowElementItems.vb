Imports System.ComponentModel
Imports System.Runtime.InteropServices

Namespace Things.Collections

  <DataObject(True), ClassInterface(ClassInterfaceType.None), Serializable()> _
  Public Class WorkflowElementItems
    Inherits Things.Collections.BaseCollection
    Implements ICollection_WorkflowElements

    'TODO: WANT TO REMOVE CLASS
    Public Function Element(ByVal [ID] As Integer) As WorkflowElement Implements COMInterfaces.ICollection_WorkflowElements.Element

      Dim objChild As Things.Base

      For Each objChild In MyBase.Items
        If objChild.ID = ID And objChild.Type = Type.WorkflowElementItem Then
          Return CType(objChild, WorkflowElement)
        End If
      Next

      Return Nothing

    End Function

    Public Function Elements() As BaseCollection Implements COMInterfaces.ICollection_WorkflowElements.Elements
      'TODO: Fails option strict, would error if called anyway IList => BaseCollection
      Return Nothing
      'Return Me.Items
    End Function

    Public Shadows Sub Add(ByVal objWorkflowElementItem As Things.WorkflowElementItem) Implements COMInterfaces.ICollection_WorkflowElements.Add
      Me.Items.Add(objWorkflowElementItem)
    End Sub

  End Class

End Namespace

