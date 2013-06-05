Imports System.ComponentModel
Imports System.Runtime.InteropServices

Namespace Things.Collections

  <DataObject(True), ClassInterface(ClassInterfaceType.None), Serializable()> _
  Public Class WorkflowElementItems
    Inherits Things.Collections.BaseCollection
    Implements iCollection_WorkflowElements

        Public Function Element(ByRef [ID] As HCMGuid) As WorkflowElement Implements COMInterfaces.iCollection_WorkflowElements.Element

            Dim objChild As Things.Base

            For Each objChild In MyBase.Items
                If objChild.ID = ID And objChild.Type = Type.WorkflowElementItem Then
                    Return objChild
                End If
            Next

            Return Nothing

        End Function

        Public Function Elements() As BaseCollection Implements COMInterfaces.iCollection_WorkflowElements.Elements
            Return Me.Items
        End Function

    Public Shadows Sub Add(ByRef objWorkflowElementItem As Things.WorkflowElementItem) Implements COMInterfaces.iCollection_WorkflowElements.Add
      Me.Items.Add(objWorkflowElementItem)
    End Sub

  End Class

End Namespace

