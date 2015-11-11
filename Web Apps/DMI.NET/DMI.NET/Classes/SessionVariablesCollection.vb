Option Strict On
Option Explicit On

Namespace Classes

    Public Class SessionVariablesCollection

        Private variables As List(Of SessionVariableState) = New List(Of SessionVariableState)

        Friend Sub RestoreState(windowId As Integer)

            Dim session = HttpContext.Current.Session

            If variables.Exists(Function(o) o.Key = windowId) Then
                Dim requestedSet = variables.Where(Function(o) o.Key = windowId).First

                session("action") = requestedSet.action
                session("currentRecCount") = requestedSet.currentRecCount
                session("firstRecPos") = requestedSet.firstRecPos
                session("lineage") = requestedSet.lineage
                session("locateValue") = requestedSet.locateValue
                session("optionAction") = requestedSet.optionAction
                session("optionRecordID") = requestedSet.optionRecordID
                session("orderID") = requestedSet.orderID
                session("parentRecordID") = requestedSet.parentRecordID
                session("parentTableID") = requestedSet.parentTableID
                session("realSource") = requestedSet.realSource
                session("recordID") = requestedSet.recordID
                session("screenID") = requestedSet.screenID
                session("selectSQL") = requestedSet.selectSQL
                session("tableID") = requestedSet.tableID
                session("viewID") = requestedSet.viewID
            End If

        End Sub

        Friend Sub UpdateState(windowId As Integer)

            Dim session = HttpContext.Current.Session
            Dim requestedSet As SessionVariableState

            If variables.Exists(Function(o) o.Key = windowId) Then
                requestedSet = variables.Where(Function(o) o.Key = windowId).First
            Else
                requestedSet = New SessionVariableState With {.Key = windowId}
                variables.Add(requestedSet)
            End If

            requestedSet.action = session("action").ToString
            requestedSet.currentRecCount = CInt(session("currentRecCount"))
            requestedSet.firstRecPos = CInt(session("firstRecPos"))
            requestedSet.lineage = session("lineage").ToString
            requestedSet.locateValue = CInt(session("locateValue"))
            requestedSet.optionAction = CType(session("optionAction"), OptionActionType)
            requestedSet.optionRecordID = CInt(session("optionRecordID"))
            requestedSet.orderID = CInt(session("orderID"))
            requestedSet.parentRecordID = CInt(session("parentRecordID"))
            requestedSet.parentTableID = CInt(session("parentTableID"))
            requestedSet.realSource = session("realSource").ToString
            requestedSet.recordID = CInt(session("recordID"))
            requestedSet.screenID = CInt(session("screenID"))
            requestedSet.selectSQL = session("selectSQL").ToString
            requestedSet.tableID = CInt(session("tableID"))
            requestedSet.viewID = CInt(session("viewID"))

        End Sub


    End Class
End Namespace