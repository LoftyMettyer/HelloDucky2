Option Strict On
Option Explicit On

Namespace Classes
    Public Class SessionVariableState
        Public Key As Integer
        Public action As String
        Public tableID As Integer
        Public viewID As Integer
        Public screenID As Integer
        Public orderID As Integer
        Public parentTableID As Integer
        Public parentRecordID As Integer
        Public realSource As String
        Public lineage As String
        Public locateValue As Integer
        Public firstRecPos As Integer
        Public currentRecCount As Integer
        Public optionRecordID As Integer
        Public optionAction As OptionActionType
        Public recordID As Integer
        Public selectSQL As String
    End Class

End Namespace