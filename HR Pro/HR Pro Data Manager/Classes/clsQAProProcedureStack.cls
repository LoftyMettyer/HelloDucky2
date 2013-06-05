VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQAProProcedureStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mobjTopProc As clsQAProProcedure

Public Function Top() As clsQAProProcedure
    ' This returns a reference to the top
    ' procedure so a caller can walk the stack
    Set Top = mobjTopProc
End Function

Private Function StackEmpty() As Boolean
    ' This makes sure the stack is not empty
    ' by checking to see if the top proc
    ' pointer is valid
    StackEmpty = (mobjTopProc Is Nothing)
End Function

Public Function EnterProc(Name As String, Optional Module As Variant) As clsQAProProcedure
 
    ' This pushes a new procedure onto the stack
    
    Dim objProc As New clsQAProProcedure
    
    ' Set the procedure's name and module properties
    objProc.Name = Name
    objProc.Module = Module
    
    ' Make its NextProc property point to
    ' the one currently at the top of the stack
    Set objProc.NextProc = mobjTopProc
    
    ' Make the new procedure the one at the top
    Set mobjTopProc = objProc
    
    ' Return a reference to the new proc
    Set EnterProc = mobjTopProc
End Function

Public Function ExitProc(Name As String) As Boolean
    ' This pops a procedure off the stack--
    ' To enforce FILO behavior we check the
    ' name passed in against that of the top
    ' procedure
    
    ' Make sure the procedure stack is not empty
    If Not StackEmpty() Then
    
        ' If the name matches, pop the proc
        ' by making the next proc the top one--
        ' this destroys the pointer to the
        ' proc currently on top and it goes away
        If mobjTopProc.Name = Name Then
            Set mobjTopProc = mobjTopProc.NextProc
            ExitProc = True
        Else
            MsgBox "Error. Trying to pop wrong procedure. " & _
             "You passed '" & Name & "'. " & _
             "Current procedure is '" & _
             mobjTopProc.Name & "'.", vbCritical
            End
        End If
    End If
End Function

