Attribute VB_Name = "modQAProFunctions"
' This file provides many useful functions for the Visual Basic integration of our API's


'*****************************************************
' Purpose:   Makes a string suitable for use by the DLLs.
' Inputs:
'   strArg: the string to be manipulated.
' Returns:   The manipulated string.
'*****************************************************

Function MakeCString(ByVal strArg As String) As String
    
    MakeCString = strArg & Chr$(0)

End Function

'*****************************************************
' Purpose:   Removes the NULL character from the end
'            of the string returned from the DLL.
' Inputs:
'   strArg: the string to be manipulated.
' Returns:   The manipulated string.
'*****************************************************

Function UnMakeCString(ByVal strArg As String) As String
    
    Dim lngNullIndex As Long

    lngNullIndex = InStr(strArg, Chr$(0))

    If lngNullIndex > 0 Then
    
        UnMakeCString = Mid$(strArg, 1, lngNullIndex - 1)
        
    Else
    
        UnMakeCString = RTrim(strArg)
        
    End If

End Function
