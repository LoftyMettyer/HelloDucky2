Option Strict On

Imports Microsoft.VisualBasic
Imports System

Public Class Utilities

  Public Shared Function PointToPixel(pointSize As Integer) As Integer

    If pointSize = 0 Then
      Throw New Exception("zero pointSize specified")
    End If

    Return CInt(pointSize * 1.3333)
  End Function

  Public Shared Function PointToPixelFontUnit(pointSize As Integer) As FontUnit
    Return New FontUnit(PointToPixel(pointSize), UnitType.Pixel)
  End Function
  
  '****************************************************************
  ' NullSafeString
  '****************************************************************
  Public Shared Function NullSafeString(ByVal arg As Object, _
  Optional ByVal returnIfEmpty As String = "") As String

    Dim returnValue As String

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CStr(arg)
      Catch
        returnValue = returnIfEmpty
      End Try

    End If

    Return returnValue

  End Function

  '****************************************************************
  ' NullSafeInteger
  '****************************************************************
  Public Shared Function NullSafeInteger(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Integer = 0) As Integer

    Dim returnValue As Integer

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CInt(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  '   NullSafeDouble
  '****************************************************************
  Public Shared Function NullSafeDouble(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Double = 0) As Double

    Dim returnValue As Double

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CDbl(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  '   NullSafeSingle
  '****************************************************************
  Public Shared Function NullSafeSingle(ByVal arg As Object, _
    Optional ByVal returnIfEmpty As Single = 0) As Single

    Dim returnValue As Single

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CSng(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  ' NullSafeBoolean
  '****************************************************************
  Public Shared Function NullSafeBoolean(ByVal arg As Object) As Boolean

    Dim returnValue As Boolean

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = False
    Else
      Try
        returnValue = CBool(arg)
      Catch
        returnValue = False
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  ' isMobileBrowser
  '****************************************************************
  Public Shared Function isMobileBrowser() As Boolean
    'GETS THE CURRENT USER CONTEXT
    Dim context As HttpContext = HttpContext.Current

    'FIRST TRY BUILT IN ASP.NT CHECK
    If context.Request.Browser.IsMobileDevice Then
      Return True
    End If
    'THEN TRY CHECKING FOR THE HTTP_X_WAP_PROFILE HEADER
    If context.Request.ServerVariables("HTTP_X_WAP_PROFILE") IsNot Nothing Then
      Return True
    End If
    'THEN TRY CHECKING THAT HTTP_ACCEPT EXISTS AND CONTAINS WAP
    If context.Request.ServerVariables("HTTP_ACCEPT") IsNot Nothing AndAlso context.Request.ServerVariables("HTTP_ACCEPT").ToLower().Contains("wap") Then
      Return True
    End If
    'AND FINALLY CHECK THE HTTP_USER_AGENT 
    'HEADER VARIABLE FOR ANY ONE OF THE FOLLOWING
    If context.Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
      'Create a list of all mobile types
      Dim mobiles As String() = New String() {"midp", "j2me", "avant", "docomo", "novarra", "palmos", _
"palmsource", "240x320", "opwv", "chtml", "pda", "windows ce", _
"mmp/", "blackberry", "mib/", "symbian", "wireless", "nokia", _
"hand", "mobi", "phone", "cdm", "up.b", "audio", _
"SIE-", "SEC-", "samsung", "HTC", "mot-", "mitsu", _
"sagem", "sony", "alcatel", "lg", "eric", "vx", _
"philips", "mmm", "xx", "panasonic", "sharp", _
"wap", "sch", "rover", "pocket", "benq", "java", _
"pt", "pg", "vox", "amoi", "bird", "compal", _
"kg", "voda", "sany", "kdd", "dbt", "sendo", _
"sgh", "gradi", "jb", "dddi", "moto", "iphone"}

      'Loop through each item in the list created above 
      'and check if the header contains that text
      For Each s As String In mobiles
        If context.Request.ServerVariables("HTTP_USER_AGENT").ToLower().Contains(s.ToLower()) Then
          Return True
        End If
      Next

    End If

    Return False
  End Function





End Class
