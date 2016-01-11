Option Strict On

Imports System
Imports System.Reflection

Public Class Utilities

  Public Shared Function ToPoint(storedSize As Integer) As Integer
    ' PBG20120419 Fault HRPRO-2157 revert to point sizing
    Return storedSize
    Return CInt(storedSize * 1.3333)
  End Function

  Public Shared Function ToPointFontUnit(storedSize As Integer) As FontUnit
    ' PBG20120419 Fault HRPRO-2157 revert to point sizing
    Return New FontUnit(storedSize)
    'Return New FontUnit(ToPoint(storedSize), UnitType.Pixel)
  End Function

  '****************************************************************
  ' NullSafeString
  '****************************************************************
  Public Shared Function NullSafeString(ByVal arg As Object, Optional ByVal returnIfEmpty As String = "") As String

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
  Public Shared Function NullSafeInteger(ByVal arg As Object, Optional ByVal returnIfEmpty As Integer = 0) As Integer

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

  Public Shared Function NullSafeShort(ByVal arg As Object, Optional ByVal returnIfEmpty As Short = 0) As Short

    Dim returnValue As Short

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) _
        OrElse (arg Is String.Empty) Then
      returnValue = returnIfEmpty
    Else
      Try
        returnValue = CShort(arg)
      Catch
        returnValue = returnIfEmpty
      End Try
    End If

    Return returnValue

  End Function

  '****************************************************************
  '   NullSafeSingle
  '****************************************************************
  Public Shared Function NullSafeSingle(ByVal arg As Object, Optional ByVal returnIfEmpty As Single = 0) As Single

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

    If (arg Is DBNull.Value) OrElse (arg Is Nothing) OrElse (arg Is String.Empty) Then
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
  Public Shared Function IsMobileBrowser() As Boolean

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
"sgh", "gradi", "jb", "dddi", "moto", "iphone", "fennec"}

      'Loop through each item in the list created above 
      'and check if the header contains that text
      For Each s As String In mobiles
        If context.Request.ServerVariables("HTTP_USER_AGENT").ToLower().Contains(s) Then
          Return True
        End If
      Next

    End If

    Return False
  End Function

  Public Shared Function BrowserRequiresOverflowScrollFix() As Boolean

    'Earlier android browsers dont support scrolling on overflowed divs
    'So we have to include a fix where neeeded
    If IsAndroidBrowser() Then

      Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower
      If userAgent Like "*android 2.*" Then Return True
      If userAgent Like "*android 3.*" Then Return True
    End If

    Return False

  End Function

  Public Shared Function IsAndroidBrowser() As Boolean

    Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower

    If userAgent.Contains("android") AndAlso userAgent.Contains("applewebkit") Then
      Return True
    End If
    Return False
  End Function

  Public Shared Function IsMacSafari() As Boolean

    Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower

    If userAgent.Contains("macintosh") Then
      Return True
    End If
    Return False
  End Function

  Public Shared Function IsTablet() As Boolean

    Dim ua As String = HttpContext.Current.Request.UserAgent.ToUpper

    If ua.Contains("IPAD") Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Shared Function GetBrowserFamily() As String

    Dim ua As String = HttpContext.Current.Request.UserAgent.ToUpper

    If ua.Contains("MSIE") Then
      Return "MSIE"
    ElseIf ua.Contains("IPHONE") OrElse ua.Contains("IPAD") Then
      Return "IOS"
    ElseIf ua.Contains("ANDROID") Then
      Return "ANDROID"
    ElseIf ua.Contains("BLACKBERRY") Then
      Return "BLACKBERRY"
    Else
      Return "UNKNOWN"
    End If
  End Function

  Public Shared Function WebSiteName(Optional contentPageName As String = "") As String

    Const ASSEMBLYNAME As String = "OPENHRWORKFLOW"
    Const DEFAULTTITLE As String = "OpenHR"
    Dim sWebSiteVersion As String
    Dim sTitle As String

    sTitle = DEFAULTTITLE

    If contentPageName.Length > 0 Then
      sTitle &= " " & contentPageName
    End If

    Try
      Dim sAssemblyName = Assembly.GetExecutingAssembly.GetName.Name.ToUpper

      sWebSiteVersion = Assembly.GetExecutingAssembly.GetName.Version.Major.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString _
       & "." & Assembly.GetExecutingAssembly.GetName.Version.Build.ToString

      If sAssemblyName = ASSEMBLYNAME Then
        ' Compiled version of the web site
        If sWebSiteVersion.Length = 0 Then
          sTitle &= " (unknown version)"
        Else
          sTitle &= " - v" & sWebSiteVersion
        End If
      Else
        ' Development version of the web site
        sTitle &= " (development)"
      End If
    Catch ex As Exception
      sTitle = DEFAULTTITLE
    End Try

    Return sTitle

  End Function


End Class
