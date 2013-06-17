Imports System
Imports System.Reflection

Public Module Utilities

   Public Function ToPoint(storedSize As Integer) As Integer
      ' PBG20120419 Fault HRPRO-2157 revert to point sizing
      Return storedSize
      Return CInt(storedSize * 1.3333)
   End Function

   Public Function ToPointFontUnit(storedSize As Integer) As FontUnit
      ' PBG20120419 Fault HRPRO-2157 revert to point sizing
      Return New FontUnit(storedSize)
      'Return New FontUnit(ToPoint(storedSize), UnitType.Pixel)
   End Function

   '****************************************************************
   ' NullSafeString
   '****************************************************************
   Public Function NullSafeString(ByVal arg As Object, Optional ByVal returnIfEmpty As String = "") As String

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
   Public Function NullSafeInteger(ByVal arg As Object, Optional ByVal returnIfEmpty As Integer = 0) As Integer

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

   Public Function NullSafeShort(ByVal arg As Object, Optional ByVal returnIfEmpty As Short = 0) As Short

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
   Public Function NullSafeSingle(ByVal arg As Object, Optional ByVal returnIfEmpty As Single = 0) As Single

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
   Public Function NullSafeBoolean(ByVal arg As Object) As Boolean

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
   Public Function IsMobileBrowser() As Boolean

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

   Public Function BrowserRequiresOverflowScrollFix() As Boolean

      'Earlier android browsers dont support scrolling on overflowed divs
      'So we have to include a fix where neeeded
      If IsAndroidBrowser() Then

         Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower
         If userAgent Like "*android 2.*" Then Return True
         If userAgent Like "*android 3.*" Then Return True
      End If

      Return False

   End Function

   Public Function IsAndroidBrowser() As Boolean

      Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower

      If userAgent.Contains("android") AndAlso userAgent.Contains("applewebkit") Then
         Return True
      End If
      Return False
   End Function

   Public Function IsMacSafari() As Boolean

      Dim userAgent = HttpContext.Current.Request.ServerVariables("HTTP_USER_AGENT").ToLower

      If userAgent.Contains("macintosh") Then
         Return True
      End If
      Return False
   End Function

   Public Function IsTablet() As Boolean

      Dim ua As String = HttpContext.Current.Request.UserAgent.ToUpper

      If ua.Contains("IPAD") Then
         Return True
      Else
         Return False
      End If
   End Function

   Public Function GetBrowserFamily() As String

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

   Public Function GetPageTitle(pageName As String) As String

      With Assembly.GetExecutingAssembly.GetName.Version
         Return String.Format("OpenHR {0} - v{1}.{2}.{3}", pageName, .Major, .Minor, .Build)
      End With

   End Function

End Module
