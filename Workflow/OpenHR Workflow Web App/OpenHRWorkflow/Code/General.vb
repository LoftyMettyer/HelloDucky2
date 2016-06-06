Imports Microsoft.VisualBasic
Imports System.Threading
Imports System.Drawing
Imports Microsoft.Win32
Imports System.IO
Imports System
Imports System.Globalization

Public Class General

  Private Sub New()
  End Sub

   Public Shared Global_WorkspaceUserId As String = ""

  Public Shared Function GetColour(ByVal piColour As Int32) As Color
    Try
      Return ColorTranslator.FromOle(piColour)
    Catch ex As Exception
      Return Color.White
    End Try

  End Function

  Public Shared Function GetHtmlColour(ByVal piColour As Int32) As String

    Try
      ' Create an instance of a Color structure.
      Dim myColor As Color = ColorTranslator.FromOle(piColour)

      ' Translate myColor to an HTML color.
      Dim htmlColor As String = ColorTranslator.ToHtml(myColor)

      Return (htmlColor)
    Catch ex As Exception
      Return ColorTranslator.ToHtml(Color.White)
    End Try

  End Function

  Public Shared Function ConvertLocaleDateToSql(ByVal psLocaleDateString As String) As String
    Dim dtDate As Date
    Dim iYear As Int16
    Dim iMonth As Int16
    Dim iDay As Int16

    Try
      iYear = CShort(GetDatePart(psLocaleDateString, "Y"))
      iMonth = CShort(GetDatePart(psLocaleDateString, "M"))
      iDay = CShort(GetDatePart(psLocaleDateString, "D"))

      dtDate = DateSerial(iYear, iMonth, iDay)

      Return Format(dtDate, "MM/dd/yyyy")
    Catch ex As Exception
      Return ""
    End Try

  End Function

  Public Shared Function ConvertSqlDateToLocale(ByVal psSQLDateString As String) As String
    ' Convert SQL Date string (mm/dd/yyyy) into locale short format.
    Dim dtDate As Date
    Dim iYear As Int16
    Dim iMonth As Int16
    Dim iDay As Int16

    If psSQLDateString = Nothing Then Return ""

    Try
      iYear = CShort(psSQLDateString.Substring(6, 4))
      iMonth = CShort(psSQLDateString.Substring(0, 2))
      iDay = CShort(psSQLDateString.Substring(3, 2))

      dtDate = DateSerial(iYear, iMonth, iDay)
      Return dtDate.ToString(Thread.CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern)
    Catch ex As Exception
      Return ""
    End Try

  End Function

  Public Shared Function GetDatePart(ByVal psLocaleDateString As String, ByVal psDatePart As String) As String
    Dim sLocaleDateFormat As String = Thread.CurrentThread.CurrentUICulture.DateTimeFormat.ShortDatePattern.ToUpper
    Dim sLocaleDateSep As String
    Dim iLoop As Integer
    Dim iRequiredPart As Integer = 0
    Dim sValuePart1 As String = ""
    Dim sValuePart2 As String = ""
    Dim sValuePart3 As String = ""
    Dim iPartCounter As Integer = 1
    Dim sTemp As String = ""
    Dim sResult As String = ""

    sLocaleDateSep = Replace(sLocaleDateFormat, "Y", "")
    sLocaleDateSep = Replace(sLocaleDateSep, "M", "")
    sLocaleDateSep = Left(Replace(sLocaleDateSep, "D", ""), 1)

    For iLoop = 1 To Len(psLocaleDateString)
      If Mid(psLocaleDateString, iLoop, 1) = sLocaleDateSep Then
        Select Case iPartCounter
          Case 1
            sValuePart1 = sTemp
          Case 2
            sValuePart2 = sTemp
        End Select

        iPartCounter = iPartCounter + 1
        sTemp = ""
      Else
        sTemp = sTemp & Mid(psLocaleDateString, iLoop, 1)
      End If

    Next iLoop
    sValuePart3 = sTemp

    Select Case psDatePart
      Case "Y"
        iRequiredPart = 1
        If InStr(sLocaleDateFormat, "M") < InStr(sLocaleDateFormat, "Y") Then
          iRequiredPart = iRequiredPart + 1
        End If
        If InStr(sLocaleDateFormat, "D") < InStr(sLocaleDateFormat, "Y") Then
          iRequiredPart = iRequiredPart + 1
        End If
      Case "M"
        iRequiredPart = 1
        If InStr(sLocaleDateFormat, "Y") < InStr(sLocaleDateFormat, "M") Then
          iRequiredPart = iRequiredPart + 1
        End If
        If InStr(sLocaleDateFormat, "D") < InStr(sLocaleDateFormat, "M") Then
          iRequiredPart = iRequiredPart + 1
        End If
      Case "D"
        iRequiredPart = 1
        If InStr(sLocaleDateFormat, "Y") < InStr(sLocaleDateFormat, "D") Then
          iRequiredPart = iRequiredPart + 1
        End If
        If InStr(sLocaleDateFormat, "M") < InStr(sLocaleDateFormat, "D") Then
          iRequiredPart = iRequiredPart + 1
        End If
    End Select

    Select Case iRequiredPart
      Case 1
        sResult = sValuePart1
      Case 2
        sResult = sValuePart2
      Case 3
        sResult = sValuePart3
    End Select

    GetDatePart = sResult
  End Function

  Public Shared Function BackgroundRepeat(ByVal piBackgroundImagePosition As Integer) As String
    Dim sBackgroundRepeat As String

    Try
      Select Case piBackgroundImagePosition
        Case 0
          'Top Left
          sBackgroundRepeat = "no-repeat"

        Case 1
          'Top Right
          sBackgroundRepeat = "no-repeat"

        Case 2
          'Centre
          sBackgroundRepeat = "no-repeat"

        Case 3
          'Left Tile
          sBackgroundRepeat = "repeat-y"

        Case 4
          'Right Tile
          sBackgroundRepeat = "repeat-y"

        Case 5
          'Top Tile
          sBackgroundRepeat = "repeat-x"

        Case 6
          'Bottom Tile
          sBackgroundRepeat = "repeat-x"

        Case 7
          'Tile
          sBackgroundRepeat = "repeat"

        Case Else
          'Centre
          sBackgroundRepeat = "no-repeat"
      End Select

      Return sBackgroundRepeat

    Catch ex As Exception
      Return "no-repeat"
    End Try

  End Function

  Public Shared Function BackgroundPosition(ByVal piBackgroundImagePosition As Integer) As String
    Dim sBackgroundPosition As String

    Try
      Select Case piBackgroundImagePosition
        Case 0
          'Top Left
          sBackgroundPosition = "top left"

        Case 1
          'Top Right
          sBackgroundPosition = "top right"

        Case 2
          'Centre
          sBackgroundPosition = "center"

        Case 3
          'Left Tile
          sBackgroundPosition = "left"

        Case 4
          'Right Tile
          sBackgroundPosition = "right"

        Case 5
          'Top Tile
          sBackgroundPosition = "top"

        Case 6
          'Bottom Tile
          sBackgroundPosition = "bottom"

        Case 7
          'Tile
          sBackgroundPosition = "top left"

        Case Else
          'Centre
          sBackgroundPosition = "center"
      End Select

      Return sBackgroundPosition

    Catch ex As Exception
      Return "center"
    End Try

  End Function

  Public Shared Function SplitMessage(ByVal psWholeMessage As String, _
      ByRef psPart1 As String, _
      ByRef psPart2 As String, _
      ByRef psPart3 As String) As Boolean

    Dim asText() As String
    Dim iTextIndex As Integer
    Dim sChar As String
    Dim sNextChar As String
    Dim fLiteral As Boolean
    Dim fIgnoreChar As Boolean
    Dim fDoingSlash As Boolean
    Dim sRTFCode As String
    Dim sRTFCodeToDo As String
    Dim iBracketLevel As Integer

    psPart1 = ""
    psPart2 = ""
    psPart3 = ""

    iTextIndex = 0
    ReDim asText(2)
    fDoingSlash = False
    iBracketLevel = 0
    sRTFCode = ""
    sRTFCodeToDo = ""

    Try

      Do While Len(psWholeMessage) > 0
        sChar = Mid(psWholeMessage, 1, 1)
        sNextChar = Mid(psWholeMessage, 2, 1)

        fLiteral = sChar = "\" _
          And ((sNextChar = "\") _
            Or (sNextChar = "{") _
            Or (sNextChar = "}"))

        fIgnoreChar = fDoingSlash Or (iBracketLevel > 0)

        If fDoingSlash Then
          If sChar = " " _
            Or sChar = "{" Then

            fDoingSlash = False
            sRTFCodeToDo = sRTFCode
            sRTFCode = ""
          ElseIf sChar = "\" Then
            sRTFCodeToDo = sRTFCode
            sRTFCode = sChar
          Else
            sRTFCode = sRTFCode & Trim(Replace(Replace(sChar, vbCr, ""), vbLf, ""))
          End If
        End If

        If (iBracketLevel > 0) And sChar = "}" Then
          iBracketLevel = iBracketLevel - 1
          sRTFCodeToDo = sRTFCode
          sRTFCode = ""
        End If

        If (Not fLiteral) Then
          If (sChar = "\") Then
            fDoingSlash = True
            sRTFCode = sChar
          ElseIf sChar = "{" Then
            iBracketLevel = iBracketLevel + 1
            sRTFCodeToDo = sRTFCode
            sRTFCode = ""
          ElseIf Not fIgnoreChar Then
            '        asText(iTextIndex) = asText(iTextIndex) & sChar
          End If
        Else
          '      asText(iTextIndex) = asText(iTextIndex) & sNextChar
        End If

        ' See if we need to interpret the RTF control code.
        If Len(sRTFCodeToDo) > 0 Then
          If ((sRTFCodeToDo = "\ul") And (iTextIndex = 0)) _
            Or ((sRTFCodeToDo = "\ulnone") And (iTextIndex = 1)) Then
            iTextIndex = iTextIndex + 1
            '    ElseIf (sRTFCodeToDo = "\tab") Or (sRTFCodeToDo = "\cell") Then
            '      asText(iTextIndex) = asText(iTextIndex) & vbTab
            '    ElseIf (sRTFCodeToDo = "\row") Then
            '      asText(iTextIndex) = asText(iTextIndex) & vbNewLine
            '    ElseIf (Mid(sRTFCodeToDo, 1, 2) = "\'") Then
            '      fFound = False
            '      sDeniedChar = Chr(Val("&H" & Mid(sRTFCodeToDo, 3)))
            '      For iLoop = 1 To UBound(asDeniedCharacters)
            '        If sDeniedChar = asDeniedCharacters(iLoop) Then
            '          fFound = True
            '          Exit For
            '        End If
            '      Next iLoop
            '      If Not fFound Then
            '        ReDim Preserve asDeniedCharacters(UBound(asDeniedCharacters) + 1)
            '        asDeniedCharacters(UBound(asDeniedCharacters)) = sDeniedChar
            '      End If
          End If

          sRTFCodeToDo = ""
        End If

        If (Not fLiteral) Then
          If (sChar = "\") Then
            '       fDoingSlash = True
            '        sRTFCode = sChar
          ElseIf sChar = "{" Then
            '        iBracketLevel = iBracketLevel + 1
            '        sRTFCodeToDo = sRTFCode
            '        sRTFCode = ""
          ElseIf Not fIgnoreChar Then
            asText(iTextIndex) = asText(iTextIndex) & sChar
          End If
        Else
          asText(iTextIndex) = asText(iTextIndex) & sNextChar
          fDoingSlash = False
        End If

        ' Move forward through the text (jump an extra character if we've just processed a literal.
        If fLiteral Then
          psWholeMessage = Mid(psWholeMessage, 3)
        Else
          psWholeMessage = Mid(psWholeMessage, 2)
        End If
      Loop

      psPart1 = Replace(Replace(Replace(asText(0), "<", "&lt;"), ">", "&gt;"), vbCrLf, "<BR>")
      psPart2 = Replace(Replace(Replace(asText(1), "<", "&lt;"), ">", "&gt;"), vbCrLf, "<BR>")
      psPart3 = Replace(Replace(Replace(asText(2), "<", "&lt;"), ">", "&gt;"), vbCrLf, "<BR>")

      SplitMessage = True

    Catch ex As Exception
      psPart1 = ""
      psPart2 = ""
      psPart3 = ""
      SplitMessage = False
    End Try
  End Function

  Public Shared Function ContentTypeFromExtension(ByVal psFileName As String) As String
    Dim psContentType As String
    Dim sExtension As String
    Dim rkClasses As RegistryKey

    psContentType = ""
    sExtension = ""

    Try
      ' Get the extension from the file name (including the '.').
      If psFileName.Length > 0 Then
        sExtension = Path.GetExtension(psFileName).Trim.ToLower
      End If

      If sExtension.Length > 0 Then
        Try
          ' Try to determine the file's MIME type from the registry.
          ' Use a try-catch as permissions might not allow registry access.
          rkClasses = Registry.ClassesRoot
          psContentType = rkClasses.OpenSubKey(sExtension).GetValue("Content Type").ToString

        Catch ex As Exception
          psContentType = ""
        End Try

        If psContentType.Length = 0 Then
          ' Unable to determine the file's MIME type from the registry, so check the 
          ' file extension in our own library or extensaions and MIME types.
          Select Case sExtension.Substring(1)
            Case "323"
              psContentType = "text/h323"
            Case "3dmf"
              psContentType = "x-world/x-3dmf"

            Case "abs"
              psContentType = "audio/x-mpeg"
            Case "acx"
              psContentType = "application/internet-property-stream"
            Case "ai"
              psContentType = "application/postscript"
            Case "aif", "aiff", "aifc"
              psContentType = "audio/x-aiff"
            Case "ano"
              psContentType = "application/x-annotator"
            Case "asf", "asr", "asx"
              psContentType = "video/x-ms-asf"
            Case "asc"
              psContentType = "text/plain"
            Case "asn"
              psContentType = "application/astound"
            Case "asp"
              psContentType = "application/x-asap"
            Case "au"
              psContentType = "audio/basic"
            Case "avi"
              psContentType = "video/x-msvideo"
            Case "axs"
              psContentType = "application/x-olescript"

            Case "bas"
              psContentType = "text/plain"
            Case "bcpio"
              psContentType = "application/x-bcpio"
            Case "bin"
              psContentType = "application/octet-stream"
            Case "bmp"
              psContentType = "image/bmp"
            Case "bw"
              psContentType = "image/x-sgi-bw"

            Case "c", "c++", "cc"
              psContentType = "text/plain"
            Case "cal"
              psContentType = "image/x-cals"
            Case "cat"
              psContentType = "application/vnd.ms-pkiseccat"
            Case "ccv"
              psContentType = "application/ccv"
            Case "cdf"
              psContentType = "application/x-cdf"
            Case "cer"
              psContentType = "application/x-x509-ca-cert"
            Case "cgm"
              psContentType = "image/cgm"
            Case "class"
              psContentType = "application/octet-stream"
            Case "clp"
              psContentType = "application/x-msclip"
            Case "cmx"
              psContentType = "image/x-cmx"
            Case "cod"
              psContentType = "image/cis-cod"
            Case "cpio"
              psContentType = "application/x-cpio"
            Case "crd"
              psContentType = "application/x-mscardfile"
            Case "crl"
              psContentType = "application/pkix-crl"
            Case "crt"
              psContentType = "application/x-x509-ca-cert"
            Case "csh"
              psContentType = "application/x-csh"
            Case "css"
              psContentType = "text/css"

            Case "dcr", "dxr"
              psContentType = "application/x-director"
            Case "der"
              psContentType = "application/x-x509-ca-cert"
            Case "dir"
              psContentType = "application/x-dirview"
            Case "dll"
              psContentType = "application/x-msdownload"
            Case "dms"
              psContentType = "application/octet-stream"
            Case "doc", "dot"
              psContentType = "application/msword"
            Case "docm"
              psContentType = "application/vnd.ms-word.document.macroEnabled.12"
            Case "docx"
              psContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case "dotm"
              psContentType = "application/vnd.ms-word.template.macroEnabled.12"
            Case "dotx"
              psContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
            Case "dsf"
              psContentType = "image/x-mgx-dsf"
            Case "dvi"
              psContentType = "application/x-dvi"
            Case "dwf"
              psContentType = "drawing/x-dwf"
            Case "dwg"
              psContentType = "image/vnd.dwg"
            Case "dxf"
              psContentType = "image/vnd.dxf"

            Case "eps"
              psContentType = "application/postscript"
            Case "epsi", "epsf"
              psContentType = "image/x-eps"
            Case "es"
              psContentType = "audio/echospeech"
            Case "etx"
              psContentType = "text/x-setext"
            Case "evy"
              psContentType = "application/envoy"
            Case "exe"
              psContentType = "application/octet-stream"

            Case "faxmgr"
              psContentType = "application/x-fax-manager"
            Case "faxmgrjob"
              psContentType = "application/x-fax-manager-job"
            Case "fif"
              psContentType = "application/fractals"
            Case "flr"
              psContentType = "x-world/x-vrml"
            Case "fm", "frame"
              psContentType = "application/vnd.framemaker"
            Case "frm"
              psContentType = "application/x-alpha-form"

            Case "g3f"
              psContentType = "image/g3fax"
            Case "gif"
              psContentType = "image/gif"
            Case "gtar"
              psContentType = "application/x-gtar"
            Case "gz"
              psContentType = "application/x-gzip"

            Case "h"
              psContentType = "text/plain"
            Case "hdf"
              psContentType = "application/x-hdf"
            Case "hlp"
              psContentType = "application/winhlp"
            Case "hqx"
              psContentType = "application/mac-binhex40"
            Case "hta"
              psContentType = "application/hta"
            Case "htc"
              psContentType = "text/x-component"
            Case "html", "htm"
              psContentType = "text/html"
            Case "htt"
              psContentType = "text/webviewhtml"

            Case "ice"
              psContentType = "x-conference/x-cooltalk"
            Case "icnbk"
              psContentType = "application/x-iconbook"
            Case "ico"
              psContentType = "image/x-icon"
            Case "ief"
              psContentType = "image/ief"
            Case "igs"
              psContentType = "application/iges"
            Case "iii"
              psContentType = "application/x-iphone"
            Case "ins"
              psContentType = "application/x-internet-signup"
            Case "insight"
              psContentType = "application/x-insight"
            Case "inst"
              psContentType = "application/x-install"
            Case "ipcall"
              psContentType = "application/x-inperson-call"
            Case "isp"
              psContentType = "application/x-internet-signup"
            Case "iv"
              psContentType = "graphics/x-inventor"

            Case "jfif"
              psContentType = "image/pipeg"
            Case "jpg", "jpeg", "jpe"
              psContentType = "image/jpeg"
            Case "js"
              psContentType = "text/javascript"

            Case "latex"
              psContentType = "application/x-latex"
            Case "lcc"
              psContentType = "application/fastman"
            Case "lha"
              psContentType = "application/octet-stream"
            Case "lic"
              psContentType = "application/x-enterlicense"
            Case "ls"
              psContentType = "text/javascript"
            Case "lsf", "lsx"
              psContentType = "video/x-la-asf"
            Case "lzh"
              psContentType = "application/octet-stream"

            Case "m13", "m14"
              psContentType = "application/x-msmediaview"
            Case "m3u"
              psContentType = "audio/x-mpegurl"
            Case "ma"
              psContentType = "application/mathematica"
            Case "mail"
              psContentType = "application/x-mailfolder"
            Case "man"
              psContentType = "application/x-troff-man"
            Case "mbd"
              psContentType = "application/mbedlet"
            Case "mdb"
              psContentType = "application/x-msaccess"
            Case "me"
              psContentType = "application/x-troff-me"
            Case "mht", "mhtml"
              psContentType = "message/rfc822"
            Case "mid"
              psContentType = "audio/mid"
            Case "mif"
              psContentType = "application/vnd.mif"
            Case "mil"
              psContentType = "image/x-cals"
            Case "mmid"
              psContentType = "x-music/x-midi"
            Case "mny"
              psContentType = "application/x-msmoney"
            Case "mocha"
              psContentType = "text/javascript"
            Case "mov"
              psContentType = "video/quicktime"
            Case "movie"
              psContentType = "video/x-sgi-movie"
            Case "mp2a", "mpa2"
              psContentType = "audio/x-mpeg2"
            Case "mp3", "mp2", "mpga"
              psContentType = "audio/mpeg"
            Case "mpa", "mpega"
              psContentType = "audio/x-mpeg"
            Case "mpc", "mpt", "mpx", "mpw", "mpp"
              psContentType = "application/vnd.ms-project"
            Case "mpe", "mpg", "mpeg"
              psContentType = "video/mpeg"
            Case "mpv2", "mp2v"
              psContentType = "video/mpeg2"
            Case "ms"
              psContentType = "application/x-troff-ms"
            Case "msg"
              psContentType = "application/vnd.ms-outlook"
            Case "msh"
              psContentType = "x-model/x-mesh"
            Case "msw"
              psContentType = "application/x-dos_ms_word"
            Case "mvb"
              psContentType = "application/x-msmediaview"

            Case "nc"
              psContentType = "application/x-netcdf"
            Case "nws"
              psContentType = "message/rfc822"

            Case "oda"
              psContentType = "application/oda"
            Case "odc"
              psContentType = "application/vnd.oasis.opendocument.chart"
            Case "odf"
              psContentType = "application/vnd.oasis.opendocument.formula"
            Case "odg"
              psContentType = "application/vnd.oasis.opendocument.graphics"
            Case "odi"
              psContentType = "application/vnd.oasis.opendocument.image"
            Case "odm"
              psContentType = "application/vnd.oasis.opendocument.text-master"
            Case "odp"
              psContentType = "application/vnd.oasis.opendocument.presentation"
            Case "ods"
              psContentType = "application/x-oleobject"
            Case "odt"
              psContentType = "application/vnd.oasis.opendocument.text"
            Case "opp"
              psContentType = "x-form/x-openscape"
            Case "otc"
              psContentType = "application/vnd.oasis.opendocument.chart-template"
            Case "otf"
              psContentType = "application/vnd.oasis.opendocument.formula-template"
            Case "otg"
              psContentType = "application/vnd.oasis.opendocument.graphics-template"
            Case "oti"
              psContentType = "application/vnd.oasis.opendocument.image-template"
            Case "oth"
              psContentType = "application/vnd.oasis.opendocument.text-web"
            Case "otp"
              psContentType = "application/vnd.oasis.opendocument.presentation-template"
            Case "ots"
              psContentType = "application/vnd.oasis.opendocument.spreadsheet-template"
            Case "ott"
              psContentType = "application/vnd.oasis.opendocument.text-template"

            Case "p10"
              psContentType = "application/pkcs10"
            Case "p12"
              psContentType = "application/x-pkcs12"
            Case "p3d"
              psContentType = "application/x-p3d"
            Case "p7b"
              psContentType = "application/x-pkcs7-certificates"
            Case "p7c", "p7m"
              psContentType = "application/x-pkcs7-mime"
            Case "p7r"
              psContentType = "application/x-pkcs7-certreqresp"
            Case "p7s"
              psContentType = "application/x-pkcs7-signature"
            Case "pac"
              psContentType = "application/x-ns-proxy-autoconfig"
            Case "pbm"
              psContentType = "image/x-portable-bitmap"
            Case "pcd"
              psContentType = "image/x-photo-cd"
            Case "pcn"
              psContentType = "application/x-pcn"
            Case "pdf"
              psContentType = "application/pdf"
            Case "pfx"
              psContentType = "application/x-pkcs12"
            Case "pgm"
              psContentType = "image/x-portable-graymap"
            Case "pict"
              psContentType = "image/x-pict"
            Case "pko"
              psContentType = "application/ynd.ms-pkipko"
            Case "pl"
              psContentType = "application/x-perl"
            Case "pma", "pmc", "pml", "pmr", "pmw"
              psContentType = "application/x-perfmon"
            Case "png"
              psContentType = "image/x-png"
            Case "pnm"
              psContentType = "image/x-portable-anymap"
            Case "potx"
              psContentType = "application/vnd.openxmlformats-officedocument.presentationml.template"
            Case "pp", "ppages"
              psContentType = "application/x-ppages"
            Case "ppm"
              psContentType = "image/x-portable-pixmap"
            Case "ppt", "ppz", "pps", "pot"
              psContentType = "application/vnd.ms-powerpoint"
            Case "ppsm"
              psContentType = "application/vnd.ms-powerpoint.slideshow.macroEnabled.12"
            Case "ppsx"
              psContentType = "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
            Case "pptm"
              psContentType = "application/vnd.ms-powerpoint.presentation.macroEnabled.12"
            Case "pptx"
              psContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            Case "ppz"
              psContentType = "application/mspowerpoint"
            Case "prf"
              psContentType = "application/pics-rules"
            Case "ps"
              psContentType = "application/postscript"
            Case "pub"
              psContentType = "application/x-mspublisher"

            Case "qt"
              psContentType = "video/quicktime"

            Case "ra", "ram"
              psContentType = "application/x-pn-realaudio"
            Case "rad"
              psContentType = "application/x-rad-powermedia"
            Case "ras"
              psContentType = "image/x-cmu-raster"
            Case "rgb"
              psContentType = "image/x-rgb"
            Case "rgba"
              psContentType = "image/x-sgi-rgba"
            Case "rm", "rpm"
              psContentType = "application/x-pn-realaudio-plugin"
            Case "rmi"
              psContentType = "audio/mid"
            Case "roff"
              psContentType = "application/x-troff"
            Case "rtf"
              psContentType = "application/rtf"
            Case "rtx"
              psContentType = "text/richtext"

            Case "scd"
              psContentType = "application/x-msschedule"
            Case "sct"
              psContentType = "text/scriptlet"
            Case "setpay"
              psContentType = "application/set-payment-initiation"
            Case "setreg"
              psContentType = "application/set-registration-initiation"
            Case "sgi"
              psContentType = "image/x-sgi-rgba"
            Case "sgi-lpr"
              psContentType = "application/x-sgi-lpr"
            Case "sh"
              psContentType = "application/x-sh"
            Case "shar"
              psContentType = "application/x-shar"
            Case "showcase", "slides", "sc", "sho", "show"
              psContentType = "application/x-showcase"
            Case "sit", "sea"
              psContentType = "application/x-stuffit"
            Case "skp"
              psContentType = "application/vnd.koan"
            Case "snd"
              psContentType = "audio/basic"
            Case "spc"
              psContentType = "application/x-pkcs7-certificates"
            Case "spl"
              psContentType = "application/futuresplash"
            Case "src", "wsrc"
              psContentType = "application/x-wais-source"
            Case "sst"
              psContentType = "application/vnd.ms-pkicertstore"
            Case "stl"
              psContentType = "application/vnd.ms-pkistl"
            Case "stm"
              psContentType = "text/html"
            Case "sv4cpio"
              psContentType = "application/x-sv4cpio"
            Case "sv4crc"
              psContentType = "application/x-sv4crc"
            Case "svd"
              psContentType = "application/vnd.svd"
            Case "svf"
              psContentType = "image/vnd.svf"
            Case "svg"
              psContentType = "image/svg+xml"
            Case "svr"
              psContentType = "x-world/x-svr"

            Case "t", "tr"
              psContentType = "application/x-troff"
            Case "talk"
              psContentType = "text/x-speech"
            Case "tar"
              psContentType = "application/x-tar"
            Case "tardist"
              psContentType = "application/x-tardist"
            Case "tcl"
              psContentType = "application/x-tcl"
            Case "tex"
              psContentType = "application/x-tex"
            Case "texinfo", "texi"
              psContentType = "application/x-texinfo"
            Case "tgz"
              psContentType = "application/x-compressed"
            Case "tif", "tiff"
              psContentType = "image/tiff"
            Case "trm"
              psContentType = "application/x-msterminal"
            Case "tsv"
              psContentType = "text/tab-separated-values"
            Case "txt"
              psContentType = "text/plain"

            Case "uls"
              psContentType = "text/iuls"
            Case "ustar"
              psContentType = "application/x-ustar"
            Case "uu"
              psContentType = "application/octet-stream"

            Case "v5d"
              psContentType = "application/vis5d"
            Case "vb"
              psContentType = "application/x-cosmobuilder"
            Case "vcf"
              psContentType = "text/x-vcard"
            Case "vdo"
              psContentType = "video/vdo"
            Case "viv"
              psContentType = "video/vnd.vivo"
            Case "vox"
              psContentType = "audio/voxware"
            Case "vrml"
              psContentType = "x-world/x-vrml"
            Case "vrw"
              psContentType = "x-world/x-vream"
            Case "vts"
              psContentType = "workbook/formulaone"

            Case "wav"
              psContentType = "audio/x-wav"
            Case "wb"
              psContentType = "application/x-inpview"
            Case "wba"
              psContentType = "application/x-webbasic"
            Case "wcm", "wdb", "wks"
              psContentType = "application/vnd.ms-works"
            Case "wfx"
              psContentType = "x-script/x-wfxclient"
            Case "wi"
              psContentType = "image/wavelet"
            Case "wkz"
              psContentType = "application/x-wingz"
            Case "wmf"
              psContentType = "application/x-msmetafile"
            Case "wps"
              psContentType = "application/vnd.ms-works"
            Case "wri"
              psContentType = "application/x-mswrite"
            Case "wrl", "wrz"
              psContentType = "x-world/x-vrml"
            Case "wvr"
              psContentType = "x-world/x-wvr"

            Case "xaf"
              psContentType = "x-world/x-vrml"
            Case "xbm"
              psContentType = "image/x-xbitmap"
            Case "xl"
              psContentType = "application/x-dos_ms_excel"
            Case "xls", "xla", "xlc", "xlt", "xll", "xlm", "xlw"
              psContentType = "application/vnd.ms-excel"
            Case "xlsb"
              psContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12"
            Case "xlsm"
              psContentType = "application/vnd.ms-excel.sheet.macroEnabled.12"
            Case "xlsx"
              psContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Case "xltx"
              psContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.template"
            Case "xof"
              psContentType = "x-world/x-vrml"
            Case "xpm"
              psContentType = "image/x-xpixmap"
            Case "xps"
              psContentType = "application/vnd.ms-xpsdocument"
            Case "xwd"
              psContentType = "image/x-xwindowdump"

            Case "z"
              psContentType = "application/x-compress"
            Case "zip"
              psContentType = "application/zip"
            Case "ztardist"
              psContentType = "application/x-ztardist"

          End Select
        End If
      End If

    Catch ex As Exception
      psContentType = ""

    Finally
      If psContentType.Length = 0 Then
        ' Unable to determine the file's MIME type, so use a default.
        psContentType = "text/html"
      End If
    End Try

    Return psContentType
  End Function

  Friend Shared Function DecryptFromQueryString(value As String) As WorkflowUrl

    Dim url As New WorkflowUrl

    Try
      'Try the latest encryption method
      'Set the culture to English(GB) to ensure the decryption works OK. Fault HRPRO-1404
      Dim currentCulture = Thread.CurrentThread.CurrentCulture

      Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-GB")
      Thread.CurrentThread.CurrentUICulture = CultureInfo.CreateSpecificCulture("en-GB")

      Dim crypt As New Crypt
      value = crypt.DecompactString(value)
      value = crypt.DecryptString(value, "", True)

      'Reset the culture to be the one used by the client. Fault HRPRO-1404
      Thread.CurrentThread.CurrentCulture = currentCulture
      Thread.CurrentThread.CurrentUICulture = currentCulture

      'Extract the required parameters from the decrypted queryString.
      Dim values = value.Split(vbTab(0))

      url.InstanceId = CInt(values(0))
      url.ElementId = CInt(values(1))
      url.User = values(2)
      url.Password = values(3)
      url.Server = values(4)
      url.Database = values(5)
      If values.Count > 6 Then url.UserName = values(6)

    Catch ex As Exception
      'Try the older encryption method
      Try
        Dim crypt As New Crypt
        value = crypt.ProcessDecryptString(value)
        value = crypt.DecryptString(value, "", False)

        Dim values = value.Split(vbTab(0))

        If url.InstanceId = 0 Then url.InstanceId = CInt(values(0))
        If url.ElementId = 0 Then url.ElementId = CInt(values(1))
        url.User = values(2)
        url.Password = values(3)
        url.Server = values(4)
        url.Database = values(5)
      Catch exx As Exception
        Throw New Exception("Invalid workflow url")
      End Try
    End Try

    Return url

  End Function 

End Class
