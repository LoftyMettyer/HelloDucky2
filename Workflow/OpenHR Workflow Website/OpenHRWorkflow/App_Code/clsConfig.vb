Option Strict On

Imports Microsoft.VisualBasic
Imports System.Xml
Imports Utilities
Imports System

Public Class Config
  Private msThemeName As String
  Private msThemeHex As String
  Private msThemeFore As String
  Private miMessageFontSize As Int32
  Private miValidationMessageFontSize As Int32
  Private miSubmissionTimeout As Int32
  Private msOLEFolder_Server As String
  Private msOLEFolder_Local As String
  Private msPhotographFolder As String
  Private miLookupRowsRange As Int32
  ' Mobile Stuff
  Private msServerName As String
  Private msDatabaseName As String
  Private msLogin As String
  Private msPassword As String
  Private msWorkflowURL As String


  Public Sub Initialise(ByVal psConfigFile As String)
    Dim sHexFileName As String
    Dim xmlReader As XmlTextReader

    Try
      msThemeName = System.Configuration.ConfigurationManager.AppSettings("Theme").Trim.ToUpper
      miMessageFontSize = NullSafeInteger(System.Configuration.ConfigurationManager.AppSettings("MessageFontSize"))
      miValidationMessageFontSize = NullSafeInteger(System.Configuration.ConfigurationManager.AppSettings("ValidationMessageFontSize"))
      msOLEFolder_Server = System.Configuration.ConfigurationManager.AppSettings("OLEFolder_Server").Trim
      msOLEFolder_Local = System.Configuration.ConfigurationManager.AppSettings("OLEFolder_Local").Trim
      msPhotographFolder = System.Configuration.ConfigurationManager.AppSettings("PhotographFolder").Trim
      miSubmissionTimeout = NullSafeInteger(System.Configuration.ConfigurationManager.AppSettings("SubmissionTimeout"))
      miLookupRowsRange = NullSafeInteger(System.Configuration.ConfigurationManager.AppSettings("LookupRowsRange"))

      ' Read the Hex and Foreground values for the defined theme.
      Try
        sHexFileName = psConfigFile
        xmlReader = New XmlTextReader(sHexFileName)

        Do While (xmlReader.ReadToFollowing("theme"))
          If xmlReader.ReadToFollowing("name") Then
            If (xmlReader.Read()) Then
              If (xmlReader.Value.Trim.ToUpper = msThemeName) Then
                If xmlReader.ReadToFollowing("hex") Then
                  If (xmlReader.Read()) Then
                    msThemeHex = xmlReader.Value.Trim.ToUpper

                    If xmlReader.ReadToFollowing("forecolour") Then
                      If (xmlReader.Read()) Then
                        msThemeFore = xmlReader.Value.Trim.ToUpper
                      End If
                    End If
                  End If
                End If

                Exit Do
              End If
            End If
          End If
        Loop
        xmlReader.Close()
      Catch ex As Exception

      End Try
    Catch ex As Exception

    End Try

  End Sub

  Public Sub Mob_Initialise()
    Try
      msServerName = System.Configuration.ConfigurationManager.AppSettings("Server").Trim
      msDatabaseName = System.Configuration.ConfigurationManager.AppSettings("Database").Trim
      msLogin = System.Configuration.ConfigurationManager.AppSettings("Login").Trim
      msPassword = System.Configuration.ConfigurationManager.AppSettings("Password").Trim
      msWorkflowURL = System.Configuration.ConfigurationManager.AppSettings("WorkflowURL").Trim
    Catch ex As Exception

    End Try

  End Sub
  Public Function Server() As String
    Server = ""

    Try
      If msServerName.Length > 0 Then
        Server = msServerName
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function Database() As String
    Database = ""

    Try
      If msDatabaseName.Length > 0 Then
        Database = msDatabaseName
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function Login() As String
    Login = ""

    Try
      If msLogin.Length > 0 Then
        Login = msLogin
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function Password() As String
    Password = ""

    Try
      If msPassword.Length > 0 Then
        Password = msPassword
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function WorkflowURL() As String
    WorkflowURL = ""

    Try
      If msWorkflowURL.Length > 0 Then
        WorkflowURL = msWorkflowURL
      End If
    Catch ex As Exception
    End Try

  End Function

  Public Function ColourThemeFolder() As String
    ColourThemeFolder = "Blanco"

    Try
      If msThemeName.Length > 0 Then
        ColourThemeFolder = msThemeName
      End If
    Catch ex As Exception
    End Try

  End Function

  Public Function ColourThemeHex() As String
    ' Default to Blanco
    ColourThemeHex = "#FFFFFF"

    Try
      If msThemeHex.Length > 0 Then
        ColourThemeHex = "#" & msThemeHex
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function ColourThemeForeColour() As String
    ColourThemeForeColour = "Black"

    Try
      If msThemeFore.Length > 0 Then
        ColourThemeForeColour = msThemeFore
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function MessageFontSize() As Int32
    MessageFontSize = 10

    Try
      If miMessageFontSize > 0 Then
        MessageFontSize = miMessageFontSize
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function SubmissionTimeout() As Int32
    ' Return the configured SubmissionTimeout in milliseconds.
    ' This value is used for WARP submission timeout, and SQL command timeout.
    ' Defaulted to 2 minutes
    SubmissionTimeout = SubmissionTimeoutInSeconds() * 1000
  End Function
  Public Function SubmissionTimeoutInSeconds() As Int32
    ' Return the configured SubmissionTimeout in seconds.
    ' This value is used for WARP submission timeout, and SQL command timeout.
    ' Defaulted to 2 minutes
    SubmissionTimeoutInSeconds = 120

    Try
      If miSubmissionTimeout > 0 Then
        SubmissionTimeoutInSeconds = miSubmissionTimeout
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function LookupRowsRange() As Int32
    ' Return the configured number of records to load by default in the lookup dropdown grids.
    ' Defaulted to 100 rows
    LookupRowsRange = 100

    Try
      If miLookupRowsRange > 0 Then
        LookupRowsRange = miLookupRowsRange
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function ValidationMessageFontSize() As Int32
    ValidationMessageFontSize = 8

    Try
      If miValidationMessageFontSize > 0 Then
        ValidationMessageFontSize = miValidationMessageFontSize
      End If
    Catch ex As Exception
    End Try

  End Function
  Public Function OLEFolder_Server() As String
    OLEFolder_Server = ""

    Try
      OLEFolder_Server = msOLEFolder_Server
    Catch ex As Exception
    End Try

  End Function
  Public Function OLEFolder_Local() As String
    OLEFolder_Local = ""

    Try
      OLEFolder_Local = msOLEFolder_Local
    Catch ex As Exception
    End Try

  End Function
  Public Function PhotographFolder() As String
    PhotographFolder = ""

    Try
      PhotographFolder = msPhotographFolder
    Catch ex As Exception
    End Try

  End Function
  Public Sub New()

  End Sub
End Class
