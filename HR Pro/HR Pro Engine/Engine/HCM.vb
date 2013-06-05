Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.None)> _
Public Class HCM
  Implements COMInterfaces.iSystemManager

  Private objDatabase As New Connectivity.SQL
  Private mobjScript As New ScriptDB.Script

  Public Property DB As Object Implements COMInterfaces.iSystemManager.CommitDB, COMInterfaces.iSystemManager.MetadataDB
    Get
      Return objDatabase
    End Get
    Set(ByVal value As Object)
      objDatabase = value
    End Set
  End Property

  Public Function Initialise() As Boolean Implements iSystemManager.Initialise

    Dim bOK As Boolean = True

    Try
      Globals.Initialise()

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function

  Public Function CloseSafely() As Boolean Implements iSystemManager.CloseSafely
    Return True
  End Function

  Public Function PopulateObjects() As Boolean Implements COMInterfaces.iSystemManager.PopulateObjects

    Dim bOK As Boolean = True

    Try

      If Options Is Nothing Then
        Globals.Initialise()
      End If

      Globals.MetadataDB = objDatabase
      Globals.CommitDB = objDatabase
      Globals.Options.DevelopmentMode = False

      'Things.PopulateSystemThings()
      '       PopulateSystemSettings()
      Things.PopulateTables()
      Things.PopulateColumns()
      Things.PopulateScreens()
      Things.PopulateExpressions()


      '        PopulateModuleSettings()

    Catch ex As Exception
      bOK = False
    End Try

    Return bOK

  End Function


  Public ReadOnly Property ErrorLog As ErrorHandler.Errors Implements COMInterfaces.iSystemManager.ErrorLog
    Get
      Return Globals.ErrorLog
    End Get
  End Property

  Public ReadOnly Property Options As HCMOptions Implements COMInterfaces.iSystemManager.Options
    Get
      Return Globals.Options
    End Get
  End Property

  Public ReadOnly Property Script As ScriptDB.Script Implements COMInterfaces.iSystemManager.Script
    Get
      Return mobjScript
    End Get
  End Property

  Public ReadOnly Property ReturnThings As Things.Collection Implements COMInterfaces.iSystemManager.Things
    Get
      Return Globals.Things
    End Get
  End Property

  Public ReadOnly Property TuningLog As Tuning.Report Implements COMInterfaces.iSystemManager.TuningLog
    Get
      Return Globals.TuningLog
    End Get
  End Property

  Public ReadOnly Property Version As System.Version Implements COMInterfaces.iSystemManager.Version
    Get
      Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
    End Get
  End Property

#Region "Connectivity"

  Public Function Connect(ByRef Login As Connectivity.Login) As Boolean

    Dim bOK As Boolean = True

    Try

      If Login.UserName = vbNullString Then
        bOK = False
      Else
        objDatabase.Login = Login
        objDatabase.Open()
      End If

    Catch ex As Exception
      bOK = False

    End Try

    Return bOK

  End Function

  Public Function Disconnect() As Boolean

    Dim bOK As Boolean

    Try
      objDatabase.Close()

    Catch ex As Exception
      bOK = False

    End Try

    Return bOK

  End Function

#End Region


End Class

