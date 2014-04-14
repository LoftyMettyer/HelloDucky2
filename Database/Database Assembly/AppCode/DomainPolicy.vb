Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.DirectoryServices

Public Class DomainPolicy
  Implements IDisposable

  Private _attribs As ResultPropertyCollection

  Public Sub New(ByVal domainRoot As DirectoryEntry)

    Dim policyAttributes As String() = New String() {"maxPwdAge", "minPwdAge", "minPwdLength", "lockoutDuration", "lockOutObservationWindow", _
                                                    "lockoutThreshold", "pwdProperties", "pwdHistoryLength", "objectClass", "distinguishedName"}

    'we take advantage of the marshaling with the DirectorySearcher for Large Int32 values...
    Dim ds As DirectorySearcher = New DirectorySearcher(domainRoot, "(objectClass=domainDNS)", policyAttributes, SearchScope.Base)

    Try
      Dim result As SearchResult = ds.FindOne()

      'do some quick validation...       
      If result Is Nothing Then
        Throw New ArgumentException("domainRoot is not a domainDNS object.")
      End If

      _attribs = result.Properties
      result = Nothing

    Finally
      ds.Dispose()
      ds = Nothing
    End Try

  End Sub

  'for some odd reason, the intervals are all stored as negative numbers.  We use this to "invert" them
  Private Function GetAbsLongValue(ByVal longInt As Object) As Long
    'AE20080311 Fault #12998
    If CType(longInt, Int64) = Long.MinValue Then
      Return 0
    Else
      Return Math.Abs(CType(Fix(longInt), Int64))
    End If
  End Function

  Public ReadOnly Property DomainDistinguishedName() As String
    Get
      Dim val As String = "distinguishedName"
      If _attribs.Contains(val) Then
        Return CStr(_attribs(val)(0))
      End If
      'default return value
      Return String.Empty
    End Get
  End Property

  Public ReadOnly Property MaxPasswordAge() As TimeSpan
    Get
      Dim val As String = "maxPwdAge"
      If _attribs.Contains(val) Then
        Dim ticks As Long = GetAbsLongValue(_attribs(val)(0))

        If ticks > 0 Then
          Return TimeSpan.FromTicks(ticks)
        End If
      End If

      Return TimeSpan.Zero
    End Get
  End Property

  Public ReadOnly Property MinPasswordAge() As TimeSpan
    Get
      Dim val As String = "minPwdAge"
      If _attribs.Contains(val) Then
        Dim ticks As Long = GetAbsLongValue(_attribs(val)(0))

        If ticks > 0 Then
          Return TimeSpan.FromTicks(ticks)
        End If
      End If

      Return TimeSpan.Zero
    End Get
  End Property

  Public ReadOnly Property LockoutDuration() As TimeSpan
    Get
      Dim val As String = "lockoutDuration"
      If _attribs.Contains(val) Then
        Dim ticks As Long = GetAbsLongValue(_attribs(val)(0))

        If ticks > 0 Then
          Return TimeSpan.FromTicks(ticks)
        End If
      End If

      Return TimeSpan.Zero
    End Get
  End Property

  Public ReadOnly Property LockoutObservationWindow() As TimeSpan
    Get
      Dim val As String = "lockoutObservationWindow"
      If _attribs.Contains(val) Then
        Dim ticks As Long = GetAbsLongValue(_attribs(val)(0))

        If ticks > 0 Then
          Return TimeSpan.FromTicks(ticks)
        End If
      End If

      Return TimeSpan.Zero
    End Get
  End Property

  Public ReadOnly Property LockoutThreshold() As Int32
    Get
      Dim val As String = "lockoutThreshold"
      If _attribs.Contains(val) Then
        Return CType(Fix(_attribs(val)(0)), Int32)
      End If

      Return 0
    End Get
  End Property

  Public ReadOnly Property MinPasswordLength() As Int32
    Get
      Dim val As String = "minPwdLength"
      If _attribs.Contains(val) Then
        Return CType(Fix(_attribs(val)(0)), Int32)
      End If

      Return 0
    End Get
  End Property

  Public ReadOnly Property PasswordHistoryLength() As Int32
    Get
      Dim val As String = "pwdHistoryLength"
      If _attribs.Contains(val) Then
        Return CType(Fix(_attribs(val)(0)), Int32)
      End If

      Return 0
    End Get
  End Property

  Public ReadOnly Property PasswordProperties() As Int32
    Get
      Dim val As String = "pwdProperties"
      'this should fail if not found
      Return CType(Fix(_attribs(val)(0)), Int32)
    End Get
  End Property

  Private disposedValue As Boolean = False    ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(ByVal disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        'Call Dispose() on managed objects
      End If

      'release unmanaged resource(s) held by this object
      _attribs = Nothing
    End If
    Me.disposedValue = True
  End Sub

#Region " IDisposable Support "
  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub

  Protected Overrides Sub Finalize()
    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    Dispose(False)
    MyBase.Finalize()
  End Sub
#End Region
End Class
