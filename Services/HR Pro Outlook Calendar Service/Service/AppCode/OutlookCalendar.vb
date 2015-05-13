Option Strict Off

Imports Redemption
Imports System.Globalization
Imports System.Runtime.InteropServices

''' <summary>
''' This Class creates an Outlook appointment using late binding
''' </summary>
''' <remarks>
''' Late binding is used as we need to support Outlook 2000 and
''' Microsoft haven't produced and PIAs for Office 2000
''' </remarks>
Public Class OutlookCalendar
    Implements IDisposable

#Region "Declarations"

    Private _session As RDOSession

    Private _sw As IO.StreamWriter = Nothing
    Private _enableTrace As Boolean = False

    Private _mailboxName As String = String.Empty
    Private _dateFormat As String = String.Empty
    Private _errorMessage As String = String.Empty
    Private _storeID As String = String.Empty
    Private _entryID As String = String.Empty
    Private _reminder As Boolean = False
    Private _reminderOffset As Int32 = 0
    Private _reminderPeriod As Int32 = 0
    Private _allDayEvent As Boolean = False
    Private _startDate As DateTime
    Private _endDate As DateTime
    Private _startTime As String = String.Empty
    Private _endTime As String = String.Empty
    Private _subject As String = String.Empty
    Private _content As String = String.Empty
    Private _busyStatus As Int32 = 0
    Private _folder As String = String.Empty
    Private _serviceAccountPassword As String = String.Empty

#End Region

    Public Sub New(ByVal sw As IO.StreamWriter, ByVal enableTrace As Boolean, serviceAccountPassword As String)
        _sw = sw
        _enableTrace = enableTrace
        _dateFormat = DateFormat()
        _serviceAccountPassword = serviceAccountPassword
    End Sub

    Public Function Logon(ByVal exchangeServer As String) As Boolean

        Try
            If _session Is Nothing Then
                _session = Redemption.RedemptionLoader.new_RDOSession
            End If

            Dim userName As String = _session.CurrentWindowsUser.NTAccountName
            '_session.LogonExchangeMailbox(userName, exchangeServer) 'Exchange 2010
            _session.LogonHostedExchangeMailbox(_session.CurrentWindowsUser.SMTPAddress, userName, _serviceAccountPassword) 'Exchange 2010 and 2013
            LogEvent("Logged in successfully to Exchange server '" & exchangeServer & "'")
        Catch ex As COMException
            _errorMessage = String.Format("Unable to connect to exchange server '{0}'.{1}{1}Check the configured user account settings.", _
                                          exchangeServer, ControlChars.NewLine)
            TraceLog("LOGON ERROR : " & ex.Message, _sw, _enableTrace)

            Return LoggedOn
        Catch ex As Exception
            _errorMessage = String.Format("Unable to connect to exchange server '{0}'.{1}{1}Check the configured user account settings.", _
                                          exchangeServer, ControlChars.NewLine)
            TraceLog("LOGON ERROR : " & ex.Message, _sw, _enableTrace)

            Return LoggedOn
        End Try

        Return LoggedOn

    End Function

    Public Function CreateEntry() As Boolean

        If _session Is Nothing Then
            _errorMessage = "Exchange Logon failed"
            Return False
        End If

        If _startDate.ToString() = String.Empty Then
            _errorMessage = "No start date entered"
            Return False
        End If

        If Not _allDayEvent Then
            If Not _startTime Like "##:##" Then
                _errorMessage = String.Concat("Invalid Start Time <", _startTime, ">")
                Return False
            End If
            If Not (_endTime Like "##:##") Then
                _errorMessage = String.Concat("Invalid End Time <", _endTime, ">")
                Return False
            End If
        End If

        If _folder = String.Empty Then
            _errorMessage = String.Concat("Outlook folder name empty")
            Return False
        End If

        _errorMessage = String.Empty

        Dim folderItem As RDOFolder = Nothing

        If Not _folder.Contains("\\Public Folders") Then
            Dim mailboxItem As RDOStore = Nothing
            Try
                Dim mailboxName As String = GetNameFromMailbox(_folder)
                TraceLog("GetSharedMailbox : " & mailboxName, _sw, _enableTrace)
                mailboxItem = _session.GetSharedMailbox(mailboxName)
                TraceLog("GetSharedMailbox OK", _sw, _enableTrace)
            Catch
                _errorMessage = String.Format("Unable to open mailbox for {0}.  Check permissions.", GetNameFromMailbox(_folder))

                ReleaseItem(mailboxItem)
                Return False
            End Try
            Try
                TraceLog("GetFolderFromPath : " & _folder, _sw, _enableTrace)
                folderItem = GetFolderFromPath(_folder, mailboxItem.IPMRootFolder.Folders)

            Catch
                _errorMessage = String.Format("Unable to open mailbox for {0}.  Check permissions.", GetNameFromMailbox(_folder))

                Return False
            Finally
                ReleaseItem(mailboxItem)
            End Try
        Else
            TraceLog("Public Folder: " & _folder, _sw, _enableTrace)
            Try
                Dim storeName = _session.Stores.DefaultStore.IPMRootFolder.Name

                For Each store As RDOStore In _session.Stores
                    If (store.StoreKind = TxStoreKind.skPublicFolders) Then
                        folderItem = GetFolderFromPath(_folder, store.IPMRootFolder.Folders)

                        If folderItem IsNot Nothing Then
                            TraceLog("GetFolderFromPath OK", _sw, _enableTrace)
                            Exit For
                        Else
                            TraceLog("GetFolderFromPath - Cant find folder", _sw, _enableTrace)
                        End If
                    End If
                Next

            Catch
                TraceLog("Public Folder: NO ACCESS TO STORE" & _folder, _sw, _enableTrace)
            End Try
        End If

        If folderItem Is Nothing Then
            _errorMessage = "Unable to obtain a valid Calendar path from :- " & _folder
            Return False
        End If

        If _errorMessage <> String.Empty Then
            Return False
        End If

        Try
            Dim apptItem As RDOAppointmentItem = folderItem.Items.Add()

            With apptItem
                If _allDayEvent Then
                    .AllDayEvent = True
                    .Start = _startDate
                    'If no times are specified then outlook correctly finishes at midnight but does not include the end date.  
                    'For OpenHR we need the event to be inclusive of both the start date and end date so if its an all day
                    'event add one day to the end date.
                    If Date.Compare(_startDate, _endDate) > 0 Then
                        'Start date after end date
                        .End = _startDate.AddDays(1)
                    Else
                        .End = _endDate.AddDays(1)
                    End If
                Else
                    .AllDayEvent = False
                    _startDate = CDate(_startDate.Date.ToString(_dateFormat) & " " & _startTime)
                    .Start = _startDate
                    _endDate = CDate(_endDate.ToString(_dateFormat) & " " & _endTime)
                    .End = _endDate
                End If
                .Subject = _subject
                .Body = _content
                '.BusyStatus = Redemption.rdoBusyStatus.olBusy
                .BusyStatus = _busyStatus
                .ReminderSet = _reminder
                If _reminder Then
                    .ReminderMinutesBeforeStart = CInt(_reminderOffset) * CInt(Choose(_reminderPeriod + 1, 1, 1440, 10080, 40240))
                End If

                .Save()

                _storeID = folderItem.StoreID
                _entryID = .EntryID

            End With
            ReleaseItem(apptItem)

        Catch ex As Exception
            _errorMessage = ex.Message
            Return False
        Finally
            ReleaseItem(folderItem)
        End Try

        Return True
    End Function

    Public Function DeleteEntry() As Boolean

        Dim apptItem As RDOAppointmentItem

        ' First lets try and grab the original entry, if this errors
        ' then it must've been deleted from Outlook
        Try
            apptItem = _session.GetMessageFromID(_entryID, _storeID)

        Catch ex As Exception
            ' Original entry does not exist in Outlook Calendar
            Return True
        End Try

        ' Now lets try deleting it..... if this part fails
        ' then there is a folder permissions issue
        Try
            apptItem.Delete()
        Catch ex As Exception
            _errorMessage = "Cannot delete entry - Check Mailbox and Calendar folder permissions"
            Return False
        Finally
            ReleaseItem(apptItem)
        End Try

        Return True

    End Function

    Private Function GetFolderFromPath( _
      ByVal path As String, _
      ByVal myfolders As RDOFolders) As RDOFolder

        Try
            If myfolders Is Nothing Then
                Return Nothing
            End If

            TraceLog("Mailbox : " & path, _sw, _enableTrace)

            Dim olFolder As RDOFolder = Nothing

            Dim pathArray() As String
            pathArray = path.Split("\"c)

            TraceLog("Folder count : " & myfolders.Count, _sw, _enableTrace)

            For count As Integer = 0 To pathArray.Length - 1
                If pathArray(count) <> "" Then
                    If olFolder Is Nothing Then
                        For Each olTempFolder As RDOFolder In myfolders
                            If (Not olTempFolder.Name.Equals(String.Empty)) _
                              AndAlso (olTempFolder.Name.ToLower.Trim = pathArray(count).ToLower.Trim) Then

                                olFolder = olTempFolder
                                Exit For
                            End If
                        Next
                    Else
                        If olFolder.Folders IsNot Nothing Then
                            For Each olTempFolder As RDOFolder In olFolder.Folders
                                If (Not olTempFolder.Name.Equals(String.Empty)) _
                                  AndAlso (olTempFolder.Name.ToLower.Trim = pathArray(count).ToLower.Trim) Then

                                    olFolder = olTempFolder
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
            Next

            If (olFolder IsNot Nothing) Then
                TraceLog("Mailbox found", _sw, _enableTrace)
            Else
                TraceLog("Mailbox NOT found", _sw, _enableTrace)
            End If

            Return olFolder

        Catch ex As Exception
            TraceLog("*** ERROR OCCURED LOCATING MAILBOX ***", _sw, _enableTrace)
            TraceLog("  -- " & ex.Message, _sw, _enableTrace)
            TraceLog(_sw, _enableTrace)

            _errorMessage = ex.Message
            Return Nothing
        Finally

        End Try
    End Function

    Private Function GetNameFromMailbox(ByVal mailboxPath As String) As String
        Dim tmpName As String = mailboxPath
        Try
            If mailboxPath.Contains("\\Mailbox") Then
                tmpName = mailboxPath.Replace("\\Mailbox - ", "")
                tmpName = tmpName.Substring(0, tmpName.IndexOf("\"))
            ElseIf mailboxPath.StartsWith("\\") Then
                tmpName = mailboxPath.Substring(mailboxPath.IndexOf("\\") + 2)
                tmpName = tmpName.Substring(0, tmpName.IndexOf("\"))
            End If
        Catch
            Return String.Empty
        End Try

        Return tmpName
    End Function

    Public Function Quit() As Boolean

        Try
            If _session IsNot Nothing Then
                _session.Logoff()
                ReleaseItem(_session)
            Else
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End If

        Catch ex As Exception
            _errorMessage = ex.Message
            Return False
        End Try

        Return True
    End Function

    Private Function DateFormat() As String
        ' NB. Windows allows the user to configure totally stupid
        ' date formats (eg. d/M/yyMydy !). This function does not cater
        ' for such stupidity, and simply takes the first occurence of the
        ' 'd', 'M', 'y' characters.
        Dim sysFormat As String = String.Empty
        Dim sysDateSeparator As String = String.Empty
        Dim sysDateFormat As String = String.Empty
        Dim daysDone As Boolean = False
        Dim monthsDone As Boolean = False
        Dim yearsDone As Boolean = False

        sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        sysDateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator

        ' Loop through the string picking out the required characters.
        For iLoop As Integer = 0 To sysFormat.Length - 1

            Select Case sysFormat.Substring(iLoop, 1)
                Case "d"
                    If Not daysDone Then
                        ' Ensure we have two day characters.
                        sysDateFormat &= "dd"
                        daysDone = True
                    End If

                Case "M"
                    If Not monthsDone Then
                        ' Ensure we have two month characters.
                        sysDateFormat &= "MM"
                        monthsDone = True
                    End If

                Case "y"
                    If Not yearsDone Then
                        ' Ensure we have four year characters.
                        sysDateFormat &= "yyyy"
                        yearsDone = True
                    End If

                Case Else
                    sysDateFormat &= sysFormat.Substring(iLoop, 1)
            End Select

        Next iLoop

        ' Ensure that all day, month and year parts of the date
        ' are present in the format.
        If Not daysDone Then
            If sysDateFormat.Substring(sysDateFormat.Length - 1, 1) <> sysDateSeparator Then
                sysDateFormat &= sysDateSeparator
            End If

            sysDateFormat &= "dd"
        End If

        If Not monthsDone Then
            If sysDateFormat.Substring(sysDateFormat.Length - 1, 1) <> sysDateSeparator Then
                sysDateFormat &= sysDateSeparator
            End If

            sysDateFormat &= "MM"
        End If

        If Not yearsDone Then
            If sysDateFormat.Substring(sysDateFormat.Length - 1, 1) <> sysDateSeparator Then
                sysDateFormat &= sysDateSeparator
            End If

            sysDateFormat &= "yyyy"
        End If

        ' Return the date format.
        Return sysDateFormat

    End Function

    Public Function ResetStoreAndEntry() As Boolean
        _storeID = String.Empty
        _entryID = String.Empty
        Return True
    End Function

    Private Sub ReleaseItem(ByVal item As Object)
        If item IsNot Nothing Then
            Runtime.InteropServices.Marshal.ReleaseComObject(item)
            item = Nothing
        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

#Region "Properties"

    Public ReadOnly Property LoggedOn() As Boolean
        Get
            If _session IsNot Nothing Then
                Return _session.LoggedOn
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property ErrorMessage() As String
        Get
            Return _errorMessage
        End Get
    End Property

    Public Property Mailbox() As String
        Get
            Return Me._mailboxName
        End Get
        Set(ByVal value As String)
            Me._mailboxName = value
        End Set
    End Property

    Public Property StoreID() As String
        Get
            Return Me._storeID
        End Get
        Set(ByVal value As String)
            Me._storeID = value
        End Set
    End Property

    Public Property EntryID() As String
        Get
            Return Me._entryID
        End Get
        Set(ByVal value As String)
            Me._entryID = value
        End Set
    End Property

    Public WriteOnly Property Reminder() As Boolean
        Set(ByVal value As Boolean)
            _reminder = value
        End Set
    End Property

    Public WriteOnly Property ReminderOffset() As Int32
        Set(ByVal value As Int32)
            _reminderOffset = value
        End Set
    End Property

    Public WriteOnly Property ReminderPeriod() As Int32
        Set(ByVal value As Int32)
            _reminderPeriod = value
        End Set
    End Property

    Public WriteOnly Property AllDayEvent() As Boolean
        Set(ByVal value As Boolean)
            _allDayEvent = value
        End Set
    End Property

    Public Property StartDate() As DateTime
        Get
            Return _startDate
        End Get
        Set(ByVal value As DateTime)
            _startDate = value
            _endDate = _startDate
        End Set
    End Property

    Public Property EndDate() As DateTime
        Get
            Return _endDate
        End Get
        Set(ByVal value As DateTime)
            _endDate = value
        End Set
    End Property

    Public WriteOnly Property StartTime() As String
        Set(ByVal value As String)
            _startTime = value
        End Set
    End Property

    Public WriteOnly Property EndTime() As String
        Set(ByVal value As String)
            _endTime = value
        End Set
    End Property

    Public Property Subject() As String
        Get
            Return _subject
        End Get
        Set(ByVal value As String)
            _subject = value
        End Set
    End Property

    Public WriteOnly Property Content() As String
        Set(ByVal value As String)
            _content = value
        End Set
    End Property

    Public WriteOnly Property BusyStatus() As Int32
        Set(ByVal value As Int32)
            _busyStatus = value
        End Set
    End Property

    Public Property Folder() As String
        Get
            Return _folder
        End Get
        Set(ByVal value As String)
            _folder = value
        End Set
    End Property

#End Region

#Region " IDisposable Support "
    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then

            End If

            If _session IsNot Nothing Then
                _session.Logoff()
                ReleaseItem(_session)
            End If
        End If

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region


End Class
