using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.VisualBasic;
using OutlookCalendarLogic.Exchange;
using OutlookCalendarLogic.Structures;
using Microsoft.VisualBasic.CompilerServices;

//Exchange 2013 101 Code Samples
//The samples in the Exchange 2013: 101 code samples package show you how to use the Exchange Web Services (EWS) Managed API to perform specific tasks
//with mailbox data on an on-premises Exchange server, Exchange Online, or Exchange Online as part of Office 365. 
//https://code.msdn.microsoft.com/Exchange-2013-101-Code-3c38582c

namespace OutlookCalendarLogic
{
  public class Worker
  {
    public delegate void MessageEventHandler(MessageEventDetails e);
    public event MessageEventHandler RaiseMessageEvent;

    #region Private variables
    private ExchangeService _exchangeService;
    private UserDefinedConfiguration _userDefinedConfiguration;
    private readonly string _exchangeURL;
    private readonly string _exchangeUser;
    private readonly string _exchangeUserPassword;
    private readonly bool _useDefaultCredentials;
    private readonly bool _enableExchangeTrace;
    private bool _isProcessing;
    private bool _loggedOn;
    private VersionNumber _svcVersion;
    private const Single MINIMUMDBVERSION = 8.0f;
    private string _storeId;
    private string _entryId;
    private string _dateFormat = string.Empty;
    private string _errorMessage = string.Empty;
    private bool _reminder = false;
    private Int32 _reminderOffset = 0;
    private Int32 _reminderPeriod = 0;
    private bool _allDayEvent = false;
    private DateTime _startDate;
    private DateTime _endDate;
    private string _startTime = string.Empty;
    private string _endTime = string.Empty;
    private string _subject = string.Empty;
    private string _content = string.Empty;
    private Int32 _busyStatus = 0;
    private string _folder = string.Empty;


    #endregion

    #region Public properties
    public List<ConfigurationError> ConfigurationErrors { get; set; }
    public List<string> Messages { get; set; }
    public UserDefinedConfiguration UserDefinedConfiguration => _userDefinedConfiguration;
    public VersionNumber ServiceVersion => _svcVersion;
    #endregion

    #region Public methods
    public Worker(string exchangeUser, bool enableExchangeTrace)
    {
      _exchangeUser = exchangeUser;
      _enableExchangeTrace = enableExchangeTrace;
      _useDefaultCredentials = true;
      InitializeWorker();
    }

    public Worker(string exchangeUser, string exchangeUserPassword, bool enableExchangeTrace)
    {
      _exchangeUser = exchangeUser;
      _exchangeUserPassword = exchangeUserPassword;
      _enableExchangeTrace = enableExchangeTrace;
      _useDefaultCredentials = false;
      InitializeWorker();
    }

    public bool GetAndCheckUserConfiguration()
    {
      var configOk = true;//Assume the config is Ok unless we find errors below
      var tempOpenHrSystems = new List<OpenHRSystem>();

      //Get the user configuration
      try
      {
        Configuration configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        RaiseMessageEvent(new MessageEventDetails("Configuration file: " + configFile.FilePath));

        _userDefinedConfiguration = new UserDefinedConfiguration
        {
          OpenHRUser = ConfigurationManager.AppSettings["openhruser"],
          OpenHRPassword = ConfigurationManager.AppSettings["openhrpassword"],
          ExchangeServer = ConfigurationManager.AppSettings["exchange"],
          ExchangeServerURL = ConfigurationManager.AppSettings["exchangeURL"],
          ServiceAccountPassword = ConfigurationManager.AppSettings["serviceAccountPassword"],
          Debug = ConfigurationManager.AppSettings["debug"] != "0"
        };

        int commandTimeout;
        _userDefinedConfiguration.CommandTimeout = int.TryParse(ConfigurationManager.AppSettings["commandtimeout"], out commandTimeout) ? commandTimeout : 200;

        foreach (string key in ConfigurationManager.AppSettings.Keys)
        {
          if (key.Length >= 6 && key.ToLower().Trim().Substring(0, 6) == "server")
          {
            var server = ConfigurationManager.AppSettings[key];
            var serverNumber = "";
            if (key.Length > 6)
              serverNumber = key.Substring(6);

            var openHrSystem = new OpenHRSystem(server, ConfigurationManager.AppSettings["database" + serverNumber]);
            openHrSystem.UserName = ConfigurationManager.AppSettings["openhruser"];
            openHrSystem.Password = ConfigurationManager.AppSettings["openhrpassword"];
            openHrSystem.Serviced = true;
            openHrSystem.ServiceServer = Environment.MachineName;
            openHrSystem.VersionOk = true;

            tempOpenHrSystems.Add(openHrSystem);
          }
        }
      }
      catch (Exception)
      {
        ConfigurationErrors.Add(new ConfigurationError("An error has occured while reading the application's settings", EventLogEntryType.Error));
        return false;
      }

      //Describe which systems we're servicing, and add the correctly-configured ones to the list of OpenHRSystems
      _userDefinedConfiguration.OpenHRSystems = new List<OpenHRSystem>();
      foreach (OpenHRSystem ohrs in tempOpenHrSystems)
      {
        //TODO: more checks to the OpenHRSystems to service
        _userDefinedConfiguration.OpenHRSystems.Add(ohrs);
      }

      //Check the user configuration and register any errors; also generate any messages than can be used by the calling application to log them or whatever
      if (_userDefinedConfiguration.ExchangeServer.Equals(String.Empty))
      {
        ConfigurationErrors.Add(new ConfigurationError("No exchange server parameter defined", EventLogEntryType.Error));
        configOk = false;
      }

      if (_userDefinedConfiguration.OpenHRUser == null)
      {
        ConfigurationErrors.Add(new ConfigurationError("'OpenHRUser' configuration value is not defined", EventLogEntryType.Error));
        configOk = false;
      }

      if (_userDefinedConfiguration.OpenHRSystems.Count == 0)
      {
        ConfigurationErrors.Add(new ConfigurationError("No server and database parameters defined", EventLogEntryType.Error));
        configOk = false;
      }

      //Raise message events for each configuration error, if any
      if (ConfigurationErrors.Count > 0)
      {
        RaiseMessageEvent(new MessageEventDetails("Configuration errors:"));
        foreach (var configError in ConfigurationErrors)
          RaiseMessageEvent(new MessageEventDetails(configError.Message, configError.Severity,
            MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
      }

      //
      return configOk;
    }

    public void ProcessEntries()
    {
      if (!_isProcessing)
      {
        _isProcessing = true;
        OutlookBatch();
        _isProcessing = false;
      }
    }
    #endregion

    #region Exchange callback methods
    private bool RedirectionUrlValidationCallback(string redirectionUrl)
    {
      RaiseMessageEvent(new MessageEventDetails("Using autodiscover Url " + redirectionUrl));

      // The default for the validation callback is to reject the URL.
      bool result = false;

      Uri redirectionUri = new Uri(redirectionUrl);

      // Validate the contents of the redirection URL. In this simple validation
      // callback, the redirection URL is considered valid if it is using HTTPS
      // to encrypt the authentication credentials. 
      if (redirectionUri.Scheme == "https")
      {
        result = true;
      }
      return result;
    }

    private static bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
      {
        if (chain != null)
        {
          foreach (X509ChainStatus status in chain.ChainStatus)
          {
            if ((certificate.Subject == certificate.Issuer) && (status.Status == X509ChainStatusFlags.UntrustedRoot))
            { // Self-signed certificates with an untrusted root are valid. 
              continue;
            }
            else {
              if (status.Status != X509ChainStatusFlags.NoError)
              {
                // If there are any other errors in the certificate chain, the certificate is invalid, so the method returns false.
                return false;
              }
            }
          }
        }

        // When processing reaches this line, the only errors in the certificate chain are 
        // untrusted root errors for self-signed certificates. These certificates are valid
        // for default Exchange server installations, so return true.
        return true;
      }
      else {
        // In all other cases, return false.
        return false;
      }
    }
    #endregion

    #region Database checks
    private bool DatabaseIsOk(OpenHRSystem openHrSystem, SqlConnection conn)
    {
      bool returnOk = true;

      // Check if the given database is locked.
      returnOk = !DatabaseIsLocked(ref openHrSystem, conn);
      RaiseMessageEvent(new MessageEventDetails("Is database locked?: " + (!returnOk)));

      if (returnOk)
      {
        // Check if the given database is in the middle of the overnight job update.
        returnOk = !DatabaseIsRunningOvernight(ref openHrSystem, conn);
        RaiseMessageEvent(new MessageEventDetails("Is database running overnight process?: " + (!returnOk)));
      }

      if (returnOk)
      {
        // Check if the given database is the correct version.
        returnOk = DatabaseVersionIsOK(ref openHrSystem, conn);
        RaiseMessageEvent(new MessageEventDetails("Is database version okay?: " + returnOk));
      }

      if (returnOk)
      {
        // Check to see if the specified database is already being service elsewhere
        returnOk = !DatabaseIsBeingServiced(ref openHrSystem, conn);
        RaiseMessageEvent(new MessageEventDetails("Is database being serviced?: " + (!returnOk)));
      }

      return returnOk;
    }

    private bool DatabaseIsLocked(ref OpenHRSystem openHRSystem, SqlConnection conn)
    {
      SqlCommand lockCheck = new SqlCommand();
      bool returnBool = false;

      try
      {
        lockCheck.CommandText = "sp_ASRLockCheck";
        lockCheck.Connection = conn;
        lockCheck.CommandType = CommandType.StoredProcedure;
        lockCheck.CommandTimeout = _userDefinedConfiguration.CommandTimeout;

        SqlDataReader reader = lockCheck.ExecuteReader();

        while (reader.Read())
        {
          if (Utilities.NullSafeInteger(reader["priority"]) != 3)
          {
            returnBool = true;
          }
        }

        reader.Close();
        reader = null;

        if (!openHRSystem.Locked && returnBool)
        {
          RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) locked.", MessageEventDetails.MessageEventType.WindowsEventsLog));
        }
        else if (openHRSystem.Locked && !returnBool)
        {
          RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) unlocked.", MessageEventDetails.MessageEventType.WindowsEventsLog));
        }

        openHRSystem.Locked = returnBool;

      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("Lock check ({openHRSystem.ToString()}) - {ex.Message} {ex.StackTrace}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }
      finally
      {
        lockCheck.Dispose();
        lockCheck = null;
      }

      return returnBool;

    }

    private bool DatabaseIsBeingServiced(ref OpenHRSystem openHRSystem, SqlConnection conn)
    {
      bool returnBool = false;
      DateTime lastRun = (DateTime)SqlDateTime.MinValue;
      string runServer = string.Empty;

      try
      {
        if (GetSystemSetting("outlook service 2", "running", conn) == "1")
        {//Is the service running?
          try
          {
            lastRun = Convert.ToDateTime(GetSystemSetting("outlook service 2", "last run", conn));
          }
          catch
          {
            // Can't convert to a datetime so we'll stick with the MinValue
          }

          runServer = GetSystemSetting("outlook service 2", "server", conn);

          if ((runServer == openHRSystem.ServiceServer))
          {
            SaveSystemSetting("outlook service 2", "running", "1", conn);
            SaveSystemSetting("outlook service 2", "server", Environment.MachineName, conn);
            SaveSystemSetting("outlook service 2", "last run", DateTime.Now.ToString(), conn);

            openHRSystem.Serviced = true;
            returnBool = false;

          }
          else if ((lastRun != SqlDateTime.MinValue) && DateTime.Now.Subtract(lastRun).Minutes >= 5)
          {
            SaveSystemSetting("outlook service 2", "running", "1", conn);
            SaveSystemSetting("outlook service 2", "server", Environment.MachineName, conn);
            SaveSystemSetting("outlook service 2", "last run", DateTime.Now.ToString(), conn);

            RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) now being serviced.", MessageEventDetails.MessageEventType.WindowsEventsLog));

            openHRSystem.ServiceServer = Environment.MachineName;
            openHRSystem.Serviced = true;
            returnBool = false;

          }
          else if (openHRSystem.Serviced)
          {
            RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) already being serviced by server {runServer}.", EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLog));
            openHRSystem.Serviced = false;
            returnBool = true;
          }
        }
        else {
          if (openHRSystem.Serviced == false)
          {
            RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) now being serviced.", MessageEventDetails.MessageEventType.WindowsEventsLog));
          }

          SaveSystemSetting("outlook service 2", "running", "1", conn);
          SaveSystemSetting("outlook service 2", "server", Environment.MachineName, conn);
          SaveSystemSetting("outlook service 2", "last run", DateTime.Now.ToString(), conn);

          openHRSystem.ServiceServer = Environment.MachineName;
          openHRSystem.Serviced = true;
          returnBool = false;
        }
      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("DatabaseIsBeingServiced - {ex.Message}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }

      return returnBool;

    }

    private bool DatabaseIsRunningOvernight(ref OpenHRSystem openHRSystem, SqlConnection conn)
    {
      bool returnBool = false;

      try
      {
        returnBool = GetSystemSetting("database", "updatingdatedependantcolumns", conn) == "1";

        if (!openHRSystem.Suspended && returnBool)
        {
          RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) suspended.", MessageEventDetails.MessageEventType.WindowsEventsLog));
        }
        else if (openHRSystem.Suspended && !returnBool)
        {
          RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) resumed.", MessageEventDetails.MessageEventType.WindowsEventsLog));
        }

        openHRSystem.Suspended = returnBool;

      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("Overnight Job check ({openHRSystem}) - {ex.Message} {ex.StackTrace}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }

      return returnBool;
    }

    private bool DatabaseVersionIsOK(ref OpenHRSystem openHRSystem, SqlConnection conn)
    {
      bool returnBool;

      VersionNumber minSvcVersion = new VersionNumber();
      string minSvcVersionString = GetSystemSetting("outlook service 2", "minimum version", conn);

      try
      {
        int start = 0;
        int length = 0;
        start = 0;
        length = minSvcVersionString.IndexOf(".", start);
        minSvcVersion.Major = Convert.ToInt32(minSvcVersionString.Substring(start, length));

        start += length + 1;
        length = (minSvcVersionString.IndexOf(".", start)) - start;
        minSvcVersion.Minor = Convert.ToInt32(minSvcVersionString.Substring(start, length));

        start += length + 1;
        if (minSvcVersionString.IndexOf(".", start) == -1)
        {
          minSvcVersion.Build = Convert.ToInt32(minSvcVersionString.Substring(start));
        }
        else {
          length = (minSvcVersionString.IndexOf(".", start)) - start;
          minSvcVersion.Build = Convert.ToInt32(minSvcVersionString.Substring(start, length));
        }
      }
      catch
      {
        minSvcVersion.Major = 0;
        minSvcVersion.Minor = 0;
        minSvcVersion.Build = 0;
      }

      bool svcVersionOK = minSvcVersion.Major <= _svcVersion.Major && minSvcVersion.Minor <= _svcVersion.Minor && minSvcVersion.Build <= _svcVersion.Build;

      float dbVersion = Convert.ToSingle(GetSystemSetting("database", "version", conn));
      bool dbVersionOK = MINIMUMDBVERSION <= dbVersion;

      returnBool = (svcVersionOK && dbVersionOK);

      if ((openHRSystem.VersionOk && (!returnBool)))
      {
        RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) is incompatible with the OpenHR Outlook Calendar Service 2 ({_svcVersion}). Contact your system administrator.", EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }
      else if ((returnBool && (!openHRSystem.VersionOk)))
      {
        RaiseMessageEvent(new MessageEventDetails("Database ({openHRSystem}) version incompatibility corrected.", MessageEventDetails.MessageEventType.WindowsEventsLog));
      }

      openHRSystem.VersionOk = returnBool;

      return returnBool;

    }
    #endregion

    #region Private methods
    private void InitializeWorker()
    {
      ConfigurationErrors = new List<ConfigurationError>();
      Messages = new List<string>();
      _svcVersion = new VersionNumber
      {
        Major = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major,
        Minor = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor,
        Build = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build
      };
    }

    private bool OutlookBatch()
    {
      string dateFormat = string.Empty;
      string sql = string.Empty;
      Int32 linkID = 0;
      Int32 folderID = 0;
      Int32 startDateColumnID = 0;
      Int32 endDateColumnID = 0;
      Int32 startTimeColumnID = 0;
      Int32 endTimeColumnID = 0;
      Int32 recordID = 0;
      Int32 recordDescExprID = 0;
      Int32 subjectExprID = 0;
      string storeId = string.Empty;
      string entryId = string.Empty;
      string title = string.Empty;
      string fixedStartTime = string.Empty;
      string fixedEndTime = string.Empty;
      bool deleted = false;
      Int32 timeRange = 0;
      Int32 folderType = 0;
      Int32 folderExprID = 0;
      string folderPath = string.Empty;
      bool createdEntry = false;
      bool doOutlookOK = false;
      string folder = string.Empty;

      if (_userDefinedConfiguration.Debug)
      {
        RaiseMessageEvent(new MessageEventDetails("-------------" + Environment.NewLine + "Batch Started" + Environment.NewLine + "-------------"));
      }

      foreach (OpenHRSystem openHrSystem in _userDefinedConfiguration.OpenHRSystems)
      {
        try
        {
          RaiseMessageEvent(new MessageEventDetails("Opening Connection {openHrSystem.ConnectionString}"));

          using (SqlConnection sqlConnection = new SqlConnection(openHrSystem.ConnectionString))
          {
            sqlConnection.Open();
            RaiseMessageEvent(new MessageEventDetails("Connection Open"));

            if (DatabaseIsOk(openHrSystem, sqlConnection))
            {
              RaiseMessageEvent(new MessageEventDetails("Database is okay"));
              _dateFormat = GetSystemSetting("email", "date format", sqlConnection);

              sql = GetOutlookEventsSQL();
              var cmdEvents = new SqlCommand(sql, sqlConnection) { CommandType = CommandType.Text };

              var reader = cmdEvents.ExecuteReader();

              if (reader.HasRows)
              {
                RaiseMessageEvent(new MessageEventDetails("Checking Exchange"));
                if (!_loggedOn)
                  LogonToExchange();

                if (!_loggedOn)
                {
                  RaiseMessageEvent(new MessageEventDetails("Could not logon to Exchange: " + _errorMessage,
                    EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLog));
                  break;
                }

                // Already logged onto Outlook
                RaiseMessageEvent(new MessageEventDetails("Connected to Exchange"));

                while (reader.Read())
                {
                  storeId = Utilities.NullSafeString(reader["StoreID"]);
                  entryId = Utilities.NullSafeString(reader["EntryID"]);
                  deleted = Utilities.NullSafeBoolean(reader["Deleted"]);
                  linkID = Utilities.NullSafeInteger(reader["LinkID"]);
                  folderID = Utilities.NullSafeInteger(reader["FolderID"]);
                  recordID = Utilities.NullSafeInteger(reader["RecordID"]);
                  startDateColumnID = Utilities.NullSafeInteger(reader["StartDate"]);
                  endDateColumnID = Utilities.NullSafeInteger(reader["EndDate"]);
                  fixedStartTime = Utilities.NullSafeString(reader["FixedStartTime"]);
                  fixedEndTime = Utilities.NullSafeString(reader["FixedEndTime"]);
                  startTimeColumnID = Utilities.NullSafeInteger(reader["ColumnStartTime"]);
                  endTimeColumnID = Utilities.NullSafeInteger(reader["ColumnEndTime"]);
                  timeRange = Utilities.NullSafeInteger(reader["TimeRange"]);
                  title = Utilities.NullSafeString(reader["Title"]);
                  subjectExprID = Utilities.NullSafeInteger(reader["Subject"]);
                  recordDescExprID = Utilities.NullSafeInteger(reader["RecordDescExprID"]);
                  folderType = Utilities.NullSafeInteger(reader["FolderType"]);
                  folderPath = Utilities.NullSafeString(reader["FixedPath"]);
                  folderExprID = Utilities.NullSafeInteger(reader["ExprID"]);
                  _content = Utilities.NullSafeString(reader["Content"]);
                  _reminder = Utilities.NullSafeBoolean(reader["Reminder"]);
                  _reminderOffset = Utilities.NullSafeInteger(reader["ReminderOffset"]);
                  _reminderPeriod = Utilities.NullSafeInteger(reader["ReminderPeriod"]);
                  _busyStatus = Utilities.NullSafeInteger(reader["BusyStatus"]);

                  _storeId = string.Empty;
                  _entryId = string.Empty;
                  doOutlookOK = true;

                  bool emptyIDs = (storeId == string.Empty && entryId == string.Empty);

                  RaiseMessageEvent(new MessageEventDetails("Servicing {serverDBName}"));
                  RaiseMessageEvent(new MessageEventDetails(title));
                  RaiseMessageEvent(
                    new MessageEventDetails(
                      string.Format("Deleted: {0}{3}StoreID: {1}{3}EntryID: {2}", deleted, storeId, entryId, Environment.NewLine)));

                  if (deleted || (!emptyIDs))
                  {
                    _storeId = storeId;
                    _entryId = entryId;
                    SqlTransaction transaction;

                    using (SqlConnection connEvents = new SqlConnection(openHrSystem.ConnectionString))
                    {
                      connEvents.Open();

                      SqlCommand cmd = connEvents.CreateCommand();

                      // Start a local transaction
                      transaction = connEvents.BeginTransaction();

                      RaiseMessageEvent(new MessageEventDetails("BEGIN TRANS"));

                      // Must assign both transaction object and connection
                      // to Command object for a pending local transaction.
                      cmd.Connection = connEvents;
                      cmd.Transaction = transaction;
                      cmd.CommandTimeout = _userDefinedConfiguration.CommandTimeout;

                      try
                      {
                        if (!deleted)
                        {
                          RaiseMessageEvent(new MessageEventDetails("Update to NULLs"));

                          // Update to NULLs
                          sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) ";
                          sql += "SET StoreID = '' ";
                          sql += ", EntryID = '' ";
                          sql += ", RefreshDate = GETDATE() ";
                          sql += "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID";
                          cmd.Parameters.AddWithValue("@LinkID", linkID);
                          cmd.Parameters.AddWithValue("@FolderID", folderID);
                          cmd.Parameters.AddWithValue("@RecordID", recordID);
                          cmd.CommandText = sql;
                          cmd.ExecuteNonQuery();
                          cmd.Parameters.Clear();
                        }
                        else {
                          RaiseMessageEvent(new MessageEventDetails("Delete from ASRSysOutlookEvents"));

                          // Delete entry
                          sql = "DELETE FROM ASRSysOutlookEvents WITH(ROWLOCK) ";
                          sql += "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID";
                          cmd.Parameters.AddWithValue("@LinkID", linkID);
                          cmd.Parameters.AddWithValue("@FolderID", folderID);
                          cmd.Parameters.AddWithValue("@RecordID", recordID);
                          cmd.CommandText = sql;
                          cmd.ExecuteNonQuery();
                          cmd.Parameters.Clear();
                        }

                        if (DeleteEntry())
                        {
                          RaiseMessageEvent(new MessageEventDetails("Delete from Outlook" + Environment.NewLine + "COMMIT TRANS"));

                          // Attempt to commit the transaction.
                          transaction.Commit();
                          doOutlookOK = true;
                        }
                        else {
                          try
                          {
                            RaiseMessageEvent(
                              new MessageEventDetails("ROLLBACK TRANS: Delete from Outlook failed with '" + _errorMessage + "'"));

                            // Rollback
                            transaction.Rollback();
                          }
                          catch
                          {
                            // This catch block will handle any errors that may have occurred
                            // on the server that would cause the rollback to fail, such as a closed connection.
                            RaiseMessageEvent(new MessageEventDetails("Error rolling back transaction duplicates may occur.",
                              EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
                          }

                          doOutlookOK = false;
                          RaiseMessageEvent(new MessageEventDetails("Delete Error: " + _errorMessage, EventLogEntryType.Warning,
                            MessageEventDetails.MessageEventType.WindowsEventsLog));

                          // Update so we don't try again
                          using (SqlConnection connUpd = new SqlConnection(openHrSystem.ConnectionString))
                          {
                            connUpd.Open();

                            cmd = connEvents.CreateCommand();

                            RaiseMessageEvent(new MessageEventDetails("BEGIN TRANS: Update to not retry"));

                            // Begin Trans
                            transaction = connUpd.BeginTransaction();
                            cmd.Connection = connUpd;
                            cmd.Transaction = transaction;
                            cmd.CommandTimeout = _userDefinedConfiguration.CommandTimeout;

                            sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) ";
                            sql += "SET Refresh = 0 ";
                            sql += ", Deleted = 0 ";
                            sql += ", RefreshDate = GETDATE() ";
                            sql += "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID";
                            cmd.Parameters.AddWithValue("@LinkID", linkID);
                            cmd.Parameters.AddWithValue("@FolderID", folderID);
                            cmd.Parameters.AddWithValue("@RecordID", recordID);
                            cmd.CommandText = sql;
                            cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();

                            // Attempt to commit the transaction.
                            transaction.Commit();
                            RaiseMessageEvent(new MessageEventDetails("COMMIT TRANS: Update to not retry"));
                          }
                        }
                      }
                      catch (SqlException sqlEx)
                      {
                        doOutlookOK = false;
                        // Rollback
                        try
                        {
                          RaiseMessageEvent(new MessageEventDetails("ROLLBACK TRANS: Update or Delete failed"));
                          transaction.Rollback();
                        }
                        catch
                        {
                          // This catch block will handle any errors that may have occurred
                          // on the server that would cause the rollback to fail, such as a closed connection.
                          RaiseMessageEvent(new MessageEventDetails("Error rolling back transaction; duplicates may occur.",
                            EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
                        }

                        if (sqlEx.Number == 1205)
                        {
                          RaiseMessageEvent(new MessageEventDetails("*** DEADLOCK ***" + Environment.NewLine));

                          // Deadlock encountered - return false an wait for a minute
                          transaction.Dispose();
                          if (!reader.IsClosed)
                          {
                            reader.Close();
                          }
                          return false;
                        }
                        RaiseMessageEvent(new MessageEventDetails(string.Format("Sql Error in transaction: {0}.", sqlEx.Message), 
                          EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));

                      }
                      catch (Exception ex)
                      {
                        doOutlookOK = false;
                        RaiseMessageEvent(new MessageEventDetails(string.Format("Error in transaction: {0}.", ex.Message),
                          EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
                      }
                    }
                  }

                  RaiseMessageEvent(new MessageEventDetails("doOutlookOK: " + doOutlookOK));

                  if (doOutlookOK && !deleted)
                  {
                    using (SqlConnection sqlConnectionSp = new SqlConnection(openHrSystem.ConnectionString))
                    {
                      // Lets call spASRNetOutlookBatch to get us the rest of our properties
                      SqlCommand cmdSp = new SqlCommand("spASRNetOutlookBatch", sqlConnectionSp);
                      cmdSp.Connection.Open();
                      cmdSp.CommandType = CommandType.StoredProcedure;

                      // Input/Output parameters
                      cmdSp.Parameters.Add("@Content", SqlDbType.VarChar, 8000).Direction = ParameterDirection.InputOutput;
                      cmdSp.Parameters["@Content"].Value = _content;

                      // Output parameters
                      cmdSp.Parameters.Add("@AllDayEvent", SqlDbType.Bit).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@StartDate", SqlDbType.DateTime).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@EndDate", SqlDbType.DateTime).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@StartTime", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@EndTime", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@Subject", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output;
                      cmdSp.Parameters.Add("@Folder", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output;

                      // Input parameters
                      cmdSp.Parameters.AddWithValue("@LinkID", linkID);
                      cmdSp.Parameters.AddWithValue("@RecordID", recordID);
                      cmdSp.Parameters.AddWithValue("@FolderID", folderID);
                      cmdSp.Parameters.AddWithValue("@StartDateColumnID", startDateColumnID);
                      cmdSp.Parameters.AddWithValue("@EndDateColumnID", endDateColumnID);
                      cmdSp.Parameters.AddWithValue("@FixedStartTime", fixedStartTime);
                      cmdSp.Parameters.AddWithValue("@FixedEndTime", fixedEndTime);
                      cmdSp.Parameters.AddWithValue("@StartTimeColumnID", startTimeColumnID);
                      cmdSp.Parameters.AddWithValue("@EndTimeColumnID", endTimeColumnID);
                      cmdSp.Parameters.AddWithValue("@TimeRange", timeRange);
                      cmdSp.Parameters.AddWithValue("@Title", title);
                      cmdSp.Parameters.AddWithValue("@SubjectExprID", subjectExprID);
                      cmdSp.Parameters.AddWithValue("@RecordDescExprID", recordDescExprID);
                      cmdSp.Parameters.AddWithValue("@DateFormat", _dateFormat);
                      cmdSp.Parameters.AddWithValue("@FolderPath", folderPath);
                      cmdSp.Parameters.AddWithValue("@FolderType", folderType);
                      cmdSp.Parameters.AddWithValue("@FolderExprID", folderExprID);
                      cmdSp.ExecuteNonQuery();

                      // Retrieve output parameters
                      _content = Utilities.NullSafeString(cmdSp.Parameters["@Content"].Value);
                      _allDayEvent = Utilities.NullSafeBoolean(cmdSp.Parameters["@AllDayEvent"].Value);

                      if (!cmdSp.Parameters["@StartDate"].Value.Equals(DBNull.Value))
                        _startDate = Convert.ToDateTime(cmdSp.Parameters["@StartDate"].Value);

                      _endDate = !cmdSp.Parameters["@EndDate"].Value.Equals(DBNull.Value) ? Convert.ToDateTime(cmdSp.Parameters["@EndDate"].Value) : _startDate;

                      if (DateTime.Compare(_startDate, _endDate) > 0)
                        _endDate = _startDate;

                      _startTime = Utilities.NullSafeString(cmdSp.Parameters["@StartTime"].Value);
                      _endTime = Utilities.NullSafeString(cmdSp.Parameters["@EndTime"].Value);
                      _subject = Utilities.NullSafeString(cmdSp.Parameters["@Subject"].Value);
                      _folder = Utilities.NullSafeString(cmdSp.Parameters["@Folder"].Value);

                      cmdSp.Parameters.Clear();
                      cmdSp.Connection.Close();
                      cmdSp.Dispose();
                    }

                    // Create the outlook appointment using the OutlookCalendar class
                    if (!CreateEntry())
                    {
                      RaiseMessageEvent(new MessageEventDetails("Create Entry: FAILED"));
                      RaiseMessageEvent(

                        new MessageEventDetails(string.Format("Could not create entry {0} : {1}", _subject, _errorMessage),
                          EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLog));
                    }
                    else { //Entry created
                      RaiseMessageEvent(
                        new MessageEventDetails(string.Format("StoreID: {0} EntryId: {1}", _storeId, _entryId)));
                      RaiseMessageEvent(new MessageEventDetails(
                        "Create Entry: " + _subject + " [" + _startDate + " to " + _endDate + "]", EventLogEntryType.Information,
                        MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
                    }

                    try
                    {
                      RaiseMessageEvent(new MessageEventDetails("Update ASRSysOutlookEvents"));

                      // Need to update the row in the events table
                      using (SqlConnection connUpd = new SqlConnection(openHrSystem.ConnectionString))
                      {
                        sql = "UPDATE ASRSysOutlookEvents WITH(ROWLOCK) ";
                        sql += "SET ErrorMessage = @ErrorMessage ";
                        sql += ", StoreID = @StoreID ";
                        sql += ", EntryID = @EntryID ";
                        sql += ", Refresh = 0 ";
                        sql += ", StartDate = @StartDate ";
                        sql += ", Subject = @Subject ";
                        sql += ", Folder = @Folder ";
                        sql += ", EndDate = @EndDate ";
                        sql += ", RefreshDate = GETDATE() ";
                        sql += "WHERE LinkID = @LinkID and FolderID = @FolderID and RecordID = @RecordID";
                        connUpd.Open();
                        SqlCommand cmdUpd = new SqlCommand(sql, connUpd);
                        cmdUpd.Parameters.AddWithValue("@ErrorMessage", _errorMessage);
                        cmdUpd.Parameters.AddWithValue("@StoreID", _storeId);
                        cmdUpd.Parameters.AddWithValue("@EntryID", _entryId);
                        cmdUpd.Parameters.AddWithValue("@StartDate", _startDate);
                        cmdUpd.Parameters.AddWithValue("@Subject", _subject);
                        cmdUpd.Parameters.AddWithValue("@Folder", _folder);
                        cmdUpd.Parameters.AddWithValue("@EndDate", _endDate);
                        cmdUpd.Parameters.AddWithValue("@LinkID", linkID);
                        cmdUpd.Parameters.AddWithValue("@FolderID", folderID);
                        cmdUpd.Parameters.AddWithValue("@RecordID", recordID);
                        cmdUpd.ExecuteNonQuery();
                        cmdUpd.Parameters.Clear();
                        cmdUpd.Connection.Close();
                        cmdUpd.Dispose();
                      }

                    }
                    catch (SqlException sqlEx)
                    {
                      string errDesc = Convert.ToString((sqlEx.Number == 1205 ? "deadlock" : "error"));

                      RaiseMessageEvent(new MessageEventDetails("Update: FAILED"));

                      if (DeleteEntry())
                      {
                        RaiseMessageEvent(new MessageEventDetails("Entry deleted"));
                        RaiseMessageEvent(
                          new MessageEventDetails("Sql {errDesc} occured: {_subject} set to retry.",
                            EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLog));
                      }
                      else {
                        RaiseMessageEvent(new MessageEventDetails("FAILED to delete entry: " + _errorMessage));
                        RaiseMessageEvent(
                          new MessageEventDetails(
                            string.Format("Sql {0} occurred in UPDATE {1}{1}{2} could not be deleted duplicates may occur.", errDesc,
                              Environment.NewLine, _subject), EventLogEntryType.Error,
                            MessageEventDetails.MessageEventType.WindowsEventsLog));
                      }

                      if (sqlEx.Number == 1205)
                      {
                        RaiseMessageEvent(new MessageEventDetails("*** DEADLOCK ***" + Environment.NewLine));
                        if (!reader.IsClosed)
                        {
                          reader.Close();
                        }

                        return false;
                      }
                    }
                  }

                  RaiseMessageEvent(new MessageEventDetails(Environment.NewLine));

                }
              }
              else {
                // No events to process
              }
            }
            else {
              RaiseMessageEvent(new MessageEventDetails("Database not okay, check previous logged information",
                MessageEventDetails.MessageEventType.WindowsEventsLog));
            }
            sqlConnection.Close();
          }
        }
        catch (Exception ex)
        {
          RaiseMessageEvent(new MessageEventDetails(ex.Message));
        }
        finally
        {
          RaiseMessageEvent(new MessageEventDetails("-------------" + Environment.NewLine + "Batch Finished" + Environment.NewLine + "-------------"));
        }
      }

      return true;
    }

    private bool CreateEntry()
    {
      string mailboxName = "";

      if (!_loggedOn)
      {
        _errorMessage = "Exchange Logon failed";
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      if (_startDate.ToString() == string.Empty)
      {
        _errorMessage = "No start date entered";
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      if (!_allDayEvent)
      {
        if (!LikeOperator.LikeString(_startTime, "##:##", CompareMethod.Text))
        {
          _errorMessage = string.Concat("Invalid Start Time <", _startTime, ">");
          RaiseMessageEvent(new MessageEventDetails(_errorMessage));
          return false;
        }

        if (!LikeOperator.LikeString(_endTime, "##:##", CompareMethod.Text))
        {
          {
            _errorMessage = string.Concat("Invalid End Time <", _endTime, ">");
            RaiseMessageEvent(new MessageEventDetails(_errorMessage));
            return false;
          }
        }
      }
      if (_folder == string.Empty)
      {
        _errorMessage = "Outlook folder name empty";
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      string dateFormat = DateFormat();
      _errorMessage = string.Empty;

      Folder calendarFolder = null;

      if (!_folder.Contains("\\\\Public Folders"))
      {
        try
        {
          mailboxName = GetNameFromMailbox(_folder, true);
          RaiseMessageEvent(new MessageEventDetails("GetSharedMailbox: " + mailboxName));
          calendarFolder = Folder.Bind(_exchangeService, WellKnownFolderName.Calendar);
          RaiseMessageEvent(new MessageEventDetails("GetSharedMailbox OK"));
        }
        catch (Exception ex)
        {
          _errorMessage = string.Format("Unable to open mailbox for {0} - {1}", mailboxName, ex.Message);
          return false;
        }
      }
      else
      {
        RaiseMessageEvent(new MessageEventDetails("Public Folder: " + _folder));
        try
        {
          var rootFolder = Folder.Bind(_exchangeService, WellKnownFolderName.Root);
          string storeName = rootFolder.DisplayName;

          // Create a new folder view, and pass in the maximum number of folders to return.
          FolderView view = new FolderView(100);

          // Create an extended property definition for the PR_ATTR_HIDDEN property,
          // so that your results will indicate whether the folder is a hidden folder.
          ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);

          // As a best practice, limit the properties returned to only those required.
          // In this case, return the folder ID, DisplayName, and the value of the isHiddenProp extended property.
          view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, isHiddenProp);

          // Indicate a Traversal value of Deep, so that all subfolders are retrieved.
          view.Traversal = FolderTraversal.Deep;

          // Call FindFolders to retrieve the folder hierarchy, starting with the MsgFolderRoot folder.
          // This method call results in a FindFolder call to EWS.
          FindFoldersResults findFolderResults = _exchangeService.FindFolders(WellKnownFolderName.Root, view);

          foreach (Folder f in findFolderResults.Folders)
          {
            //Debugger.Break();
          }


          //  foreach (RDOStore store in _session.Stores) {
          //	if ((store.StoreKind == TxStoreKind.skPublicFolders)) {
          //	  folderItem = GetFolderFromPath(_folder, store.IPMRootFolder.Folders);

          //	  if (folderItem != null) {
          //		RaiseMessageEvent(new MessageEventDetails("GetFolderFromPath OK"));
          //		break; // TODO: might not be correct. Was : Exit For
          //	  } else {
          //		RaiseMessageEvent(new MessageEventDetails("GetFolderFromPath - Cant find folder"));
          //	  }
          //	}
          //		      }

        }
        catch
        {
          RaiseMessageEvent(new MessageEventDetails("Public Folder: NO ACCESS TO STORE " + _folder));
        }
      }

      if (calendarFolder == null)
      {
        _errorMessage = "Unable to obtain a valid Calendar path from: " + _folder;
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      if (_errorMessage != string.Empty)
      {
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      try
      {
        var appointment = new Appointment(_exchangeService)
        {
          Subject = _subject,
          Body = _content,
          IsAllDayEvent = _allDayEvent
        };

        if (_allDayEvent)
        {
          appointment.Start = _startDate;
          //If no times are specified then outlook correctly finishes at midnight but does not include the end date. 
          //For OpenHR we need the event to be inclusive of both the start date and end date so if its an all day
          //event add one day to the end date.
          if (DateTime.Compare(_startDate, _endDate) > 0)
            appointment.End = _startDate.AddDays(1); //Start date after end date
          else
            appointment.End = _endDate.AddDays(1);
        }
        else {
          _startDate = Convert.ToDateTime(_startDate.Date.ToString(dateFormat) + " " + _startTime);
          appointment.Start = _startDate;
          _endDate = Convert.ToDateTime(_endDate.ToString(dateFormat) + " " + _endTime);
          appointment.End = _endDate;
        }

        appointment.LegacyFreeBusyStatus = (LegacyFreeBusyStatus)_busyStatus;
        appointment.IsReminderSet = _reminder;
        var reminderInMinutes = Convert.ToInt32(_reminderOffset) *
                    Convert.ToInt32(Interaction.Choose(_reminderPeriod + 1, 1, 1440, 10080, 40240));
        if (_reminder)
        {
          appointment.ReminderMinutesBeforeStart = reminderInMinutes;
        }

        appointment.Save(new FolderId(WellKnownFolderName.Calendar, new Mailbox(mailboxName)));

        _entryId = appointment.Id.ToString();
      }
      catch (Exception ex)
      {
        _errorMessage = ex.Message;
        RaiseMessageEvent(new MessageEventDetails(_errorMessage));
        return false;
      }

      return true;
    }

    private bool DeleteEntry()
    {
      try
      {
        if (_entryId == string.Empty)
          return true;
        var appointment = Appointment.Bind(_exchangeService, new ItemId(_entryId));
        appointment.Delete(DeleteMode.HardDelete);
      }
      catch (Exception e)
      {
        RaiseMessageEvent(new MessageEventDetails("Cannot delete entry - Check Mailbox and Calendar folder permissions - " + e.Message));
        _errorMessage = "Cannot delete entry - Check Mailbox and Calendar folder permissions - " + e.Message;
        return false;
      }

      return true;
    }

    private void LogonToExchange()
    {
      RaiseMessageEvent(new MessageEventDetails("Attempting to autodiscover Exchange"));
      _loggedOn = false;

      try
      {
        ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

        _exchangeService = new ExchangeService(ExchangeVersion.Exchange2013)
        {
          UseDefaultCredentials = _useDefaultCredentials,
          TraceListener = new ExchangeTraceListener(),
          TraceEnabled = _enableExchangeTrace,
          TraceFlags = _enableExchangeTrace ? TraceFlags.All : TraceFlags.None
        };

        if (!_useDefaultCredentials) //Need to provide username and password
          _exchangeService.Credentials = new NetworkCredential(_exchangeUser, _exchangeUserPassword, "");

        if (_userDefinedConfiguration.ExchangeServerURL == null)
        {
          _exchangeService.AutodiscoverUrl(_exchangeUser, RedirectionUrlValidationCallback);
        }
        else
        {
          _exchangeService.Url = new Uri(_userDefinedConfiguration.ExchangeServerURL);
        }

        try
        {
          RaiseMessageEvent(
            new MessageEventDetails(
            "Successfully logged in to Exchange Server {_exchangeService.RequestedServerVersion}; autodiscovery Url: {_exchangeService.Url}",
            EventLogEntryType.Information, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
        }
        catch
        {
          RaiseMessageEvent(
          new MessageEventDetails(
            "Successfully logged in to Exchange Server; no Server details available",
            EventLogEntryType.Information, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
        }

        _loggedOn = true;
      }
      catch (Exception e)
      {
        _errorMessage = "Error connecting to Exchange: " + e.Message;
        RaiseMessageEvent(new MessageEventDetails(_errorMessage, EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
      }
    }

    private string GetOutlookEventsSQL()
    {
      var sqlString = new StringBuilder();

      sqlString.Append("SELECT ASRSysOutlookEvents.LinkID,");
      sqlString.Append(" ASRSysOutlookEvents.FolderID,");
      sqlString.Append(" ASRSysOutlookEvents.TableID,");
      sqlString.Append(" ASRSysOutlookEvents.RecordID,");
      sqlString.Append(" ASRSysOutlookEvents.Refresh,");
      sqlString.Append(" ASRSysOutlookEvents.Deleted,");
      sqlString.Append(" ASRSysOutlookEvents.StoreID,");
      sqlString.Append(" ASRSysOutlookEvents.EntryID,");
      sqlString.Append(" ASRSysOutlookLinks.Title,");
      sqlString.Append(" ASRSysOutlookLinks.BusyStatus,");
      sqlString.Append(" ASRSysOutlookLinks.StartDate,");
      sqlString.Append(" ASRSysOutlookLinks.EndDate,");
      sqlString.Append(" ASRSysOutlookLinks.TimeRange,");
      sqlString.Append(" ASRSysOutlookLinks.FixedStartTime,");
      sqlString.Append(" ASRSysOutlookLinks.FixedEndTime,");
      sqlString.Append(" ASRSysOutlookLinks.ColumnStartTime,");
      sqlString.Append(" ASRSysOutlookLinks.ColumnEndTime,");
      sqlString.Append(" ASRSysOutlookLinks.Subject,");
      sqlString.Append(" ISNULL(ASRSysOutlookLinks.Content,'') [Content],");
      sqlString.Append(" ASRSysOutlookLinks.Reminder,");
      sqlString.Append(" ASRSysOutlookLinks.ReminderOffset,");
      sqlString.Append(" ASRSysOutlookLinks.ReminderPeriod,");
      sqlString.Append(" ASRSysOutlookFolders.FolderType,");
      sqlString.Append(" ASRSysOutlookFolders.FixedPath,");
      sqlString.Append(" ASRSysOutlookFolders.ExprID,");
      sqlString.Append(" ASRSysTables.RecordDescExprID ");
      sqlString.Append("FROM ASRSysOutlookEvents WITH(READPAST) ");
      sqlString.Append("LEFT OUTER JOIN ASRSysOutlookLinks WITH(READPAST)");
      sqlString.Append(" ON ASRSysOutlookEvents.LinkID = ASRSysOutlookLinks.LinkID ");
      sqlString.Append("LEFT OUTER JOIN ASRSysOutlookFolders WITH(READPAST)");
      sqlString.Append(" ON ASRSysOutlookEvents.FolderID = ASRSysOutlookFolders.FolderID ");
      sqlString.Append("LEFT OUTER JOIN ASRSysTables WITH(READPAST)");
      sqlString.Append(" ON ASRSysOutlookEvents.TableID = ASRSysTables.TableID ");
      sqlString.Append("WHERE(ASRSysOutlookEvents.Refresh = 1)");
      sqlString.Append(" OR ASRSysOutlookEvents.Deleted = 1 ");
      sqlString.Append("ORDER BY ASRSysOutlookEvents.RecordID");

      return sqlString.ToString();
    }

    private string GetSystemSetting(string section, string key, SqlConnection conn)
    {
      string returnString = string.Empty;

      try
      {
        using (SqlCommand cmd = new SqlCommand())
        {
          cmd.CommandText = "spASRGetSetting";
          cmd.Connection = conn;
          cmd.CommandType = CommandType.StoredProcedure;
          cmd.CommandTimeout = _userDefinedConfiguration.CommandTimeout;

          cmd.Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psSection"].Value = section;

          cmd.Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psKey"].Value = key;

          cmd.Parameters.Add("@psDefault", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psDefault"].Value = "0";

          cmd.Parameters.Add("@pfUserSetting", SqlDbType.Bit).Direction = ParameterDirection.Input;
          cmd.Parameters["@pfUserSetting"].Value = false;

          cmd.Parameters.Add("@psResult", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output;
          cmd.ExecuteNonQuery();

          returnString = cmd.Parameters["@psResult"].Value.ToString();
        }

      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("GetSystemSetting - {ex.Message} {ex.StackTrace}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }

      return returnString;
    }

    private void SaveSystemSetting(string section, string key, string value, SqlConnection conn)
    {
      try
      {
        using (SqlCommand cmd = new SqlCommand())
        {
          cmd.CommandText = "spASRSaveSetting";
          cmd.Connection = conn;
          cmd.CommandType = CommandType.StoredProcedure;
          cmd.CommandTimeout = _userDefinedConfiguration.CommandTimeout;

          cmd.Parameters.Add("@psSection", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psSection"].Value = section;

          cmd.Parameters.Add("@psKey", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psKey"].Value = key;

          cmd.Parameters.Add("@psValue", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Input;
          cmd.Parameters["@psValue"].Value = value;

          cmd.ExecuteNonQuery();
        }
      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("SaveSystemSetting - {ex.Message} {ex.StackTrace}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLog));
      }
    }

    private string GetNameFromMailbox(string mailboxPath, bool resolveName)
    {
      string tmpName = mailboxPath;
      try
      {
        if (mailboxPath.Contains("\\\\Mailbox")) {
          tmpName = mailboxPath.Replace("\\\\Mailbox - ", "");
          tmpName = tmpName.Substring(0, tmpName.IndexOf("\\"));
        }
        else if (mailboxPath.StartsWith("\\\\")) {
          tmpName = mailboxPath.Substring(mailboxPath.IndexOf("\\\\") + 2);
          tmpName = tmpName.Substring(0, tmpName.IndexOf("\\"));
        }

        if (!tmpName.Contains("@") && resolveName) {
          var resolved = _exchangeService.ResolveName(tmpName);
          if (resolved.Count == 1) {
            tmpName = resolved[0].Mailbox.Address;
            }
        }

      }
      catch
      {
        return string.Empty;
      }

      return tmpName;
    }

    private string DateFormat()
    {
      // NB. Windows allows the user to configure totally stupid
      // date formats (eg. d/M/yyMydy !). This function does not cater
      // for such stupidity, and simply takes the first occurence of the
      // 'd', 'M', 'y' characters.
      string sysFormat = string.Empty;
      string sysDateSeparator = string.Empty;
      string sysDateFormat = string.Empty;
      bool daysDone = false;
      bool monthsDone = false;
      bool yearsDone = false;

      sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
      sysDateSeparator = CultureInfo.CurrentCulture.DateTimeFormat.DateSeparator;

      // Loop through the string picking out the required characters.

      for (int iLoop = 0; iLoop <= sysFormat.Length - 1; iLoop++)
      {
        switch (sysFormat.Substring(iLoop, 1))
        {
          case "d":
            if (!daysDone)
            {
              // Ensure we have two day characters.
              sysDateFormat += "dd";
              daysDone = true;
            }

            break;
          case "M":
            if (!monthsDone)
            {
              // Ensure we have two month characters.
              sysDateFormat += "MM";
              monthsDone = true;
            }

            break;
          case "y":
            if (!yearsDone)
            {
              // Ensure we have four year characters.
              sysDateFormat += "yyyy";
              yearsDone = true;
            }

            break;
          default:
            sysDateFormat += sysFormat.Substring(iLoop, 1);
            break;
        }

      }

      // Ensure that all day, month and year parts of the date
      // are present in the format.
      if (!daysDone)
      {
        if (sysDateFormat.Substring(sysDateFormat.Length - 1, 1) != sysDateSeparator)
        {
          sysDateFormat += sysDateSeparator;
        }

        sysDateFormat += "dd";
      }

      if (!monthsDone)
      {
        if (sysDateFormat.Substring(sysDateFormat.Length - 1, 1) != sysDateSeparator)
        {
          sysDateFormat += sysDateSeparator;
        }

        sysDateFormat += "MM";
      }

      if (!yearsDone)
      {
        if (sysDateFormat.Substring(sysDateFormat.Length - 1, 1) != sysDateSeparator)
        {
          sysDateFormat += sysDateSeparator;
        }

        sysDateFormat += "yyyy";
      }

      // Return the date format.
      return sysDateFormat;
    }

    #endregion

    #region Exchange server test methods
    public void TestAutodiscoverLogon()
    {
      LogonToExchange();
    }

    public void CreateTestCalendarEntry()
    {
      try
      {
        LogonToExchange();

        var appointment = new Appointment(_exchangeService)
        {
          Subject = "OpenHR Outlook Calendar Service test",
          Body = "OpenHR Outlook Calendar Service test",
          Start = DateTime.Now.AddHours(1),
          IsAllDayEvent = false,
          IsReminderSet = false,
          LegacyFreeBusyStatus = LegacyFreeBusyStatus.Busy
        };
        appointment.End = appointment.Start.AddHours(4);

        appointment.Save(new FolderId(WellKnownFolderName.Calendar, new Mailbox(_exchangeUser)), SendInvitationsMode.SendToNone);

        RaiseMessageEvent(new MessageEventDetails("Created test calendar entry successfully"));
      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("Create test calendar entry failed: " + ex.Message));
      }
    }
    #endregion

    public void OutputUserConfiguration()
    {
      //Output the user configuration values
      RaiseMessageEvent(new MessageEventDetails("OpenHRUser: " + _userDefinedConfiguration.OpenHRUser));
      RaiseMessageEvent(new MessageEventDetails("OpenHRPassword: " + _userDefinedConfiguration.OpenHRPassword));
      RaiseMessageEvent(new MessageEventDetails("ExchangeServer: " + _userDefinedConfiguration.ExchangeServer));
      RaiseMessageEvent(new MessageEventDetails("ServiceAccountPassword: *****"));
      RaiseMessageEvent(new MessageEventDetails("Debug: " + _userDefinedConfiguration.Debug));
    }

    public void DescribeServicedSystems()
    {
      foreach (OpenHRSystem ohrs in _userDefinedConfiguration.OpenHRSystems)
      {
        if (ohrs.DatabaseName == string.Empty)
          RaiseMessageEvent(
            new MessageEventDetails(
              "No database parameter defined for the server ({ohrs.ServerName}). System will be ignored.",
              EventLogEntryType.Warning, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
        else
          RaiseMessageEvent(
            new MessageEventDetails(
              "OpenHR Outlook Calendar Service 2 configured for {ohrs.ServerName}\\{ohrs.DatabaseName}. Exchange Server = {_userDefinedConfiguration.ExchangeServer}", EventLogEntryType.Information, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));

      }
    }

    public void Stop()
    {
      foreach (OpenHRSystem ohrs in _userDefinedConfiguration.OpenHRSystems)
      {
        if (ohrs.Serviced)
        {
          try
          {
            using (SqlConnection sqlConnection = new SqlConnection(ohrs.ConnectionString))
            {
              sqlConnection.Open();
              SaveSystemSetting("outlook service 2", "running", "0", sqlConnection);
              sqlConnection.Close();
            }
          }
          catch (Exception ex)
          {
            RaiseMessageEvent(
              new MessageEventDetails(
                "Error clearing system parameters for database ({ohrs.DatabaseName}). {Environment.NewLine}{ex.Message}", EventLogEntryType.Error, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
          }
        }
      }
      RaiseMessageEvent(
        new MessageEventDetails(
          "OpenHR Outlook Calendar Service 2 ({_userDefinedConfiguration.ExchangeServer}) stopped successfully.", EventLogEntryType.Information, MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog));
    }


    public void CreateTestCalendarAndPopulate()
    {

      LogonToExchange();

      try
      {
        var appointment = new Appointment(_exchangeService)
        {
          Subject = "OpenHR",
          Body = _content,
          IsAllDayEvent = _allDayEvent
        };



        if (_allDayEvent)
        {
          appointment.Start = _startDate;
          //If no times are specified then outlook correctly finishes at midnight but does not include the end date. 
          //For OpenHR we need the event to be inclusive of both the start date and end date so if its an all day
          //event add one day to the end date.
          if (DateTime.Compare(_startDate, _endDate) > 0)
            appointment.End = _startDate.AddDays(1); //Start date after end date
          else
            appointment.End = _endDate.AddDays(1);
        }
        else
        {
          var dateFormat = "103";
          _startTime = "2016-02-05 09:00:00.000";
          _endTime = "2016-02-05 17:30:00.000";


          _startDate = Convert.ToDateTime(_startTime);
          appointment.Start = _startDate;
          _endDate = Convert.ToDateTime(_endTime);
          appointment.End = _endDate;
        }

        appointment.LegacyFreeBusyStatus = (LegacyFreeBusyStatus)_busyStatus;
        appointment.IsReminderSet = _reminder;
        var reminderInMinutes = Convert.ToInt32(_reminderOffset) *
                                Convert.ToInt32(Interaction.Choose(_reminderPeriod + 1, 1, 1440, 10080, 40240));
        if (_reminder)
        {
          appointment.ReminderMinutesBeforeStart = reminderInMinutes;
        }

        appointment.Save(new FolderId(WellKnownFolderName.Calendar, new Mailbox(_exchangeUser)));

        //_exchangeService.LoadPropertiesForItems(appointment, PropertySet.FirstClassProperties);
        //     var calendarView = new CalendarView();
        //   var getProps = _exchangeService.FindAppointments(appointment.ParentFolderId, calendarView);



        _storeId = "unused"; //Encoding.UTF8.GetString(appointment.StoreEntryId);
        _entryId = appointment.Id.ToString();

        appointment.Delete(DeleteMode.HardDelete);

        RaiseMessageEvent(
          new MessageEventDetails(string.Format("update success calendar entry success: {0} - {1} ", _storeId,
            _entryId)));
      }


      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("updated test calendar entry failed: " + ex.Message));
      }
    }

    public void CreateTestCalendarDelete(string entryID)
    {
      try
      {
        LogonToExchange();

        var appointment = Appointment.Bind(_exchangeService, new ItemId(entryID));
        appointment.Delete(DeleteMode.HardDelete);

        RaiseMessageEvent(new MessageEventDetails("deleted test calendar entry successfully"));
      }
      catch (Exception ex)
      {
        RaiseMessageEvent(new MessageEventDetails("deleted test calendar entry failed: " + ex.Message));
      }

    }
  }
}
