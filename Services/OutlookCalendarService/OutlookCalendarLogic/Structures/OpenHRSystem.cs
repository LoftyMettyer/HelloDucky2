using System.Data.SqlClient;
using System.Reflection;

namespace OutlookCalendarLogic.Structures {
  public struct OpenHRSystem {
	public string ServerName;
	public string DatabaseName;
	//public bool DefaultServer;
	public bool Locked;
	public bool Suspended;
	public bool Serviced;
	public string ServiceServer;
	public bool VersionOk;
	public string UserName;
	public string Password;

	public OpenHRSystem(string server, string database) : this() {
	  ServerName = server;
	  DatabaseName = database;
	}

	/// <summary>
	/// Gets the ConnectionString for the given OpenHR System
	/// </summary>
	public string ConnectionString {
	  get {
		SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder {
		  DataSource = ServerName,
		  InitialCatalog = DatabaseName
		};

		if (!UserName.Equals(string.Empty)) {
		  builder.UserID = UserName;
		  builder.Password = Password;
		} else {
		  builder.IntegratedSecurity = true;
		}

		return builder.ConnectionString;
	  }
	}

	/// <summary>
	/// String representation of OpenHR Server in format ServerName.DatabaseName
	/// </summary>
	/// <returns></returns>
	public override string ToString() {
	  return string.Concat(ServerName, ".", DatabaseName);
	}
  }
}
