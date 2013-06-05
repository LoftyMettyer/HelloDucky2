using System;
using System.Data.SqlClient;

namespace Fusion
{
	public enum LockLevel
	{
		ReadWrite = 3,
		Manual = 2,
		Saving = 1
	}

	public class Database : IDisposable
	{
		//wouldn't normally keep an open live connection to db but the locking system 
		//relies on having an open connection to track live locks
		private readonly string _connectionString;
		private readonly SqlConnection _connection;

		public Database(string connectionString)
		{
			_connectionString = connectionString;
			_connection = new SqlConnection(_connectionString);
			try {
				_connection.Open();
			}
			catch {
				_connection = null;
			}
		}

		public string ConnectionString
		{
			get { return _connectionString; }
		}

		public bool IsValid()
		{
			return _connection != null;
		}

		public bool IsAdmin()
		{
			using (var cmd = new SqlCommand("SELECT IS_SRVROLEMEMBER('sysadmin')", _connection)) {
				return (int) cmd.ExecuteScalar() == 1;
			}
		}

		public bool Lock(LockLevel level)
		{
			if (level == LockLevel.ReadWrite) {
				//cant take a read-write lock if there are any locks
				using (var cmd = new SqlCommand("exec sp_ASRLockCheck", _connection)) {
					using (var dr = cmd.ExecuteReader()) {
						if (dr.HasRows)
							return false;
					}
				}
			} else if (level == LockLevel.Saving) {
				//already have a read-write lock, can only upgrade to a save lock if no users are in the system
				using (var cmd = new SqlCommand("exec spASRGetCurrentUsers", _connection)) {
					using (var dr = cmd.ExecuteReader()) {
						//two users will be me, one for each lock, so a third means someone else
						if (dr.Read() && dr.Read() && dr.Read())
							return false;
					}
				}
			}

			//take the lock
			using (var cmd = new SqlCommand("exec sp_ASRLockWrite " + (int) level, _connection)) {
				cmd.ExecuteNonQuery();
			}

			return true;
		}

		public void Unlock(LockLevel level)
		{
			using (var cmd = new SqlCommand("exec sp_ASRLockDelete " + (int) level, _connection)) {
				cmd.ExecuteNonQuery();
			}
		}

		public void Dispose()
		{
			_connection.Dispose();
		}
	}
}