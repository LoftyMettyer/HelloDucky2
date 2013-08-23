namespace RCVS.Interfaces
{
	public interface iModel
	{
		long UserID { get; set; }
		void Load();
		void Save();
	}
}