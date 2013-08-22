namespace RCVS.Models
{
	public abstract class BaseModel
	{
		public long UserID { get; set; }
		public abstract void Load();
		public abstract void Save();
	}
}