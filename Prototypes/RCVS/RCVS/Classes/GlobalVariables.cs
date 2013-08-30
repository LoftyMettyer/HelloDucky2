using System.Web.Hosting;

namespace RCVS.Classes
{
	public static class GlobalVariables
	{
		public static string LogFileFullPath
		{
			get { return HostingEnvironment.MapPath("/Logs/WebServiceCalls.txt"); }
		}
	}
}