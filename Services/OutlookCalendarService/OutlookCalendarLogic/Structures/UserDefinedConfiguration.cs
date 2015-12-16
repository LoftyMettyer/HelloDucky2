using System.Collections.Generic;

namespace OutlookCalendarLogic.Structures {
  public struct UserDefinedConfiguration {
	public string OpenHRUser { get; set; }
	public string OpenHRPassword { get; set; }
	public string ExchangeServer { get; set; }
	public List<OpenHRSystem> OpenHRSystems { get; set; }
	public string ServiceAccountPassword { get; set; }
	public bool Debug { get; set; }
	public int CommandTimeout { get; set; }
  }
}
