using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarLogic.Structures {
  public class VersionNumber {
	public int Major;
	public int Minor;

	public int Build;
	public override string ToString() {
	  if (Major == 0 && Minor == 0 && Build == 0)
		return "<unknown>";

	  //  Return String.Format("v{0}.{1}.{2}", Major, Minor, Build)
	  return string.Format("{0}.{1}.{2}", Major, Minor, Build);
	}
  }
}
