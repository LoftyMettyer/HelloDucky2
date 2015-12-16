using System;
using System.Collections.Generic;
using System.Configuration;
using OutlookCalendarLogic.Structures;

namespace OutlookCalendarLogic {
  public static class Utilities {
	public static string NullSafeString(object arg) {
	  try { return (string)arg; } catch (Exception) { return string.Empty; }
	}

	public static bool NullSafeBoolean(object arg) {
	  try { return (bool)arg; } catch (Exception) { return false; }
	}

	public static int NullSafeInteger(object arg) {
	  try { return (int)arg; } catch (Exception) { return 0; }
	}
  }
}
