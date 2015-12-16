using System;
using System.Windows.Forms;

namespace OutlookCalendarService2Tester
{
  public static class Utils
  {
	public static void WriteMessageToTextBox(TextBox textbox, string message)
	{
	  textbox.AppendText(message + Environment.NewLine);
	}
  }
}
