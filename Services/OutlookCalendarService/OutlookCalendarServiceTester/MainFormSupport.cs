using System;
using System.IO;
using System.Windows.Forms;
using OutlookCalendarLogic;
using static OutlookCalendarService2Tester.Utils;

namespace OutlookCalendarService2Tester
{
  public partial class MainFormSupport : Form
  {
	private Worker _worker;
	private readonly StreamWriter _traceWriter;
	private const string DebugFile = "\\OutlookCalendarService2_Support.txt";

	public MainFormSupport()
	{
	  InitializeComponent();
	  try
	  {
		_traceWriter = File.AppendText(Directory.GetCurrentDirectory() + DebugFile);
	  } catch (Exception ex)
	  {
		WriteMessageToTextBox(txtDebugLog, ex.Message);
	  }
	}

	private void btnRunOnce_Click(object sender, EventArgs e)
	{
	  try
	  {
		WriteMessageToTextBox(txtDebugLog, "*** Started ***");
		_worker = new Worker(txtExchangeUser.Text, txtExchangeUserPassword.Text, true);
		_worker.RaiseMessageEvent += HandleRaisedMessageEvent;

		if (!_worker.GetAndCheckUserConfiguration())
		  return;

		_worker.OutputUserConfiguration();
		_worker.DescribeServicedSystems();
		_worker.ProcessEntries();

	  } catch (Exception ex)
	  {
		WriteMessageToTextBox(txtDebugLog, ex.Message);
	  }

	  WriteMessageToTextBox(txtDebugLog, "*** Finished ***");
	  _worker.Stop();
	}

	private void btnTestAutodiscoverLogon_Click(object sender, EventArgs e)
	{
	  WriteMessageToTextBox(txtDebugLog, "*** Autodiscover Logon ***");
	  _worker = new Worker(txtExchangeUser.Text, txtExchangeUserPassword.Text, true);
	  _worker.RaiseMessageEvent += HandleRaisedMessageEvent;

	  if (!_worker.GetAndCheckUserConfiguration())
		return;

	  _worker.OutputUserConfiguration();
	  _worker.DescribeServicedSystems();
	  _worker.TestAutodiscoverLogon();
	  WriteMessageToTextBox(txtDebugLog, "*** End of Autodiscover Logon ***");
	  _worker.Stop();
	}

	private void btnCreateTestCalendarEntry_Click(object sender, EventArgs e)
	{
	  WriteMessageToTextBox(txtDebugLog, "*** Create test calendar entry ***");
	  _worker = new Worker(txtExchangeUser.Text, txtExchangeUserPassword.Text, true);
	  _worker.RaiseMessageEvent += HandleRaisedMessageEvent;

	  if (!_worker.GetAndCheckUserConfiguration())
		return;

	  _worker.OutputUserConfiguration();
	  _worker.DescribeServicedSystems();
	  _worker.CreateTestCalendarEntry();
	  WriteMessageToTextBox(txtDebugLog, "*** End of Create test calendar entry ***");
	  _worker.Stop();
	}

	private void btnCheckConfiguration_Click(object sender, EventArgs e)
	{
	  WriteMessageToTextBox(txtDebugLog, "*** Check configuration ***");
	  _worker = new Worker(txtExchangeUser.Text, txtExchangeUserPassword.Text, true);
	  _worker.RaiseMessageEvent += HandleRaisedMessageEvent;
	  var configOk = _worker.GetAndCheckUserConfiguration();
	  _worker.OutputUserConfiguration();
	  WriteMessageToTextBox(txtDebugLog, "Is configuration Ok?: " + configOk);
	  WriteMessageToTextBox(txtDebugLog, "*** End of Check configuration ***");
	  _worker.DescribeServicedSystems();
	  WriteMessageToTextBox(txtDebugLog, Environment.NewLine);
	  _worker.Stop();
	}

	#region
	private void HandleRaisedMessageEvent(MessageEventDetails e)
	{
	  if (e.TriggeredEventType == MessageEventDetails.MessageEventType.DebugLog || e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog)
	  {
		WriteMessageToTextBox(txtDebugLog, e.Message);
		_traceWriter.WriteLine(e.Message);
		_traceWriter.Flush();
	  }
	  if (e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLog || e.TriggeredEventType == MessageEventDetails.MessageEventType.WindowsEventsLogAndDebugLog)
	  {
		WriteMessageToTextBox(txtWindowsEventsLog, e.Message);
	  }
	}

	private void btnClearDebugMessages_Click(object sender, EventArgs e)
	{
	  txtDebugLog.Clear();
	}

	private void btnClearWindowsEventLog_Click(object sender, EventArgs e)
	{
	  txtWindowsEventsLog.Clear();
	}
	#endregion

	private void MainFormSupport_FormClosed(object sender, FormClosedEventArgs e)
	{
	  _traceWriter.WriteLine("-----------------------------------------------------------------------------------");
	  _traceWriter.Flush();
	  _traceWriter.Close();
	}
  }
}
