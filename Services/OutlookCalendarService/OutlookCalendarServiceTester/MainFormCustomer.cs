using System;
using System.IO;
using System.Windows.Forms;
using OutlookCalendarLogic;
using static OutlookCalendarService2Tester.Utils;

namespace OutlookCalendarService2Tester
{
  public partial class MainFormCustomer : Form
  {
	private Worker _worker;
	private readonly StreamWriter _traceWriter;
	private const string DebugFile = "\\OutlookCalendarService2_Customer.txt";

	public MainFormCustomer()
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

	private void btnClearDebugMessages_Click(object sender, EventArgs e)
	{
	  txtDebugLog.Clear();
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

	private void HandleRaisedMessageEvent(MessageEventDetails e)
	{
	  WriteMessageToTextBox(txtDebugLog, e.Message);
	  _traceWriter.WriteLine(e.Message);
	  _traceWriter.Flush();
	}

	private void MainFormCustomer_FormClosed(object sender, FormClosedEventArgs e)
	{
	  _traceWriter.WriteLine("-----------------------------------------------------------------------------------");
	  _traceWriter.Flush();
	  _traceWriter.Close();
	}
  }
}
