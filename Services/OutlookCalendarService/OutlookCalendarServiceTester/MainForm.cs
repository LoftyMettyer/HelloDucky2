using System;
using System.Windows.Forms;

namespace OutlookCalendarService2Tester
{
  public partial class MainForm : Form
  {
	public MainForm()
	{
	  InitializeComponent();
	}

	private void btnSupport_Click(object sender, EventArgs e)
	{
	  var today = DateTime.Now;
	  var todaysPassword =
		  string.Format(today.Day.ToString(), "00") +
		  string.Format(today.Month.ToString(), "00") +
		  string.Format((today.Day + 10).ToString(), "00") +
		  string.Format((today.Month + 10).ToString(), "00");

	  if (txtSupportPassword.Text == todaysPassword)
	  {
		var supportForm = new MainFormSupport();
		supportForm.ShowDialog();
	  } else
	  {
		MessageBox.Show("Wrong support password, please try again");
		txtSupportPassword.Focus();
		txtSupportPassword.SelectAll();
	  }
	}

	private void btnCustomer_Click(object sender, EventArgs e)
	{
	  var customerForm = new MainFormCustomer();
	  customerForm.ShowDialog();
	}

	private void txtSupportPassword_KeyPress(object sender, KeyPressEventArgs e)
	{
	  if (e.KeyChar == 13) //Enter
		btnSupport_Click(new object(), new EventArgs());
	}
  }
}
