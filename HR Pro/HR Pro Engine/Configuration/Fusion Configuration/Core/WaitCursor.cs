using System;
using System.Windows.Forms;

namespace Fusion
{
	public class WaitCursor: IDisposable
	{
		private readonly Cursor _previous;

		public WaitCursor()
		{
			_previous = Cursor.Current;
			Cursor.Current = Cursors.WaitCursor;
		}

		public void Dispose()
		{
			Cursor.Current = _previous;
		}
	}
}
