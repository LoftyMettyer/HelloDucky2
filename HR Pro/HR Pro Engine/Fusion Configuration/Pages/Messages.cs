using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinTree;
using NHibernate;
using NHibernate.Linq;

namespace Fusion.Pages
{
	public partial class Messages : UserControl
	{
		private ISession _session;

		public Messages()
		{
			InitializeComponent();

			messageTree.ViewStyle = ViewStyle.Standard;
			messageTree.DisplayStyle = UltraTreeDisplayStyle.WindowsVista;
			messageTree.NodeConnectorStyle = NodeConnectorStyle.None;
			messageTree.Override.ItemHeight = 20;
			messageTree.Override.BorderStyleNode = UIElementBorderStyle.None;
			messageTree.NodeLevelOverrides[1].ShowExpansionIndicator = ShowExpansionIndicator.Never;

			//tree control does pickup on infragistics themes, fix it
			messageTree.Appearance.BorderColor = Color.FromArgb(255, 176, 202, 235);
			messageTree.Override.ActiveNodeAppearance.ForeColor = Color.White;
			messageTree.Override.ActiveNodeAppearance.BackColor = Color.FromArgb(255, 255, 175, 75);
			messageTree.Override.ActiveNodeAppearance.BackColor2 = Color.FromArgb(255, 254, 222, 119);
			messageTree.Override.ActiveNodeAppearance.BackGradientStyle = GradientStyle.GlassTop50;
		}

		public void Display(ISession session)
		{
			if(_session == null) {
				_session = session;

				using (new WaitCursor()) {
					messageBindingSource.DataSource = _session.Query<FusionMessage>().OrderBy(m => m.Name)
						.AsEnumerable()
						.Select(x => new FusionMessageNode(x))
						.ToList();

					//tree doesnt bother showing anything, for some reason it thinks its hidden, fix it
					messageTree.BringToFront();
					messageTree.SendToBack();
				}
			}
			messageTree.Select();
		}
	}


	//tree wont do multiple level data binding to FusionMessage.Elements as it needs Elements to be an IBindingList
	//so use MVVM and create a ViewModel that all controls on the screen can data bind to correctly
	public class FusionMessageNode
	{
		public readonly FusionMessage Message;
		private BindingList<FusionMessageElementNode> _elements;

		public FusionMessageNode(FusionMessage message)
		{
			Message = message;
		}

		public string Name
		{
			get { return Message.Name; }
		}

		public bool AllowPublish
		{
			get { return Message.AllowPublish; }
			set { Message.AllowPublish = value; }
		}

		public bool AllowSubscribe
		{
			get { return Message.AllowSubscribe; }
			set { Message.AllowSubscribe = value; }
		}
		public bool Publish
		{
			get { return Message.Publish; }
			set { Message.Publish = value; }
		}

		public bool Subscribe
		{
			get { return Message.Subscribe; }
			set { Message.Subscribe = value; }
		}

		public string XmlTemplate
		{
			get
			{
				var sb = new StringBuilder(200);

				sb.AppendFormat("<{0}>", Message.Name);

				foreach (var item in Message.Elements.OrderBy(m => m.Position).ToList())
				{
					if (item.Element.Column != null)
						sb.AppendFormat("\n\t<{0}>{1}.{2}</{0}>", item.NodeKey, item.Element.Column.Table.Name, item.Element.Column.Name);
					else
						sb.AppendFormat("\n\t<{0}>{{MAPPING UNDEFINED}}</{0}>", item.NodeKey);
				}
				sb.AppendFormat("\n</{0}>", Message.Name);

				return sb.ToString()
					.Replace("<", "&lt;")
					.Replace(">", "&gt;")
					.Replace("\n", "<br/>")
					.Replace("\t", "&nbsp;")
					.Replace("{", "<span style='color:Red;'>")
					.Replace("}","</span>");
			}
		}

		public BindingList<FusionMessageElementNode> Items
		{
			get
			{
				if (_elements == null) {
					var q = Message.Elements.Select(x => new FusionMessageElementNode(x)).ToList();
					_elements = new BindingList<FusionMessageElementNode>(q);
				}
				return _elements;
			}
		}
	}

	public class FusionMessageElementNode
	{
		private readonly FusionMessageElement _element;

		public FusionMessageElementNode(FusionMessageElement element)
		{
			_element = element;
		}

		public string Name
		{
			get { return _element.NodeKey; }
		}
		public int Position
		{
			get { return _element.Position; }
		}
	}
}