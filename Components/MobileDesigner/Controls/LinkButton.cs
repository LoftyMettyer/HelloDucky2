using System.Drawing;
using System.Windows.Forms;

namespace MobileDesigner.Controls
{
    public partial class LinkButton : UserControl
    {
        public LinkButton()
        {
            InitializeComponent();
            BackColor = Color.Transparent;
            tableLayoutPanel1.Dock = DockStyle.Fill;
        }

        public Font TitleFont
        {
            get { return titleLabel.Font; }
            set { titleLabel.Font = value; }
        }
        public Color TitleForeColor
        {
            get { return titleLabel.ForeColor; }
            set { titleLabel.ForeColor = value; }
        }
        public Font DescriptionFont
        {
            get { return descriptionLabel.Font; }
            set { descriptionLabel.Font = value; }
        }
        public Color DescriptionForeColor
        {
            get { return descriptionLabel.ForeColor; }
            set { descriptionLabel.ForeColor = value; }
        }
    }
}
