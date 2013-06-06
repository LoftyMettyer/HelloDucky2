using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MobileDesigner.Controls
{
    public partial class IconButton : UserControl
    {
        private Picture _icon;

        public IconButton()
        {
            InitializeComponent();
            captionLabel.ForeColor = Color.White;
        }

        [DefaultValue("Caption")]
        public string Caption
        {
            get { return captionLabel.Text; }
            set { captionLabel.Text = value; }
        }

        public Picture Image
        {
            get { return _icon; }
            set
            {
                _icon = value;
                pictureBox.Image = value != null ? value.Image.ToImage() : null;
                pictureBox.BackColor = value == null ? Color.White : Color.Transparent;
            }
        }

        [DefaultValue(typeof(Color), "Transparent")]
        public override Color BackColor
        {
            get { return base.BackColor; }
            set { base.BackColor = value; }
        }

        [DefaultValue(typeof(Color), "White")]
        public Color CaptionColor
        {
            get { return captionLabel.ForeColor; }
            set { captionLabel.ForeColor = value; }
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            pictureBox.Left = (int)Math.Round((Width - pictureBox.Width) / 2M);
            pictureBox.Top = 4;
            captionLabel.Left = (int)Math.Round((Width - captionLabel.Width) / 2M);
            base.OnLayout(e);
        }

       
    }
}
