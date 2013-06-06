using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Drawing;
using System.Windows.Forms;

namespace MobileDesigner.Controls
{
    [Designer("System.Windows.Forms.Design.ParentControlDesigner, System.Design", typeof(IDesigner))]
    public partial class HeaderPanel : BackgroundControl
    {
        private Picture _backgroundImage;
        private Picture _logoImage;
        private ImageLocation _logoImageLocation = ImageLocation.TopLeft;
        private Point _logoImageOffset;

        public HeaderPanel()
        {
            InitializeComponent();
        }

        public new Picture BackgroundImage
        {
            get { return _backgroundImage; }
            set
            {
                _backgroundImage = value;
                base.BackgroundImage = _backgroundImage != null ? _backgroundImage.Image.ToImage() : null;
            }
        }

        public Picture LogoImage
        {
            get { return _logoImage; }
            set
            {
                _logoImage = value;
                pictureBox.Image = _logoImage == null ? null : _logoImage.Image.ToImage();
                pictureBox.Size = _logoImage == null ? Size.Empty : pictureBox.Image.Size;
                PerformLayout();
            }
        }

        public Size LogoImageSize
        {
            get { return pictureBox.Size; }
            set
            {
                pictureBox.Size = value;
                PerformLayout();
            }
        }

        [DefaultValue(typeof (ImageLocation), "TopLeft")]
        public ImageLocation LogoImageLocation
        {
            get { return _logoImageLocation; }
            set
            {
                _logoImageLocation = value;
                PerformLayout();
            }
        }

        public Point LogoImageOffset
        {
            get { return _logoImageOffset; }
            set
            {
                _logoImageOffset = value; 
                PerformLayout();
            }
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);

            var point = Point.Empty;

            switch(_logoImageLocation)
            {
                case ImageLocation.TopLeft :
                    point = new Point(_logoImageOffset.X, _logoImageOffset.Y);
                    break;
                case ImageLocation.TopRight :
                    point = new Point(ClientRectangle.Width - pictureBox.Width - _logoImageOffset.X, _logoImageOffset.Y);
                    break;
                case ImageLocation.BottomLeft :
                    point = new Point(_logoImageOffset.X, ClientRectangle.Height - pictureBox.Height - _logoImageOffset.Y);
                    break;
                case ImageLocation.BottomRight :
                    point = new Point(ClientRectangle.Width - pictureBox.Width - _logoImageOffset.X, ClientRectangle.Height - pictureBox.Height - _logoImageOffset.Y);
                    break;
            }
            pictureBox.Location = point;
        }
    }
}
