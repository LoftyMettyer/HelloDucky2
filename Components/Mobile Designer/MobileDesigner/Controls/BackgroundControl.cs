using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MobileDesigner.Controls
{
    public partial class BackgroundControl : UserControl
    {
        private Image _backgroundImage;
        private ImageLayout _backgroundImageLayout = ImageLayout.TopLeft;

        public BackgroundControl()
        {
            InitializeComponent();
        }

        public new Image BackgroundImage
        {
            get { return _backgroundImage; }
            set 
            { 
                _backgroundImage = value;
                OnBackgroundImageChanged(EventArgs.Empty);
            }
        }

        [DefaultValue(typeof(ImageLayout), "TopLeft")]
        public new ImageLayout BackgroundImageLayout
        {
            get { return _backgroundImageLayout; }
            set
            {
                _backgroundImageLayout = value;
                OnBackgroundImageLayoutChanged(EventArgs.Empty);
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);

            if(BackgroundImage != null)
            {
                DrawBackgroundImage(e.Graphics, BackgroundImage, BackgroundImageLayout, ClientRectangle);
            }
        }

        static void DrawBackgroundImage(Graphics g, Image image, ImageLayout layout, Rectangle bounds)
        {
            if(layout == ImageLayout.TopLeft || layout == ImageLayout.TopRight || layout == ImageLayout.Centre)
            {
                var location = Point.Empty;                

                if (layout == ImageLayout.TopRight)
                    location = new Point(bounds.Width - image.Width, 0);
                else if(layout == ImageLayout.Centre)
                    location = new Point((bounds.Width - image.Width) / 2, (bounds.Height - image.Height) / 2);

                g.DrawImage(image, location.X, location.Y, image.Size.Width, image.Size.Height);
                return;
            }

            if(layout == ImageLayout.Tile || layout == ImageLayout.TopTile || layout == ImageLayout.LeftTile || layout == ImageLayout.RightTile || layout == ImageLayout.BottomTile)
            {
                Rectangle rectangle = bounds;

                if(layout == ImageLayout.TopTile)
                    rectangle = new Rectangle(0, 0, bounds.Width, image.Height);
                else if(layout == ImageLayout.LeftTile)
                    rectangle = new Rectangle(0, 0, image.Width, bounds.Height);
                else if(layout == ImageLayout.RightTile)
                    rectangle = new Rectangle(bounds.Width - image.Width, 0, image.Width, bounds.Height);
                else if (layout == ImageLayout.BottomTile)
                    rectangle = new Rectangle(0, bounds.Height - image.Height, bounds.Width, image.Height);

                using (var brush = new TextureBrush(image))
                {
                    brush.TranslateTransform(rectangle.X, rectangle.Y);
                    g.FillRectangle(brush, rectangle);
                    return;
                }
            }
        }
    }
}
