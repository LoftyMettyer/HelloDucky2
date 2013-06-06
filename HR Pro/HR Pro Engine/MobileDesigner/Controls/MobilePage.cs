using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace MobileDesigner.Controls
{
    public partial class MobilePage : UserControl
    {
        protected new Layout Layout;
        private IEnumerable<Element> _elements;
 
        public MobilePage()
        {
            InitializeComponent();
        }

        public virtual void SetLayout(Layout layout)
        {
            Layout = layout;

            Header.BackColor = GetColor(layout.HeaderBackColor);
            Header.BackgroundImage = layout.HeaderPicture;
            Header.BackgroundImageLayout = layout.HeaderPictureLocation;

            Header.LogoImage = layout.HeaderLogo;
            Header.LogoImageSize = new Size(layout.HeaderLogoWidth, layout.HeaderLogoHeight);
            Header.LogoImageLocation = GetImageLocation(layout.HeaderLogoHorizontalOffsetBehaviour, layout.HeaderLogoVerticalOffsetBehaviour);
            Header.LogoImageOffset = new Point(layout.HeaderLogoHorizontalOffset, layout.HeaderLogoVerticalOffset);

            Main.BackColor = GetColor(layout.MainBackColor);
            Main.BackgroundImage = layout.MainPicture;
            Main.BackgroundImageLayout = layout.MainPictureLocation;

            Footer.BackColor = GetColor(layout.FooterBackColor);
            Footer.BackgroundImage = layout.FooterPicture;
            Footer.BackgroundImageLayout = layout.FooterPictureLocation;
        }

        public virtual void UpdateLayout()
        {
            Layout.HeaderBackColor = GetColor(Header.BackColor);
            Layout.HeaderPicture = Header.BackgroundImage;
            Layout.HeaderPictureLocation = Header.BackgroundImageLayout;
            Layout.HeaderLogo = Header.LogoImage;

            Layout.HeaderLogoWidth = Header.LogoImageSize.Width;
            Layout.HeaderLogoHeight = Header.LogoImageSize.Height;
            Layout.HeaderLogoHorizontalOffsetBehaviour = GetOffsetBehaviour(Header.LogoImageLocation, false);
            Layout.HeaderLogoVerticalOffsetBehaviour = GetOffsetBehaviour(Header.LogoImageLocation, true);
            Layout.HeaderLogoHorizontalOffset = Header.LogoImageOffset.X;
            Layout.HeaderLogoVerticalOffset = Header.LogoImageOffset.Y;

            Layout.MainBackColor = GetColor(Main.BackColor);
            Layout.MainPicture = Main.BackgroundImage;
            Layout.MainPictureLocation = Main.BackgroundImageLayout;

            Layout.FooterBackColor = GetColor(Footer.BackColor);
            Layout.FooterPicture = Footer.BackgroundImage;
            Layout.FooterPictureLocation = Footer.BackgroundImageLayout;
        }

        public void SetElements(IEnumerable<Element> elements)
        {
            _elements = elements;

            foreach (var element in _elements)
            {
                var control = FindControl(element.Name);

                switch (element.Type)
                {
                    case ElementType.Button:
                        var iconButton = control as IconButton;
                        if(iconButton != null)
                        {
                            iconButton.Image = element.Picture;
                            iconButton.Caption = element.Caption;
                            iconButton.CaptionColor = GetColor(element.ForeColor);
                        }
                        break;
                    case ElementType.Label:
                        var label = (Label)control;
                        label.Text = string.IsNullOrWhiteSpace(element.Caption) ? Design.LabelProperties.EmptyLabelText : element.Caption;
                        label.Font = GetFont(element.Font);
                        label.ForeColor = GetColor(element.ForeColor);
                        break;
                    case ElementType.TextBox:
                        var textBox = (TextBox)control;
                        textBox.Font = GetFont(element.Font);
                        textBox.ForeColor = GetColor(element.ForeColor);
                        break;
                }
            }
        }

        public void UpdateElements()
        {
            foreach (var element in _elements)
            {
                var control = FindControl(element.Name);

                switch (element.Type)
                {
                    case ElementType.Button:
                        var iconButton = control as IconButton;
                        if(iconButton != null)
                        {
                            element.Picture = iconButton.Image;
                            element.Caption = iconButton.Caption;
                            element.ForeColor = GetColor(iconButton.CaptionColor);
                        }
                        break;
                    case ElementType.Label:
                        var label = (Label)control;
                        element.Caption = label.Text == Design.LabelProperties.EmptyLabelText ? null : label.Text;
                        element.Font = GetFont(label.Font);
                        element.ForeColor = GetColor(label.ForeColor);
                        break;
                    case ElementType.TextBox:
                        var textBox = (TextBox)control;
                        element.Font = GetFont(textBox.Font);
                        element.ForeColor = GetColor(textBox.ForeColor);
                        break;
                }
            }
        }


        public virtual Control GetInitalControl()
        {
            return Header;
        }

        protected Font GetFont(FontSetting font)
        {
            if (font == null)
                return null;

            var style = FontStyle.Regular;
            if (font.Bold) style |= FontStyle.Bold;
            if (font.Italic) style |= FontStyle.Italic;
            if (font.Underline) style |= FontStyle.Underline;
            if (font.Strikeout) style |= FontStyle.Strikeout;

            return new Font(font.Name, font.Size, style);
        }

        protected FontSetting GetFont(Font font)
        {
            if (font == null)
                return null;

            return new FontSetting { Name = font.Name, Size = font.Size, Bold = font.Bold, Italic = font.Italic, Underline = font.Underline, Strikeout = font.Strikeout };
        }

        protected Color GetColor(int? color)
        {
            if (!color.HasValue)
                return Color.Black;

            return ColorTranslator.FromOle(color.Value);
        }

        protected int GetColor(Color color)
        {
            return ColorTranslator.ToOle(color);
        }

        private Control FindControl(string name)
        {
            return this.FlattenChildren().FirstOrDefault(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        private static ImageLocation GetImageLocation(int horizontalOffsetBehaviour, int verticalOffsetBehaviour)
        {
            if (verticalOffsetBehaviour == 0 && horizontalOffsetBehaviour == 0)
                return ImageLocation.TopLeft;
            if (verticalOffsetBehaviour == 0 && horizontalOffsetBehaviour == 1)
                return ImageLocation.TopRight;
            if (verticalOffsetBehaviour == 1 && horizontalOffsetBehaviour == 0)
                return ImageLocation.BottomLeft;
            if (verticalOffsetBehaviour == 1 && horizontalOffsetBehaviour == 1)
                return ImageLocation.BottomRight;

            return ImageLocation.TopLeft;
        }

        private static byte GetOffsetBehaviour(ImageLocation location, bool vertical)
        {
            if ((location == ImageLocation.BottomLeft || location == ImageLocation.BottomRight) && vertical)
                return 1;

            if ((location == ImageLocation.TopRight || location == ImageLocation.BottomRight) && !vertical)
                return 1;

            return 0;
        }

        protected void SetupControlTableLayout(TableLayoutPanel control)
        {
            control.Dock = DockStyle.Fill;
            control.BackColor = Color.Transparent;
            control.Padding = new Padding(3, 0, 5, 0);

            control.ColumnStyles.Clear();
            control.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            control.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
 
            foreach (Control c in control.Controls)
            {
                if (control.GetColumn(c) == 0 && control.GetColumnSpan(c) == 2)
                    if(c is Label)
                        c.Margin = new Padding(0, 10, 0, 10);
                    else
                        c.Margin = new Padding(0, 0, 0, 3);
                else if (control.GetColumn(c) == 0 && control.GetColumnSpan(c) == 1)
                    c.Margin = new Padding(0, 0, 2, 10);
                else
                    c.Margin = new Padding(0, 0, 0, 8);

                if (control.GetColumn(c) == 1 && control.GetColumnSpan(c) == 1 && (c is TextBox || c is CheckBox))
                {
                    c.Dock = DockStyle.None;
                    c.Anchor = AnchorStyles.Left;
                    c.Width = 160;
                }
                    
            }
        }

        protected void SetupButtonTableLayout(TableLayoutPanel control)
        {
            control.Dock = DockStyle.Fill;
            control.BackColor = Color.Transparent;
            control.Padding = new Padding(0);
            control.Margin = new Padding(0);
        }
    }
}
