using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace MobileDesigner.Controls
{
    public partial class SelectControl : UserControl 
    {
        private const int HandleSize = 6;
        private const int FocusInset = 2;

        public SelectControl()
        {
            InitializeComponent();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00000020; //WS_EX_TRANSPARENT 
                return cp;
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // do nothing
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var graphics = e.Graphics;

            // draw a focus rectangle
            var rect = new Rectangle(FocusInset, FocusInset, Width - (FocusInset * 2), Height - (FocusInset * 2));
            ControlPaint.DrawFocusRectangle(e.Graphics, rect, Color.Black, Color.Black);

            // draw some grab handles
            rect = new Rectangle(0, 0, HandleSize, HandleSize);
            ControlPaint.DrawGrabHandle(graphics, rect, true, true);
            rect.X = Width - HandleSize;
            ControlPaint.DrawGrabHandle(graphics, rect, true, true);
            rect.Y = Height - HandleSize;
            ControlPaint.DrawGrabHandle(graphics, rect, true, true);
            rect.X = 0;
            ControlPaint.DrawGrabHandle(graphics, rect, true, true);

        }

        // control should have a big whole through the middle of it
        // so we can click through to the controls underneath
        private void SetRegion()
        {
            if(AllowClickThrough)
            {
                var path = new GraphicsPath();
                var rect = ClientRectangle;
                rect.Inflate(1, 1);
                path.AddRectangle(rect);
                rect.Inflate(-1, -1);
                rect.Inflate(-HandleSize, -HandleSize);
                path.AddRectangle(rect);
                Region = new Region(path);
            }
            else
                Region = null;
        }

        protected override void OnResize(System.EventArgs e)
        {
            SetRegion();
            base.OnResize(e);
        }

        private bool _allowClickThrough;
        public bool AllowClickThrough
        {
            get { return _allowClickThrough; }
            set { 
                _allowClickThrough = value;
                SetRegion();
            }
        }
    }
}
