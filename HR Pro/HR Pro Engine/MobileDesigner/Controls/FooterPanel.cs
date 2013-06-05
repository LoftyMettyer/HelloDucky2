using System.ComponentModel;
using System.ComponentModel.Design;

namespace MobileDesigner.Controls
{
    [Designer("System.Windows.Forms.Design.ParentControlDesigner, System.Design", typeof(IDesigner))]
    public partial class FooterPanel : BackgroundControl
    {
        private Picture _backgroundImage;

        public FooterPanel()
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
    }
}
