using MobileDesigner.Controls;

namespace MobileDesigner.Pages
{
    public partial class ForgotPasswordPage : MobilePage
    {
        public ForgotPasswordPage()
        {
            InitializeComponent();
            SetupControlTableLayout(controlsPanel);
            SetupButtonTableLayout(buttonsPanel);
        }
    }
}
