using MobileDesigner.Controls;

namespace MobileDesigner.Pages
{
    public partial class LoginPage : MobilePage
    {
        public LoginPage()
        {
            InitializeComponent();
            SetupControlTableLayout(controlsPanel);
            SetupButtonTableLayout(buttonsPanel);
        }
    }
}
