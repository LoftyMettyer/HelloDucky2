using MobileDesigner.Controls;

namespace MobileDesigner.Pages
{
    public partial class NewRegistrationPage : MobilePage
    {
        public NewRegistrationPage()
        {
            InitializeComponent();
            SetupControlTableLayout(controlsPanel);
            SetupButtonTableLayout(buttonsPanel);
        }
    }
}
