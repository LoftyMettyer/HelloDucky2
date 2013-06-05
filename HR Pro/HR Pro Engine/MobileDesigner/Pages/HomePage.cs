using System.Windows.Forms;
using System.Drawing;
using MobileDesigner.Controls;

namespace MobileDesigner.Pages
{
    public partial class HomePage : MobilePage
    {
        public HomePage()
        {
            InitializeComponent();
            SetupControlTableLayout(controlsPanel);
            SetupButtonTableLayout(buttonsPanel);

            ItemsList.BackColor = Color.Transparent;
        }

        public override void SetLayout(Layout layout)
        {
            base.SetLayout(layout);

            ItemsList.ItemFont = GetFont(Layout.HomeItemFont);
            ItemsList.ItemForeColor = GetColor(Layout.HomeItemForeColor);
        }

        public override void UpdateLayout()
        {
            base.UpdateLayout();

            Layout.HomeItemFont = GetFont(ItemsList.ItemFont);
            Layout.HomeItemForeColor = GetColor(ItemsList.ItemForeColor);
        }

        public override Control GetInitalControl()
        {
            return ItemsList;
        }
    }
}
