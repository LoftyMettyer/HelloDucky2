using MobileDesigner.Controls;

namespace MobileDesigner.Pages
{
    public partial class TodoListPage : MobilePage
    {
        public TodoListPage()
        {
            InitializeComponent();
            SetupControlTableLayout(controlsPanel);
            SetupButtonTableLayout(buttonsPanel);
        }

        public override void  SetLayout(Layout layout)
        {
 	    base.SetLayout(layout);

        todoLinkButton.TitleFont = GetFont(Layout.TodoTitleFont);
        todoLinkButton.TitleForeColor = GetColor(Layout.TodoTitleForeColor);
        todoLinkButton.DescriptionFont = GetFont(Layout.TodoDescFont);
        todoLinkButton.DescriptionForeColor = GetColor(Layout.TodoDescForeColor);
        }

        public override void UpdateLayout()
        {
            base.UpdateLayout();

            Layout.TodoTitleFont = GetFont(todoLinkButton.TitleFont);
            Layout.TodoTitleForeColor = GetColor(todoLinkButton.TitleForeColor);
            Layout.TodoDescFont = GetFont(todoLinkButton.DescriptionFont);
            Layout.TodoDescForeColor = GetColor(todoLinkButton.DescriptionForeColor);
        }
    }
}
