using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using MobileDesigner.Design;
using MobileDesigner.Pages;
using MobileDesigner.Properties;
using NHibernate;
using NHibernate.Linq;
using MobileDesigner.Controls;
using System.Runtime.InteropServices;
using NHibernate.Util;

namespace MobileDesigner
{
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class DesignerForm : Form
    {
        private bool _readOnly;
        private bool _isDirty;
        private bool _changesMade;
        private ISession _session;
        private Layout _pageLayout;
        private BindingList<Picture> _pictures; 
        private MobilePage _selectedPage;
        private Control _selectedControl;
        private readonly IList<Control> _hookedControls = new List<Control>();
        private readonly Dictionary<MobileForm, MobilePage> _loadedPages = new Dictionary<MobileForm, MobilePage>();

        public static DesignerForm Instance { get; set; }

        public DesignerForm()
        {
            InitializeComponent();

            //Force mnemonics to show all the time for consistency with vb6
            bool hotKey = true;
            Win32.SystemParametersInfo(Win32.SPI_SETKEYBOARDCUES, 0, ref hotKey, 0);
 
            showLoginPage.Click += (s, e) => ShowForm(MobileForm.Login, ((Button)s).Text);
            showHomePage.Click += (s, e) => ShowForm(MobileForm.HomePage, ((Button)s).Text);
            showTodoListPage.Click += (s, e) => ShowForm(MobileForm.TodoList, ((Button)s).Text);
            showNewRegistrationPage.Click += (s, e) => ShowForm(MobileForm.NewRegistration, ((Button)s).Text);
            showChangePasswordPage.Click += (s, e) => ShowForm(MobileForm.ChangePassword, ((Button)s).Text);
            showForgotPasswordPage.Click += (s, e) => ShowForm(MobileForm.ForgotPassword, ((Button)s).Text);
            propertyGrid.PropertyValueChanged += (s, e) => { 
                if(e.ChangedItem.PropertyDescriptor.Name == "SelectedUserGroup") return;
                IsDirty = true;
                Refresh();
            };
            saveToolButton.Click += (s, e) => Save();
            this.Shown += (s, e) => Cursor.Current = Cursors.Default;
        }

        private void DesignerFormLoad(object sender, EventArgs e)
        {
            _session = MobileDesignerSerivce.SessionFactory.OpenSession();
            _pageLayout = _session.Get<Layout>(1);
            Instance = this;
            IsDirty = false;
            showLoginPage.PerformClick();
        }

        public bool ReadOnly
        {
            get { return _readOnly; }
            set
            {
                _readOnly = value;
                propertyGrid.Enabled = !_readOnly;
                saveToolButton.Enabled = !_readOnly;
            }
        }

        public bool ShowForVB6()
        {
            Cursor.Current = Cursors.WaitCursor;
            ShowDialog();
            return _changesMade;
        }

        public void ShowForm(MobileForm form, string caption)
        {
            MobilePage page;
            _loadedPages.TryGetValue(form, out page);

            if(page == null) 
            {
                switch (form)
                {
                    case MobileForm.Login:
                        page = new LoginPage();
                        break;
                    case MobileForm.HomePage:
                        page = new HomePage();
                        var userGroups = _session.Query<UserGroup>().OrderBy(x => x.Name).ToList();
                        var workflows = _session.Query<Workflow>().Where(w => w.InitiationType == 0 && w.Deleted == false).ToList();

                        ((HomePage)page).ItemsList.UserGroups = userGroups;
                        ((HomePage)page).ItemsList.SelectedUserGroup = userGroups.FirstOrDefault();
                        ((HomePage)page).ItemsList.Workflows = workflows;
                        ((HomePage)page).ItemsList.ItemsChanged += (s, e) => IsDirty = true;
                        break;
                    case MobileForm.TodoList:
                        page = new TodoListPage();
                        break;
                    case MobileForm.NewRegistration:
                        page = new NewRegistrationPage();
                        break;
                    case MobileForm.ChangePassword:
                        page = new ChangePasswordPage();
                        break;
                    case MobileForm.ForgotPassword:
                        page = new ForgotPasswordPage();
                        break;
                }

                var elements = GetElements(form);
                page.SetElements(elements);
                page.TabStop = false;
                page.Location = new Point(-1000,0);
                designPanel.Controls.Add(page);
                page.Size = new Size(320, 450);
                page.Location = new Point(33, 31);
                page.PerformLayout();
                _loadedPages.Add(form, page);
            }

            if (page != _selectedPage)
            {
                if (_selectedPage != null)
                    _selectedPage.UpdateLayout();
                page.SetLayout(_pageLayout);
                SelectedControl = null;
                UnhookControls();
                HookControls(page);
                page.BringToFront();

                _selectedPage = page;
                SelectedControl = _selectedPage.GetInitalControl();
                Text = Resources.DesignerForm_Name + " - " + caption.Replace("&", string.Empty);
            }
        }

        private void HookControls(Control parentControl)
        {
            var controls = parentControl.FlattenChildren().ToList();

            foreach(Control c in controls) {
                c.MouseDown += HandleMouseDown;
                c.Resize += HandleResize;
                c.TabStop = false;
                _hookedControls.Add(c);
            }
        }

        private void UnhookControls()
        {
            _hookedControls.ForEach(c => c.MouseDown -= HandleMouseDown);
            _hookedControls.ForEach(c => c.Resize -= HandleResize);
            _hookedControls.Clear();
        }

        private void HandleMouseDown(object sender, MouseEventArgs e)
        {
            SelectedControl = GetSelectableControl((Control)sender);
        }

        private void HandleResize(object sender, EventArgs e)
        {
            if(SelectedControl == sender)
                SetupFocusControl(true);
        }

        private static bool CanClickThroughControl(Control control)
        {
            return !(control is TextBox || control is CheckBox);
        }

        private static Control GetSelectableControl(Control control)
        {
            if (control is TableLayoutPanel || (string)control.Tag == "NOSELECT")
                return GetSelectableControl(control.Parent);

            if (control.Parent is HeaderPanel || control.Parent is IconButton || control.Parent is ItemsList || control.Parent is LinkButton)
                control = control.Parent;
            if (control.Parent.Parent is HeaderPanel || control.Parent.Parent is IconButton || control.Parent.Parent is ItemsList || control.Parent.Parent is LinkButton)
                control = control.Parent.Parent;

            return control;
        }

        private Control SelectedControl
        {
            get { return _selectedControl; }
            set
            {
                if (_selectedControl != value)
                {
                    _selectedControl = value;
                    SetupFocusControl();
                    propertyGrid.SelectedObject = GetEditableProperties(_selectedControl);
                }
            }
        }

        private void SetupFocusControl(bool boundsOnly = false)
        {
            if (_selectedControl == null)
                focusControl.Visible = false;
            else
            {
                var parentScreenPoint = _selectedControl.Parent.PointToScreen(Point.Empty);
                var parentLocalPoint = designPanel.PointToClient(parentScreenPoint);
                var localPoint = new Point(parentLocalPoint.X + _selectedControl.Location.X, parentLocalPoint.Y + _selectedControl.Location.Y);

                focusControl.Bounds = new Rectangle(localPoint.X - 6, localPoint.Y - 6, _selectedControl.Width + 12, _selectedControl.Height + 12);
                if (boundsOnly) return;

                focusControl.AllowClickThrough = CanClickThroughControl(_selectedControl);
                focusControl.BringToFront();
                focusControl.Show();
                if (!CanClickThroughControl(_selectedControl))
                    focusControl.Focus();
                Refresh();
            }
        }

        private static object GetEditableProperties(Control control)
        {
            if (control is Label)
                return new LabelProperties((Label)control);
            if (control is TextBox)
                return new TextBoxProperties((TextBox)control);
            if(control is HeaderPanel)
                return new HeaderPanelProperties((HeaderPanel)control);
            if(control is MainPanel)
                return new MainPanelProperties((MainPanel)control);
            if(control is FooterPanel)
                return new FooterPanelProperties((FooterPanel)control);
            if (control is IconButton)
                return new IconButtonProperties((IconButton)control);
            if (control is LinkButton)
                return new LinkButtonProperties((LinkButton)control);
            if(control is ItemsList)
                return new WorkflowItemsListProperties((ItemsList)control);

            return null;
        }

        public IList<Picture> Pictures
        {
            get {
                if(_pictures == null) {
                    _pictures = new BindingList<Picture>(_session.Query<Picture>().ToList());
                    _pictures.ListChanged += (s, e) => IsDirty = true;
                }
                return _pictures;
            }
        }

        private IEnumerable<Element> GetElements(MobileForm form)
        {
            return _session.Query<Element>().Where(e => e.Form == form).ToList();
        }


        private bool IsDirty
        {
            get { return _isDirty; }
            set { 
                _isDirty = value;
                saveToolButton.Enabled = _isDirty;
            }
        }

        public void Save()
        {
            if (!IsDirty) return;

            Cursor.Current = Cursors.WaitCursor;

            if (_selectedPage != null)
                _selectedPage.UpdateLayout();

            foreach (var page in _loadedPages.Values)
                page.UpdateElements();

            using (var tx = _session.BeginTransaction())
            {
                var maxId = _session.CreateSQLQuery("select max(pictureid) from tmpPictures").UniqueResult<int>();
                
                Pictures.Where(p => p.Id == 0).ForEach(p =>
                {
                    p.Id = ++maxId;
                    p.New = true;
                    _session.Save(p);
                });

                _session.Flush();
                tx.Commit();
            }
            IsDirty = false;
            _changesMade = true;

            Cursor.Current = Cursors.Default;
        }

        private void DesignerFormFormClosing(object sender, FormClosingEventArgs e)
        {
            if(IsDirty) 
            {
                var result = MessageBox.Show(Resources.DesignerForm_SaveChanges, Resources.DesignerForm_Name, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);

                switch (result) {
                    case DialogResult.Yes:
                        Save();
                        break;
                    case DialogResult.No:
                        break;
                    case DialogResult.Cancel:
                        e.Cancel = true;
                        return;
                }
            }

            _session.Close();
        }

        private void DesignerFormKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                Help.ShowHelp(this, "OpenHR System Manager.chm"); //, HelpNavigator.TopicId, xxxx.ToString());
        }
    }
}
