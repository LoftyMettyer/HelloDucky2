using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace MobileDesigner.Controls
{
    public partial class ItemsList : UserControl
    {
        public event EventHandler ItemsChanged;

        private UserGroup _selectedUserGroup;

        public ItemsList()
        {
            InitializeComponent();
            ItemFont = Font;
            ItemForeColor = ForeColor;
            tableLayout.Dock = DockStyle.Fill;
            tableLayout.Scroll += (s,e) => tableLayout.Refresh();
        }

        public IList<UserGroup> UserGroups { get; set; }
        public IList<Workflow> Workflows { get; set; }

        public UserGroup SelectedUserGroup
        {
            get { return _selectedUserGroup; }
            set 
            { 
                _selectedUserGroup = value;
                RenderItems(); 
            }
        }

        public IList<Workflow> SelectedWorkflows
        {
            get 
            {
                if (SelectedUserGroup != null)
                    return SelectedUserGroup.MobileWorkflows;

                return null;
            }
            set
            {
                if (SelectedUserGroup != null)
                {
                    SelectedUserGroup.MobileWorkflows.Clear();
                    foreach (var item in value)
                        SelectedUserGroup.MobileWorkflows.Add(item);
                    RenderItems();

                    if (ItemsChanged != null)
                        ItemsChanged(this, EventArgs.Empty);
                }
            }
        }

        private Font _itemFont;
        public Font ItemFont
        {
            get { return _itemFont; }
            set {
                _itemFont = value;
                OnItemFontChanged();
            }
        }

        private Color _itemForeColor;
        public Color ItemForeColor
        {
            get { return _itemForeColor; }
            set
            {
                _itemForeColor = value;
                OnItemForeColorChanged();
            }
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);

            ControlPaint.DrawBorder(e.Graphics, ClientRectangle, Color.Gray, ButtonBorderStyle.Dashed);
        }

        public void RenderItems()
        {
            tableLayout.SuspendLayout();
            tableLayout.Controls.Clear();
            tableLayout.RowStyles.Clear();

            if (SelectedUserGroup != null)
            {
                tableLayout.RowCount = SelectedUserGroup.MobileWorkflows.Count + 1;
                var row = 0;
                foreach (var item in SelectedUserGroup.MobileWorkflows)
                {
                    var picture = new PictureBox()
                    {
                        Size = new Size(32, 32),
                        SizeMode = PictureBoxSizeMode.StretchImage,
                        Image = item.Picture == null ? null : item.Picture.Image.ToImage(),
                        BorderStyle = item.Picture == null ? BorderStyle.FixedSingle : BorderStyle.None,
                        Anchor = AnchorStyles.Left,
                        Margin = new Padding(0, 3, 0, 3)
                    };
                    var label = new Label() { Text = item.Name, Font = this.ItemFont, ForeColor = this.ItemForeColor, AutoSize = true, Anchor = AnchorStyles.Left, UseMnemonic = false };

                    tableLayout.Controls.Add(picture, 0, row);
                    tableLayout.Controls.Add(label, 1, row);
                    tableLayout.RowStyles.Add(new RowStyle());
                    row++;
                }
                tableLayout.AutoScrollPosition = new Point(0, 0);
                tableLayout.Padding = new Padding(0, 0, 8, 0);
            }
            tableLayout.ResumeLayout(true);
        }

        protected void OnItemFontChanged()
        {
            foreach (var item in tableLayout.Controls.OfType<Label>())
                item.Font = ItemFont;
        }
        protected void OnItemForeColorChanged()
        {
            foreach (var item in tableLayout.Controls.OfType<Label>())
                item.ForeColor = ItemForeColor;
        }
    }
}