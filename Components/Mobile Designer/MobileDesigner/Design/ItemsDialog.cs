using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

namespace MobileDesigner.Design
{
    public partial class ItemsDialog : Form
    {
        public IList<Workflow> Items { get; set; }
        public IList<Workflow> SelectedItems { get; set; }

        public ItemsDialog()
        {
            InitializeComponent();
            addButton.Click += (o, args) => Add();
            removeButton.Click += (o, args) => Remove();
            moveUpButton.Click += (o, args) => MoveUp();
            moveDownButton.Click += (o, args) => MoveDown();
            selectedItemsList.SelectedIndexChanged += (o, args) => SetEnabled();
        }

        private void ItemsDialogLoad(object sender, EventArgs e)
        {
            availableItemsList.DisplayMember = "Name";
            availableItemsList.Items.Clear();
            availableItemsList.Items.AddRange(Items.Except(SelectedItems).ToArray());
            selectedItemsList.DisplayMember = "Name";
            selectedItemsList.Items.Clear();
            selectedItemsList.Items.AddRange(SelectedItems.ToArray());
            availableItemsList.SelectedIndex = availableItemsList.Items.Count > 0 ? 0 : -1;
            selectedItemsList.SelectedIndex = selectedItemsList.Items.Count > 0 ? 0 : -1;
            SetEnabled();
            IsDirty = false;
        }

        private void Add()
        {
            if (!CanAdd()) return;
            var item = availableItemsList.SelectedItem;
            var index = availableItemsList.SelectedIndex;
            availableItemsList.Items.Remove(item);
            selectedItemsList.Items.Add(item);
            selectedItemsList.SelectedItem = item;
            availableItemsList.SelectedIndex = index > availableItemsList.Items.Count - 1 ? index - 1 : index;
            IsDirty = true;
            SetEnabled();
        }

        private void Remove()
        {
            if (!CanRemove()) return;
            var item = selectedItemsList.SelectedItem;
            var index = selectedItemsList.SelectedIndex;
            selectedItemsList.Items.Remove(item);
            availableItemsList.Items.Add(item);
            availableItemsList.SelectedItem = item;
            selectedItemsList.SelectedIndex = index > selectedItemsList.Items.Count - 1 ? index - 1 : index;
            IsDirty = true;
            SetEnabled();
        }

        private void MoveUp()
        {
            if (!CanMoveUp()) return;
            var item = selectedItemsList.SelectedItem;
            var index = selectedItemsList.SelectedIndex;
            selectedItemsList.Items.Remove(item);
            selectedItemsList.Items.Insert(index - 1, item);
            selectedItemsList.SelectedItem = item;
            IsDirty = true;
            SetEnabled();
        }

        private void MoveDown()
        {
            if (!CanMoveDown()) return;
            var item = selectedItemsList.SelectedItem;
            var index = selectedItemsList.SelectedIndex;
            selectedItemsList.Items.Remove(item);
            selectedItemsList.Items.Insert(index + 1, item);
            selectedItemsList.SelectedItem = item;
            IsDirty = true;
            SetEnabled();
        }

        private bool CanAdd()
        {
            return availableItemsList.Items.Count > 0;
        }

        private bool CanRemove()
        {
            return selectedItemsList.Items.Count > 0;
        }

        private bool CanMoveUp()
        {
            return selectedItemsList.Items.Count > 1 &&
                selectedItemsList.SelectedIndex > 0;
        }

        private bool CanMoveDown()
        {
            return selectedItemsList.Items.Count > 1 &&
                selectedItemsList.SelectedIndex < selectedItemsList.Items.Count - 1;
        }

        private void SetEnabled()
        {
            addButton.Enabled = CanAdd();
            removeButton.Enabled = CanRemove();
            moveUpButton.Enabled = CanMoveUp();
            moveDownButton.Enabled = CanMoveDown();
        }

        private void AvailableItemsListMouseDoubleClick(object sender, MouseEventArgs e)
        {
            var index = availableItemsList.IndexFromPoint(e.Location);
            if(index >= 0 && CanAdd())
                Add();
        }

        public bool IsDirty{ get; private set; }

        private void SelectedItemsListMouseDoubleClick(object sender, MouseEventArgs e)
        {
            var index = selectedItemsList.IndexFromPoint(e.Location);
            if (index >= 0 && CanRemove())
                Remove();
        }

        private void ItemsDialogFormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
                this.SelectedItems = selectedItemsList.Items.Cast<Workflow>().ToList();
            else
                this.SelectedItems = null;
        }
    }
}
