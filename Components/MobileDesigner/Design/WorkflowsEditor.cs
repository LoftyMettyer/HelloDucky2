using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.Windows.Forms;
using System.Linq;

namespace MobileDesigner.Design
{
    public class ItemsEditor : UITypeEditor
    {
        private ItemsDialog _dialog;

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            if (_dialog == null)
                _dialog = new ItemsDialog();

            var instance = (WorkflowItemsListProperties) context.Instance;

            _dialog.Items = instance.Workflows;
            _dialog.SelectedItems = instance.SelectedUserGroup.MobileWorkflows;

            if (_dialog.ShowDialog() == DialogResult.OK && _dialog.IsDirty)
                value = _dialog.SelectedItems;

            return value;
        }

        public override UITypeEditorEditStyle GetEditStyle(System.ComponentModel.ITypeDescriptorContext context)
        {
            return UITypeEditorEditStyle.Modal;
        }
    }
}
