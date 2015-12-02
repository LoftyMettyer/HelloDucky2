using System;
using System.ComponentModel;
using System.Drawing.Design;
using System.Windows.Forms;

namespace MobileDesigner.Design
{
    public class PictureEditor : UITypeEditor
    {
        private PictureDialog _dialog;

        public override object EditValue(ITypeDescriptorContext context, IServiceProvider provider, object value)
        {
            if(_dialog == null)
                _dialog = new PictureDialog();

            _dialog.Pictures = DesignerForm.Instance.Pictures;
            _dialog.Picture = (Picture)value;

            if (_dialog.ShowDialog() == DialogResult.OK)
            {
                value = _dialog.Picture;

                foreach(var item in _dialog.AddedPictures)
                    DesignerForm.Instance.Pictures.Add(item);
            }

            return value;
        }

        public override UITypeEditorEditStyle GetEditStyle(ITypeDescriptorContext context)
        {
            return UITypeEditorEditStyle.Modal;
        }
    }
}
