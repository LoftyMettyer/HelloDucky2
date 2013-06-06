using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using MobileDesigner.Properties;

namespace MobileDesigner.Design
{
    public partial class PictureDialog : Form
    {
        public PictureDialog()
        {
            InitializeComponent();
        }

        public IList<Picture> Pictures { get; set; } 
        public Picture Picture { get; set; }
        public IList<Picture> AddedPictures { get; private set; }

        private void PictureDialogLoad(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            AddedPictures = new List<Picture>();
            pictureList.Items.Clear();
            imageList.Images.Clear();
            
            var pictures = Pictures.OrderBy(p => p.Name).ToArray();
            var imageItems = pictures.Select(p => p.Image.ToImage()).ToArray();
            var listItems = pictures.Select((p,i) => new ListViewItem(p.Name) {ImageIndex = i, Tag = p}).ToArray();
            
            imageList.Images.AddRange(imageItems);
            pictureList.Items.AddRange(listItems);

            if (Picture != null)
                SelectPicture(Picture);
            else if (pictureList.Items.Count > 0)
                pictureList.Items[0].Selected = true;

            Cursor.Current = Cursors.Default;
        }

        private void AddPicture(Picture picture)
        {
            imageList.Images.Add(picture.Image.ToImage());

            var listItem = new ListViewItem(picture.Name)
            {
                ImageIndex = imageList.Images.Count - 1,
                Tag = picture
            };

            pictureList.Items.Add(listItem);
        }

        private void SelectPicture(Picture picture)
        {
            var item = pictureList.Items.Cast<ListViewItem>().FirstOrDefault(i => (Picture)i.Tag == picture);

            if(item != null) {
                item.Selected = true;
                item.EnsureVisible();
                pictureList.Select();
            }
        }

        private void PictureDialogFormClosing(object sender, FormClosingEventArgs e)
        {
            if (pictureList.SelectedItems.Count > 0)
                Picture = (Picture)pictureList.SelectedItems[0].Tag;

            if (DialogResult == DialogResult.OK && Picture == null)
                e.Cancel = true;
        }

        private void PictureListDoubleClick(object sender, EventArgs e)
        {
            buttonOK.PerformClick();
        }

        private void OpenPicture()
        {
            using (var fd = new OpenFileDialog())
            {
                fd.Title = Resources.Dialog_OpenPicture_Title;
                fd.Filter = Resources.Dialog_OpenPicture_Filter;

                if (fd.ShowDialog() == DialogResult.OK)
                {
                    var image = Image.FromFile(fd.FileName);

                    var picture = new Picture() { 
                        Name = Path.GetFileName(fd.FileName),
                        Image = image.ToByteArray(),
                        Type = image.GetPictureType()
                    };

                    if(Pictures.Any(p => p.Name == picture.Name))
                    {
                        MessageBox.Show(Resources.Dialog_OpenPicture_AlreadyExists, Resources.DesignerForm_Name);
                        return;
                    }
                    AddedPictures.Add(picture);
                    AddPicture(picture);
                    SelectPicture(picture);
                }
            }
        }

        private void ButtonNewClick(object sender, EventArgs e)
        {
            OpenPicture();
        }

        #region Fix to force an item to always be selected in listview

        private bool _busy;
        private void PictureListItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (!e.IsSelected && !_busy) {
                _busy = true;
                this.BeginInvoke(new ListViewItemSelectionChangedEventHandler(FixupSelection), new object[] { sender, e });
            }
        }
        private void FixupSelection(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            var lv = sender as ListView;
            if (lv.SelectedItems.Count == 0) e.Item.Selected = true;
            _busy = false;
        }
        #endregion
    }
}
