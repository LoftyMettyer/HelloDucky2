using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Collections.Generic;

namespace MobileDesigner
{
    public static class Extensions
    {
        public static Image ToImage(this byte[] instance)
        {
            var ms = new MemoryStream(instance);
            var image = Image.FromStream(ms);
            return image;

        }

        public static byte[] ToByteArray(this Image instance)
        {
            using (var ms = new MemoryStream())
            {
                instance.Save(ms, instance.RawFormat);
                return ms.ToArray();
            } 
        }

        public static PictureType GetPictureType(this Image instance)
        {
            if (instance.RawFormat.Equals(ImageFormat.Icon))
                return PictureType.Icon;
            else if (instance.RawFormat.Equals(ImageFormat.Png))
                return PictureType.Png;
            else
                return PictureType.Bitmap;
        }

        public static IEnumerable<Control> FlattenChildren(this Control control)
        {
            var children = control.Controls.Cast<Control>();
            return children.SelectMany(c => FlattenChildren(c)).Concat(children);
        }
    }
}
