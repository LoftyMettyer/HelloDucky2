using System;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;

namespace MobileDesigner.Design
{
    public class FontConverter : System.Drawing.FontConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            if (sourceType == typeof(string))
                return false;

            return base.CanConvertFrom(context, sourceType);
        }

        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string) && value is Font)
            {
                var font = (Font) value;
                var size = (int) Math.Round(font.Size);

                return font.Name + " " + size.ToString(CultureInfo.InvariantCulture);
            }
                

            return base.ConvertTo(context, culture, value, destinationType);
        }

        public override PropertyDescriptorCollection GetProperties(ITypeDescriptorContext context, object value, Attribute[] attributes)
        {
            var props = base.GetProperties(context, value, attributes);

            props.Remove(props["Name"]);
            props.Remove(props["Size"]);
            props.Remove(props["Unit"]);
            props.Remove(props["GdiCharSet"]);
            props.Remove(props["GdiVerticalFont"]);
            return props;
        }

        public override bool GetCreateInstanceSupported(ITypeDescriptorContext context)
        {
            return false;
        }
    }
}
