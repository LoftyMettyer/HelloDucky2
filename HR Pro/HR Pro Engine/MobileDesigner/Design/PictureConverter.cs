using System;
using System.ComponentModel;
using System.Linq;

namespace MobileDesigner.Design
{
    public class PictureConverter : TypeConverter 
    {
        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            if (destinationType == typeof(string) && value is Picture)
                return ((Picture)value).Name;

            return base.ConvertTo(context, culture, value, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value)
        {
            var s = value as string;
            if (s != null)
            {
                if(string.IsNullOrWhiteSpace(s))
                    return null;

                return DesignerForm.Instance.Pictures.SingleOrDefault(p => string.Equals(p.Name, s, StringComparison.OrdinalIgnoreCase));
            }

            return base.ConvertFrom(context, culture, value);
        }

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            if (sourceType == typeof(string))
                return true;

            return base.CanConvertFrom(context, sourceType);
        }
    }
}
