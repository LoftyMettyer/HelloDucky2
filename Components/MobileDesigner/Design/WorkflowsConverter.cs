using System;
using System.ComponentModel;

namespace MobileDesigner.Design
{
    public class ItemsConverter : TypeConverter
    {
        public override object ConvertTo(ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, Type destinationType)
        {
            return "Click to select";
        }
    }
}
