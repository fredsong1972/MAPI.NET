using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace SpamBounce
{
    class BounceMailConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter,  System.Globalization.CultureInfo culture)
        {
            var param = new BounceMail();
            if (values.Length != 3)
                throw new ArgumentException("BounceMailConverter failed");
            param.EmailAddress = values[0] as string;
            param.AttachOriginalMessage = (bool)values[1];
            param.DeleteOriginalMessage = (bool)values[2];
            
            return param;

        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {

            throw new NotImplementedException();

        }
    }

    [ValueConversion(typeof(bool), typeof(bool))]
    public class InverseBooleanConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool))
                throw new InvalidOperationException("The target must be a boolean");

            return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        #endregion
    }
}
