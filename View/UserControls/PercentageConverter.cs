using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace DocumentsModule.View.UserControls
{
    class PercentageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double originalSize=0;
            if (value is double)
                originalSize = (double)value;
            else
                return value;
            double percentage = 0;
            if (double.TryParse(parameter.ToString(), out percentage))
                    return originalSize * percentage;
            return value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
