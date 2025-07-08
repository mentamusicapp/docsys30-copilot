using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Data;
using System.Windows.Data;

namespace DocumentsModule.View.UserControls
{
    public class DataTableFilters : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            DataView dt;
            if (value is DataView )
            {
                dt = value as DataView;
                var rows = dt.Cast<DataRowView>();
                return rows.Where(f => f["Sug"].ToString() == "סוג מסמך").ToList();
            }
            return new List<DataRowView>();
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class DataTableFilters2 : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            DataView dt;
            if (value is DataView)
            {
                dt = value as DataView;
                var rows = dt.Cast<DataRowView>();
                return rows.Where(f => f["Sug"].ToString() != "סוג מסמך").ToList();
            }
            return new List<DataRowView>();
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
