using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Data;

namespace MyControl.Utility
{
    /// <summary>
    /// A converter class that represents a double value as a string in percentages
    /// </summary>
    public class ZoomLevelConverter : IValueConverter
    {
        /// <summary>
        /// Converts a double value to a percentage string
        /// </summary>
        /// <param name="value">the value in double format</param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            int zoomPercent = (int)(Math.Round((double)value, 2) * 100);
            return (zoomPercent.ToString());            
        }

        /// <summary>
        /// Converts a percentage string back to a double value
        /// </summary>
        /// <param name="value">the value in percentage as string</param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (((string)value) == string.Empty)
            {
                return 1;
            }
            else
            {
                string zoomString = ((string)value);
                char[] charsToTrim = { ' ', '%' };
                string trimmedString = zoomString.TrimEnd(charsToTrim);
                double result;

                if (double.TryParse((string)trimmedString, out result) && result != 0)
                    return result * 0.01;
                else
                    return 1;
            }
        }
    }



}
