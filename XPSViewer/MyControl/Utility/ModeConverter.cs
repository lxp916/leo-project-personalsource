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
//using MyControl.Controls;

namespace MyControl.Utility
{

    /// <summary>
    /// Represents a class that that converts between DocumentViewer enum types  and integers for index binding
    /// </summary>
    public class ModeConverter : IValueConverter
    {
        
        /// <summary>
        /// Convert DocumentViewer enum to integer
        /// </summary>
        /// <param name="value">DocumentViewer enum type</param>
        /// <param name="targetType"></param>
        /// <param name="parameter">type of enum, tool or pageview</param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            //if(((string)parameter).Equals("tool"))
            //{
            //    switch ((DocumentViewer.ToolModes)value)
            //    {
            //        case DocumentViewer.ToolModes.PanAndAnnotationEdit:
            //            return 0;
            //        case DocumentViewer.ToolModes.TextSelect:
            //            return 1;
            //    }
            //    return -1;
            //}
            //else if (((string)parameter).Equals("pageview"))
            //{

            //    switch ((ReaderControl.PageViewModes)value)
            //    {
            //        case ReaderControl.PageViewModes.Zoom:
            //            return -1;
            //        case ReaderControl.PageViewModes.FitWidth:
            //            return 0;
            //        case ReaderControl.PageViewModes.FitPage:
            //            return 1;
            //    }
            
            //}
            //else if (((string)parameter).Equals("fit"))
            //{
                
            //    switch ((ReaderControl.PageViewModes)value)
            //    {
            //        case ReaderControl.PageViewModes.Zoom:
            //            return -1;
            //        case ReaderControl.PageViewModes.FitWidth:
            //            return 0;
            //        case ReaderControl.PageViewModes.FitPage:
            //            return 1;
            //    }
            //}
            
            return -1;
        }

        /// <summary>
        /// Converts integer to DocumentViewer enum
        /// </summary>
        /// <param name="value">integer index</param>
        /// <param name="targetType"></param>
        /// <param name="parameter">type of enum, tool or pageview</param>
        /// <param name="culture"></param>
        /// <returns></returns>
 
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            //if (((string)parameter).Equals("tool"))
            //{
            //    switch ((int)value)
            //    {
            //        case -1:
            //            return DocumentViewer.ToolModes.PanAndAnnotationEdit;
            //        case 0:
            //            return DocumentViewer.ToolModes.PanAndAnnotationEdit;
            //        case 1:
            //            return DocumentViewer.ToolModes.TextSelect;
            //    }

            //}
            //else if (((string)parameter).Equals("pageview"))
            //{

            //    switch ((int)value)
            //    {
            //        case -1:
            //            return ReaderControl.PageViewModes.Zoom;
            //        case 0:
            //            return ReaderControl.PageViewModes.FitWidth;
            //        case 1:
            //            return ReaderControl.PageViewModes.FitPage;
            //    }
            //}
            //else if (((string)parameter).Equals("fit"))
            //{
            //    switch ((ReaderControl.PageViewModes)value)
            //    {
            //        case ReaderControl.PageViewModes.Zoom:
            //            return -1;
            //        case ReaderControl.PageViewModes.FitWidth:
            //            return 0;
            //        case ReaderControl.PageViewModes.FitPage:
            //            return 1;
            //    }
            //}

            return null;            
        }
        

    }

}
