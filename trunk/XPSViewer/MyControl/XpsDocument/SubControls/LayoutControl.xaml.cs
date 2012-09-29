using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Media.Imaging;

namespace MyControl.XpsDocument.SubControls
{
    public partial class LayoutControl : UserControl
    {
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register("Value", typeof(LayoutMode), typeof(LayoutControl), new PropertyMetadata(LayoutMode.SinglePage));
        public LayoutMode Value
        {
            get
            {
                return (LayoutMode)this.GetValue(LayoutControl.ValueProperty);
            }
            set
            {
                this.SetValue(LayoutControl.ValueProperty, value);
                (this.LayoutIcon).Source = new BitmapImage(new Uri(images[this.Value], UriKind.Relative));
            }
        }
        
        private Dictionary<LayoutMode, string> images;
        public event EventHandler PageLayoutChanged;
        public delegate void EventHandler(object sender, RoutedEventArgs e);
        public LayoutControl()
        {
            InitializeComponent();
            images = new Dictionary<LayoutMode, string>();
            images.Add(LayoutMode.SinglePage, "/Test;component/Assets/Images/page_single.png");
            images.Add(LayoutMode.DoublePage, "/Test;component/Assets/Images/page_facing.png");
            //images.Add(LayoutMode.Thumbnail, "/Test;component/Assets/Images/pictures_thumbs.png");
        }

       
        private void btnLayoutButton_Click(object sender, RoutedEventArgs e)
        {
            GeneralTransform gt = ((Button)sender).TransformToVisual(Application.Current.RootVisual);
            Point offset = gt.Transform(new Point(0, 0));

            ContextMenu conMenu = new ContextMenu()
            {
                VerticalOffset = offset.Y + (sender as Button).ActualHeight,
                HorizontalOffset = offset.X,
            };

            foreach (var item in images)
            {
                MenuItem menuItem = new MenuItem()
                {
                    Icon = new Image() { Source = new BitmapImage(new Uri(item.Value, UriKind.Relative)) },
                    Header = item.Key.ToString(),
                    Tag = item.Key,
                };
                menuItem.Click += new RoutedEventHandler(MenuItem_Click);
                if (item.Key == this.Value) menuItem.Background = new SolidColorBrush(Color.FromArgb(0x55, 0x87, 0xCE, 0xFA));
                conMenu.Items.Add(menuItem);
            }
            conMenu.IsOpen = true;
            conMenu.Closed += new RoutedEventHandler(conMenu_Closed);
        }
        void conMenu_Closed(object sender, RoutedEventArgs e)
        {
            this.btnLayoutButton.Focus();
        }

        void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Value = (LayoutMode)(sender as MenuItem).Tag;
            PageLayoutChanged(sender, e);
            
        }
    }
}
