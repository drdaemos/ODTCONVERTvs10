using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace ODTCONVERT
{
    public class SideFlyout : StackPanel
    {
        public TextBlock Header
        {
            get
            {
                return Header;
            }
            set
            {
                Header = value;
            }
        }

        public static readonly DependencyProperty HeaderProperty =
          DependencyProperty.Register("Header", typeof(TextBlock), typeof(SideFlyout));
    }
}
