using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SpamBounce.Wizard
{
    /// <summary>
    /// Interaction logic for Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        public Page1(BounceMail mail)
        {
            InitializeComponent();
            ((System.Windows.Data.ObjectDataProvider)FindResource("bounceEmail")).ObjectInstance = mail;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Window thisWindow = (Window)this.Parent;
            thisWindow.DialogResult = false;
            thisWindow.Close();
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Window thisWindow = (Window)this.Parent;
            thisWindow.DialogResult = true;
            thisWindow.Close();
        }
    }
}
