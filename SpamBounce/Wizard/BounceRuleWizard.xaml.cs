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
    /// Interaction logic for BounceRuleWizard.xaml
    /// </summary>
    public partial class BounceRuleWizard : NavigationWindow
    {
        private List<Page> pages_ = new List<Page>();

        public BounceRuleWizard(BounceMail bounce)
        {
            InitializeComponent();
            BounceMail = bounce;
            CreatePages();
        }

        public BounceMail BounceMail
        { get; private set; }


        private void CreatePages()
        {
            pages_.Add(new Page1(BounceMail));
            Navigated += new NavigatedEventHandler(NavigationService_Navigated);
            Navigate(pages_[0]); // navigate to the last page
        }

        private void NavigationService_Navigated(object sender, NavigationEventArgs e)
        {
        }
    }
}
