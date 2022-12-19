using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SpamBounce
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class OptionsPropertyPage : UserControl, Outlook.PropertyPage
    {
        private Outlook.PropertyPageSite site_;
        private bool isDirty_;
        private SpamBounceViewModel viewModel_;

        public OptionsPropertyPage()
        {
            InitializeComponent();
            viewModel_ = SpamBounceViewModel.LoadViewModel();
            ((System.Windows.Data.ObjectDataProvider)View.FindResource("bounceEmailsDataProvider")).ObjectInstance = viewModel_;
            viewModel_.PropertyChanged += new PropertyChangedEventHandler(ViewModel__PropertyChanged);
        }

        [Browsable(false)]
        public PropertyTab View
        {
            get { return propertyTab1; }
        }

        [Browsable(false)]
        public SpamBounceViewModel ViewModel
        {
            get { return viewModel_; }
        }

        [Browsable(false)]
        public bool IsDirty
        {
            get { return isDirty_; }
            set
            {
                if (isDirty_ != value)
                {
                    isDirty_ = value;
                }
            }
        }

        protected virtual void OnApply()
        {
            if (IsDirty)
            {
                viewModel_.Save();
            }
        }
        
        protected override void OnLoad(EventArgs e)
        {
            Type type = typeof(UserControl);
            Type oleType = type.Assembly.GetType("System.Windows.Forms.UnsafeNativeMethods+IOleObject");
            if (oleType == null) throw new InvalidOperationException("Could not get 'System.Windows.Forms.UnsafeNativeMethods+IOleObject'.");
            System.Reflection.MethodInfo method = oleType.GetMethod("GetClientSite");
            if (method == null) throw new InvalidOperationException("Could not get method 'IOleObject.GetClientSite'.");
            site_ = method.Invoke(this, null) as Outlook.PropertyPageSite;
            base.OnLoad(e);
        }

        private void ViewModel__PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName != "SelectedItem")
            {
                IsDirty = true;
                if (site_ != null) site_.OnStatusChange();
            }
        }

        #region 'Implement Outlook.PropertyPage'

        void Outlook.PropertyPage.Apply()
        {
            OnApply();
            isDirty_ = false;
        }

        void Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {
        }

        bool Outlook.PropertyPage.Dirty
        {
            get
            {
                return IsDirty;
            }
        }

        #endregion

    }
}
