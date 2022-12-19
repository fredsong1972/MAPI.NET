using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using MAPI.NET;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace SpamBounce
{
    public partial class SpamBounceAddIn
    {
        const string defaultBounceMessage_ = "";
        OptionsPropertyPage page_;
        SpamBounceViewModel viewModel_;
        MAPISession session_;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           
            this.Application.OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);
            page_ = new OptionsPropertyPage();
            viewModel_ = page_.ViewModel;
            session_ = new MAPISession();
            session_.Initialize();
            session_.OpenMessageStore("");
            session_.MsgStore.RegisterEvents(EEventMask.fnevNewMail |
                                           EEventMask.fnevObjectCreated |
                                           EEventMask.fnevObjectCopied |
                                           EEventMask.fnevObjectDeleted |
                                           EEventMask.fnevObjectModified |
                                           EEventMask.fnevObjectMoved);
            viewModel_.MessageStore = session_.MsgStore;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            viewModel_.MessageStore = null;
            if (session_ != null)
            {
                session_.Dispose();
                session_ = null;
            }
            viewModel_.Save();
            this.Application.OptionsPagesAdd -= new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(Application_OptionsPagesAdd);
        }

        private void Application_OptionsPagesAdd(Outlook.PropertyPages pages)
        {

            pages.Add(page_, "Spam Bounce");
        }
             

        public static Outlook.Application OutlookApp
        {
            get;
            private set;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
