using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.ComponentModel;
using System.Windows.Input;
using System.Xml;
using System.Xml.Serialization;
using MAPI.NET;
using System.Threading.Tasks;
using SpamBounce.Wizard;

namespace SpamBounce
{
    [Serializable]
    [XmlRoot("BounceMailSettings")]
    public class SpamBounceViewModel : INotifyPropertyChanged
    {
        static string serializePath_ = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"SpamBounce\BounceMails.xml");
        public event PropertyChangedEventHandler PropertyChanged;

        [XmlArray("BounceMails")]
        [XmlArrayItem("BounceMail", typeof(BounceMail))]
        private ObservableCollection<BounceMail> bounceEmails_ = new ObservableCollection<BounceMail>();

        [NonSerialized]
        private BounceMail selectedItem_;

        [NonSerialized]
        private MsgStore msgStore_;

        private Dictionary<string, BounceMail> email2Bounce_ = new Dictionary<string, BounceMail>();

        public SpamBounceViewModel()
        {
            bounceEmails_.CollectionChanged += new System.Collections.Specialized.NotifyCollectionChangedEventHandler(OnBounceEmailsChanged);
            AddCommand = new CreateBounceCommand(this);
            RemoveCommand = new RemoveBounceCommand(this);
            ChangeCommand = new ChangeBounceCommand(this);
            BounceCommand = new BounceCommand(this);
        }

        public void NotifyPropertyChanged(string PropertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(PropertyName));
        }


        public ObservableCollection<BounceMail> BounceEmails
        {
            get { return bounceEmails_; }
        }

        [XmlIgnore]
        public BounceMail SelectedItem
        {
            get 
            {
                if (selectedItem_ == null && BounceEmails.Count > 0)
                    selectedItem_ = BounceEmails[0];
                return selectedItem_; 
            }
            set
            {
                if (value != selectedItem_)
                {
                    selectedItem_ = value;
                    NotifyPropertyChanged("SelectedItem");
                }
            }
        }

        private void OnBounceEmailsChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            switch (e.Action)
            {
                case System.Collections.Specialized.NotifyCollectionChangedAction.Add:
                    {
                        foreach (BounceMail item in e.NewItems)
                        {
                            item.IsNew = false;
                            email2Bounce_[item.EmailAddress] = item;
                        }
                    }
                    break;

                case System.Collections.Specialized.NotifyCollectionChangedAction.Remove:
                    foreach (BounceMail item in e.OldItems)
                    {
                        if (email2Bounce_.ContainsKey(item.EmailAddress))
                            email2Bounce_.Remove(item.EmailAddress);
                    }
                    break;
            }
            
            NotifyPropertyChanged("BounceEmails");
        }

        [XmlIgnore]
        public CreateBounceCommand AddCommand
        {
            get; 
            private set;
        }

        [XmlIgnore]
        public RemoveBounceCommand RemoveCommand
        {
            get;
            private set;
        }

        [XmlIgnore]
        public ChangeBounceCommand ChangeCommand
        {
            get;
            private set;
        }

        [XmlIgnore]
        public BounceCommand BounceCommand
        {
            get;
            private set;
        }

        [XmlIgnore]
        public MsgStore MessageStore
        {
            get { return msgStore_; }
            set
            {
                if (msgStore_ != null)
                {
                    msgStore_.OnObjectCreated -= new EventHandler<MsgStoreObjectEventArgs>(OnMAPINewMessage);
                    msgStore_.Dispose();
                }
                msgStore_ = value;
                if (msgStore_ != null)
                {
                    msgStore_.OnObjectCreated += new EventHandler<MsgStoreObjectEventArgs>(OnMAPINewMessage);
                }
            }
        }

        public void Save()
        {
            XmlSerializer s = new XmlSerializer(typeof(SpamBounceViewModel));
            string dir = System.IO.Path.GetDirectoryName(serializePath_);
            if (!System.IO.Directory.Exists(dir))
                System.IO.Directory.CreateDirectory(dir);
            XmlTextWriter xw = new XmlTextWriter(serializePath_, Encoding.UTF8);
            xw.Formatting = Formatting.Indented;
            s.Serialize(xw, this);
            xw.Close();
        }

        public BounceMail GetBounceMail(string emailAddress)
        {
            foreach (BounceMail mail in BounceEmails)
            {
                if (mail.EmailAddress.Trim().ToLower() == emailAddress.Trim().ToLower())
                    return mail;
            }
            return null;
        }

        public bool HasBounceMail(string emailAddress)
        {
            return email2Bounce_.ContainsKey(emailAddress);
        }

        private void OnMAPINewMessage(object sender, MsgStoreObjectEventArgs e)
        {
            if (e.EntryID != null && e.ObjectType == ObjectType.MAPI_MESSAGE)
            {
                MAPIMessage message = MessageStore.GetMAPIObject(e.EntryID) as MAPIMessage;
                Recipient sndr = message.Sender;
                if (sndr == null)
                    return;
                BounceMail bounce = GetBounceMail(sndr.Email);
                MAPIFolder parentFolder = null;
                if (bounce != null && e.ParentID != null && e.ParentID.Equals(MessageStore.InboxEntryID))
                {
                    parentFolder = MessageStore.OpenFolder(e.ParentID);
                    BounceParameter parameter = new BounceParameter { Bounce = bounce, OriginalMessage = message, Folder = parentFolder };
                    BounceCommand.Execute(parameter);
                 }
                if (message != null)
                    message.Dispose();
                if (parentFolder != null)
                    parentFolder.Dispose();
            }
        }


        public static SpamBounceViewModel LoadViewModel()
        {
            XmlSerializer s = new XmlSerializer(typeof(SpamBounceViewModel));
            if (!System.IO.File.Exists(serializePath_))
                return new SpamBounceViewModel();
            try
            {
                XmlTextReader xr = new XmlTextReader(serializePath_);
                SpamBounceViewModel viewModel = (SpamBounceViewModel)s.Deserialize(xr);
                return viewModel;
            }
            catch (Exception)
            {
                return new SpamBounceViewModel();
            }
        }

    }

    public class CreateBounceCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;
        private SpamBounceViewModel viewModel_;

        public CreateBounceCommand(SpamBounceViewModel viewModel)
        {
            viewModel_ = viewModel;
        }

        public void Execute(object parameter)
        {
            System.Windows.Controls.UserControl control = parameter as System.Windows.Controls.UserControl;
            System.Windows.Point location =  control.PointToScreen(new System.Windows.Point(0,0));
            //location.Offset(control.ActualWidth / 2, control.ActualHeight / 2);

            BounceMail bounceMail = new BounceMail();
            BounceRuleWizard wizard = new BounceRuleWizard(bounceMail);
            wizard.Left = location.X;
            wizard.Top = location.Y;
            if (wizard.ShowDialog() == true && !string.IsNullOrEmpty(bounceMail.EmailAddress) && !viewModel_.HasBounceMail(bounceMail.EmailAddress))
            {
                viewModel_.BounceEmails.Add(bounceMail);
                viewModel_.SelectedItem = bounceMail;
            }
            control.Focus();
         }

        public bool CanExecute(object parameter)
        {
            return true;
        }
    }

    public class ChangeBounceCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;
        private SpamBounceViewModel viewModel_;

        public ChangeBounceCommand(SpamBounceViewModel viewModel)
        {
            viewModel_ = viewModel;
        }

        public void Execute(object parameter)
        {
            System.Windows.Controls.UserControl control = parameter as System.Windows.Controls.UserControl;
            System.Windows.Point location = control.PointToScreen(new System.Windows.Point(0, 0));
            //location.Offset(control.ActualWidth / 2, control.ActualHeight / 2);

            BounceMail bounceMail = viewModel_.SelectedItem;
            if (bounceMail != null)
            {
                BounceRuleWizard wizard = new BounceRuleWizard(bounceMail.Clone());
                wizard.Left = location.X;
                wizard.Top = location.Y;
                if (wizard.ShowDialog() == true)
                {
                    bounceMail.Copy(wizard.BounceMail);
                }
            }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }
    }

    public class RemoveBounceCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;
        private SpamBounceViewModel viewModel_;

        public RemoveBounceCommand(SpamBounceViewModel viewModel)
        {
            viewModel_ = viewModel;
            viewModel.PropertyChanged += new PropertyChangedEventHandler(viewModel_PropertyChanged);
        }

        public void Execute(object parameter)
        {
            int index = viewModel_.BounceEmails.IndexOf(viewModel_.SelectedItem);
            if (index > -1)
            {
                viewModel_.BounceEmails.Remove(viewModel_.SelectedItem);
                if (index < viewModel_.BounceEmails.Count)
                    viewModel_.SelectedItem = viewModel_.BounceEmails[index];
                else if (viewModel_.BounceEmails.Count > 0)
                    viewModel_.SelectedItem = viewModel_.BounceEmails[0];
                else
                    viewModel_.SelectedItem = null;
            }

        }

        public bool CanExecute(object parameter)
        {
            return viewModel_.SelectedItem != null;
        }

        // This catches changes to the view model and causes the GUI to call CanExecute() (above)
        void viewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (CanExecuteChanged != null)
                CanExecuteChanged(this, EventArgs.Empty);
        }
    }

    public class BounceCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;
        private SpamBounceViewModel viewModel_;

        public BounceCommand(SpamBounceViewModel viewModel)
        {
            viewModel_ = viewModel;
            viewModel.PropertyChanged += new PropertyChangedEventHandler(viewModel_PropertyChanged);
        }

        public void Execute(object parameter)
        {
            Action<Object> action = (Object bounceParam) =>
            {
                if (viewModel_.MessageStore.Session == null)
                    return;
                MAPISession session = viewModel_.MessageStore.Session;
                MAPIMessage message = viewModel_.MessageStore.CreateMessage();
                MAPIMessage originalMessage = (bounceParam as BounceParameter).OriginalMessage;
                BounceMail bounce = (bounceParam as BounceParameter).Bounce;
                MAPIFolder parentFolder = (bounceParam as BounceParameter).Folder;
                MAPIAddressBook addressBook = session.OpendAddressBook();
                message.Subject = "Undeliverable: " + (originalMessage == null ? "test" : originalMessage.Subject);
                message.Body = bounce.Message;
                message.AddRecipient(bounce.EmailAddress, addressBook);
                if (originalMessage != null && bounce.AttachOriginalMessage)
                    message.AddAttachment(originalMessage, "original message");
                EntryID defaultSender = session.PrimaryIdentity;
                bool useDefaultSender = true;
                if (!string.IsNullOrEmpty(bounce.SentOnBehalfName))
                {
                    useDefaultSender = false;
                    string defaultEmail = addressBook.GetEmailAddress(defaultSender);
                    message.SetSender(bounce.SentOnBehalfName, defaultEmail, addressBook);
                 }

                if (!viewModel_.MessageStore.SendMessage(message, false))
                {
                    if (!useDefaultSender)
                    {
                        message.SetSender(session.PrimaryIdentity);
                        message.Send();
                    }
                }
                message.Dispose();
                if (!string.IsNullOrEmpty(bounce.Forward))
                {
                    MAPIMessage forward = viewModel_.MessageStore.ForwardMessage(originalMessage);
                    if (forward.AddRecipient(bounce.Forward, addressBook))
                    {
                        forward.Send();
                        forward.Dispose();
                    }

                }
                if (addressBook != null)
                    addressBook.Dispose();
                if (originalMessage != null && parentFolder != null && bounce.DeleteOriginalMessage)
                    parentFolder.DeleteMessage(originalMessage);
             
            };

            BounceParameter parm = parameter as BounceParameter;
            if (parm == null)
                parm = new BounceParameter { Bounce = viewModel_.SelectedItem, OriginalMessage = null };
            Task t1 = new Task(action, parm);
            t1.Start();
            t1.Wait();
            System.Threading.Thread.Sleep(5000);
            if (viewModel_.MessageStore.Session != null)
                viewModel_.MessageStore.Session.SendAll();
        }

        public bool CanExecute(object parameter)
        {
            return viewModel_.SelectedItem != null;
        }

        // This catches changes to the view model and causes the GUI to call CanExecute() (above)
        void viewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (CanExecuteChanged != null)
                CanExecuteChanged(this, EventArgs.Empty);
        }
    }

    public class BounceParameter
    {
        public BounceMail Bounce { get; set; }
        public MAPIMessage OriginalMessage { get; set; }
        public MAPIFolder Folder { get; set; }
    }
}
