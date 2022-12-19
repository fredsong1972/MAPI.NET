using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

namespace SpamBounce
{
    public class BounceMail : INotifyPropertyChanged
    {
        private string sendbehalfOfName_ = "Outlook";
        private string emailAddress_;
        private bool bDelete_;
        private bool bAttach_;
        private string message_;
        private string forward_;

        public event PropertyChangedEventHandler PropertyChanged;

        public BounceMail()
        {
            IsNew = true;
        }

        public string EmailAddress 
        {
            get { return emailAddress_; }
            set
            {
                if (value != emailAddress_)
                {
                    emailAddress_ = value;
                    NotifyPropertyChanged("EmailAddress");
                }
            }
        }

        public bool DeleteOriginalMessage 
        {
            get { return bDelete_; }
            set
            {
                if (value != bDelete_)
                {
                    bDelete_ = value;
                    NotifyPropertyChanged("DeleteOriginalMessage");
                }
            }
        }

        public bool AttachOriginalMessage 
        {
            get { return bAttach_; }
            set
            {
                if (value != bAttach_)
                {
                    bAttach_ = value;
                    NotifyPropertyChanged("AttachOriginalMessage");
                }
            }
        }

        public string SentOnBehalfName 
        { 
            get 
            { 
                return sendbehalfOfName_; 
            }

            set
            {
                if (value != sendbehalfOfName_)
                {
                    sendbehalfOfName_ = value;
                    NotifyPropertyChanged("SentOnBehalfName");
                }
            }
        }

        public string Message 
        {
            get
            {
                return message_;
            }
            set
            {
                if (value != message_)
                {
                    message_ = value;
                    NotifyPropertyChanged("Message");
                }
            }
        }

        public string Forward
        {
            get { return forward_; }
            set
            {
                forward_ = value;
                NotifyPropertyChanged("Forward");
            }
        }

        public List<string> Properties
        {
            get
            {
                List<string> properties = new List<string>();
                properties.Add(string.Format("Bounce back the message from {0}", EmailAddress));
                if (!string.IsNullOrEmpty(Forward))
                    properties.Add("And forward it to " + Forward);
                if (AttachOriginalMessage)
                    properties.Add("And attach it");
                if (DeleteOriginalMessage)
                    properties.Add("And delete it");
                return properties;
            }
        }

        public bool IsNew { get; set; }

        public void NotifyPropertyChanged(string PropertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(PropertyName));
                PropertyChanged(this, new PropertyChangedEventArgs("Properties"));
            }
        }

        public BounceMail Clone()
        {
            BounceMail clone = new BounceMail();
            clone.EmailAddress = this.EmailAddress;
            clone.DeleteOriginalMessage = this.DeleteOriginalMessage;
            clone.AttachOriginalMessage = this.AttachOriginalMessage;
            clone.Forward = this.Forward;
            clone.Message = this.Message;
            clone.IsNew = this.IsNew;
            return clone;
        }

        public void Copy(BounceMail copyFrom)
        {
            this.EmailAddress = copyFrom.EmailAddress;
            this.DeleteOriginalMessage = copyFrom.DeleteOriginalMessage;
            this.AttachOriginalMessage = copyFrom.AttachOriginalMessage;
            this.Forward = copyFrom.Forward;
            this.Message = copyFrom.Message;
            this.IsNew = copyFrom.IsNew;
        }
    }
}
