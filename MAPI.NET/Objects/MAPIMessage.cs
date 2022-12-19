using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

using Ole = System.Runtime.InteropServices.ComTypes;
namespace MAPI.NET
{
    /// <summary>
    /// Manages messages, attachments, and recipients.
    /// </summary>
    [ComVisible(false)]
    [ComImport()]
    [Guid("00020307-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMessage : IMAPIProp
    {
        /// <exclude/>
        new HRESULT GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        /// <exclude/>
        new HRESULT SaveChanges(uint ulFlags);
        /// <exclude/>
        new HRESULT GetProps([In, MarshalAs(UnmanagedType.LPArray)] uint[] lpPropTagArray, uint ulFlags, out uint lpcValues, ref IntPtr lppPropArray);
        /// <exclude/>
        new HRESULT GetPropList(uint ulFlags, out IntPtr lppPropTagArray);
        /// <exclude/>
        new HRESULT OpenProperty(uint ulPropTag, ref Guid lpiid, uint ulInterfaceOptions, uint ulFlags, out IntPtr lppUnk);
        /// <exclude/>
        new HRESULT SetProps(uint cValues, IntPtr lpPropArray, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT DeleteProps(IntPtr lpPropTagArray, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT CopyTo(uint ciidExclude, ref Guid rgiidExclude, [In, MarshalAs(UnmanagedType.LPArray)] uint[] lpExcludeProps, IntPtr ulUIParam,
            IntPtr lpProgress, ref Guid lpInterface, IntPtr lpDestObj, uint ulFlags, IntPtr lppProblems);
        /// <exclude/>
        new HRESULT CopyProps(IntPtr lpIncludeProps, uint ulUIParam, IntPtr lpProgress, ref Guid lpInterface,
            IntPtr lpDestObj, uint ulFlags, out IntPtr lppProblems);
        /// <exclude/>
        new HRESULT GetNamesFromIDs(out IntPtr lppPropTags, ref Guid lpPropSetGuid, uint ulFlags,
            out uint lpcPropNames, out IntPtr lpppPropNames);
        /// <exclude/>
        new HRESULT GetIDsFromNames(uint cPropNames, ref IntPtr lppPropNames, uint ulFlags, out IntPtr lppPropTags);
        /// <summary>
        /// Returns the message's attachment table.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags that relate to the creation of the table.</param>
        /// <param name="lppTable">Pointer to a pointer to the attachment table.</param>
        /// <returns>S_OK, if the attachment table was successfully retrieved; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT GetAttachmentTable(uint ulFlags, out IntPtr lppTable);
        /// <summary>
        /// Opens an attachment.
        /// </summary>
        /// <param name="ulAttachmentNum">Index number of the attachment to open, as stored in the attachment's PR_ATTACH_NUM property. </param>
        /// <param name="lpInterface">Pointer to the interface identifier (IID) representing the interface to be used to access the attachment. Passing NULL results in the attachment's standard interface, or IAttach, being returned.</param>
        /// <param name="uFlags">Bitmask of flags that controls how the attachment is opened.</param>
        /// <param name="lpAttach"> Pointer to a pointer to the open attachment.</param>
        /// <returns>S_OK, if the attachment was successfully opened; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT OpenAttach(uint ulAttachmentNum, IntPtr lpInterface, uint uFlags, out IntPtr lpAttach);
        /// <summary>
        /// Creates a new attachment.
        /// </summary>
        /// <param name="lpInterface">Pointer to the interface identifier (IID) representing the interface to be used to access the message. Passing NULL results in the message's standard interface, or IMessage, being returned.</param>
        /// <param name="uFlags">Bitmask of flags that controls how the attachment is created.</param>
        /// <param name="ulAttachmentNum">An index number identifying the newly created attachment. </param>
        /// <param name="lpAttach">Pointer to a pointer to the open attachment object.</param>
        /// <returns>S_OK, if the attachment was successfully created; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT CreateAttach(IntPtr lpInterface, uint uFlags, out uint ulAttachmentNum, out IntPtr lpAttach);
        /// <summary>
        /// Deletes an attachment.
        /// </summary>
        /// <param name="ulAttachmentNum"> Index number of the attachment to delete. This is the value for the attachment's PR_ATTACH_NUM property.</param>
        /// <param name="ulUIParam">Handle to the parent window of any dialog boxes or windows this method displays. </param>
        /// <param name="lpProgress">Pointer to a progress object that displays a progress indicator. </param>
        /// <param name="ulFlags">Bitmask of flags that controls the display of a user interface.</param>
        /// <returns>S_OK, if the attachment was successfully deleted; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT DeleteAttach(uint ulAttachmentNum, uint ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Returns the message's recipient table.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags that controls the return of the table. </param>
        /// <param name="lppTable">Pointer to a pointer to the recipient table.</param>
        /// <returns>S_OK, if the recipient table was returned successfully; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT GetRecipientTable(uint ulFlags, out IntPtr lppTable);
        /// <summary>
        /// Adds, deletes, or modifies message recipients.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags that controls the recipient changes. If zero is passed for the ulFlags parameter, ModifyRecipients replaces all existing recipients with the recipient list pointed to by the lpMods parameter.</param>
        /// <param name="lpMods">Pointer to an ADRLIST structure that contains a list of recipients to be added, deleted, or modified in the message.</param>
        /// <returns>S_OK, if the recipient list was successfully modified; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT ModifyRecipients(uint ulFlags, IntPtr lpMods);
        /// <summary>
        /// Saves all of the message's properties and marks the message as ready to be sent.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags used to control how a message is submitted</param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT SubmitMessage(uint ulFlags);
        /// <summary>
        /// Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of the message and manages the sending of read reports.
        /// </summary>
        /// <param name="ulFlags">Bitmask of flags that controls the setting of a message's read flag that is, the message's MSGFLAG_READ flag in its PR_MESSAGE_FLAGS property and the processing of read reports. </param>
        /// <returns>S_OK, if the message's read flag has been successfully set or cleared; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT SetReadFlag(uint ulFlags);
    }

    /// <summary>
    /// IMessage .Net Wrapper object
    /// </summary>
    public class MAPIMessage : MAPIObject
    {
        static Guid IID_MailMessage = new Guid("00020D0B-0000-0000-C000-000000000046");
        static Guid IID_IMessage = new Guid("00020307-0000-0000-C000-000000000046");

        const int BufSize = 4096;

        const uint MSGFLAG_READ = 1;
        const uint CLEAR_READ_FLAG = 4;

        Recipient sender_;
        MAPITable recipientTable_;
        string subject_;
        string messageClass_;
        uint messageFlags_;
        MessageFormat messageFormat_ = MessageFormat.DONTKNOW;
        /// <summary>
        /// Initializes a new instance of the MAPIMessage class. 
        /// </summary>
        /// <param name="message">IMessage object</param>
        public MAPIMessage(IMessage message)
            : this(message, 0, string.Empty)
        {
        }
        /// <summary>
        /// Initializes a new instance of the MAPIMessage class.
        /// </summary>
        /// <param name="message">IMessage object</param>
        /// <param name="messageClass">MessageClass propery value of the IMessage object</param>
        public MAPIMessage(IMessage message, string messageClass)
            : this(message, 0, messageClass)
        {
        }
        /// <summary>
        /// Initializes a new instance of the MAPIMessage class.
        /// </summary>
        /// <param name="message">IMessage object</param>
        /// <param name="messageFlags">MessageFlag property value of the IMessage object</param>
        public MAPIMessage(IMessage message, uint messageFlags)
            : this(message, messageFlags, string.Empty)
        {
        }
        /// <summary>
        /// Initializes a new instance of the MAPIMessage class.
        /// </summary>
        /// <param name="message">IMessage object</param>
        /// <param name="messageFlags">MessageFlg property value of the IMessage object.</param>
        /// <param name="messageClass">MessageClass property value of the IMessage object.</param>
        public MAPIMessage(IMessage message, uint messageFlags, string messageClass)
            : base(message, null)
        {
            List<PropTags> tags = new List<PropTags>();
            tags.Add(PropTags.PR_ENTRYID);
            tags.Add(PropTags.PR_MSG_EDITOR_FORMAT);
            if (messageFlags == 0)
                tags.Add(PropTags.PR_MESSAGE_FLAGS);
            if (string.IsNullOrEmpty(messageClass))
                tags.Add(PropTags.PR_MESSAGE_CLASS);
            Dictionary<PropTags, IPropValue> values = GetProperties(tags.ToArray());
            if (values.ContainsKey(PropTags.PR_ENTRYID) && values[PropTags.PR_ENTRYID].AsBinary != null)
                Id_ = new EntryID(values[PropTags.PR_ENTRYID].AsBinary);
            if (values.ContainsKey(PropTags.PR_MESSAGE_CLASS))
                messageClass_ = values[PropTags.PR_MESSAGE_CLASS].AsString;
            if (values.ContainsKey(PropTags.PR_MESSAGE_FLAGS))
                messageFlags_ = (uint)values[PropTags.PR_MESSAGE_FLAGS].AsInt32;
            if (values.ContainsKey(PropTags.PR_MSG_EDITOR_FORMAT))
                messageFormat_ = (MessageFormat)values[PropTags.PR_MSG_EDITOR_FORMAT].AsInt32;
        }

        #region Public Properties
        /// <summary>
        /// IMessage inteface GUID
        /// </summary>
        public override Guid InterfaceID
        {
            get
            {
                return IID_IMessage;
            }
        }

        /// <summary>
        /// Returns or sets a Boolean value that is True if the item has not been opened (read). 
        /// </summary>
        public bool Unread
        {
            get
            {
                return (MessageFlags & MSGFLAG_READ) == 0;
            }
            set
            {
                Message.SetReadFlag(value ? CLEAR_READ_FLAG : MSGFLAG_READ);
            }
        }

        /// <summary>
        /// Gets or sets Subject property value 
        /// </summary>
        public string Subject
        {
            get
            {
                if (string.IsNullOrEmpty(subject_))
                    subject_ = GetProperty(PropTags.PR_SUBJECT).AsString;
                return subject_;
            }
            set
            {
                SetPropertyValue(PropTags.PR_SUBJECT, value);
                subject_ = value;
            }
        }
        /// <summary>
        /// Gets MessageClass property value
        /// </summary>
        public virtual string MessageClass
        {
            get
            {
                return messageClass_;
            }
        }
        /// <summary>
        /// Gets MessageFlags property value
        /// </summary>
        public uint MessageFlags
        {
            get
            {
                return messageFlags_;
            }
        }
        /// <summary>
        /// Gets or sets Body of the message
        /// </summary>
        public string Body
        {
            get
            {
                string body = "";
                if (MessageFormat == MessageFormat.HTML)
                {
                    IPropValue val = GetProperty(PropTags.PR_BODY_HTML);
                    if (val != null)
                        body = val.AsString;
                }
                else if (MessageFormat == MessageFormat.PLAINTEXT)
                {
                    IPropValue val = GetProperty(PropTags.PR_BODY);
                    if (val != null)
                        body = val.AsString;
                }
                else
                {
                    body = GetRTFBody();
                    if (RtfHtmlConverter.IsHTMLinRtf(body))
                    {
                        string html = RtfHtmlConverter.HtmlFromRtf(body);
                        if (!string.IsNullOrEmpty(html))
                        {
                            MessageFormat = MessageFormat.RTFHTML;
                            return html;
                        }
                    }
                    MessageFormat = MessageFormat.RTF;
                }
                return body;
            }
            set
            {
                switch (MessageFormat)
                {
                    case MessageFormat.HTML:
                        SetPropertyValue(PropTags.PR_BODY_HTML, value);
                        break;
                    case MessageFormat.PLAINTEXT:
                        SetPropertyValue(PropTags.PR_BODY, value);
                        break;
                    case MessageFormat.RTF:
                        SetRTFBody(value);
                        break;
                    case MessageFormat.RTFHTML:
                        SetRTFHTMLBody(value);
                        break;
                }
            }
        }


        /// <summary>
        /// Gets Sender of the message
        /// </summary>
        public Recipient Sender
        {
            get
            {
                if (sender_ == null)
                {
                    EntryID entryId = null;
                    Dictionary<PropTags, IPropValue> values = GetProperties(new PropTags[] { PropTags.PR_SENDER_NAME, PropTags.PR_SENDER_ADDRTYPE, PropTags.PR_SENDER_EMAIL_ADDRESS, PropTags.PR_SENDER_ENTRYID });
                    if (values.ContainsKey(PropTags.PR_SENDER_ENTRYID))
                    {
                        entryId = new EntryID(values[PropTags.PR_SENDER_ENTRYID].AsBinary);
                    }
                    if (entryId != null)
                    {
                        string senderName = values.ContainsKey(PropTags.PR_SENDER_NAME) ? values[PropTags.PR_SENDER_NAME].AsString : string.Empty;
                        string senderAddress = values.ContainsKey(PropTags.PR_SENDER_EMAIL_ADDRESS) ? values[PropTags.PR_SENDER_EMAIL_ADDRESS].AsString : string.Empty;
                        AddressType addrType = AddressType.SMTP;
                        if (values.ContainsKey(PropTags.PR_SENDER_ADDRTYPE))
                            addrType = values[PropTags.PR_SENDER_ADDRTYPE].AsString == AddressType.EX.ToString() ? AddressType.EX : AddressType.SMTP;
                        sender_ = new Recipient(entryId, senderName, senderAddress, addrType);
                    }
                }
                return sender_;
            }
        }
        /// <summary>
        /// Gets a list of recipients of the message
        /// </summary>
        public List<Recipient> Recipients
        {
            get
            {
                List<Recipient> recipients = new List<Recipient>();
                if (recipientTable_ == null)
                    recipientTable_ = GetRecipientTable();
                if (recipientTable_ == null)
                    return recipients;
                recipientTable_.SeekRow(BookMark.BEGINNING, 0);
                if (recipientTable_.SetColumns(new PropTags[] { PropTags.PR_RECIPIENT_TYPE, PropTags.PR_DISPLAY_NAME, PropTags.PR_EMAIL_ADDRESS, PropTags.PR_ADDRTYPE, PropTags.PR_ENTRYID }))
                {
                    SRow[] sRows;

                    while (recipientTable_.QueryRows(1, out sRows))
                    {
                        if (sRows.Length != 1)
                            break;
                        int nType = sRows[0].propVals[0].AsInt32;
                        string name = sRows[0].propVals[1].AsString;
                        string email = sRows[0].propVals[2].AsString;
                        string addrType = sRows[0].propVals[3].AsString;
                        EntryID entryId = new EntryID(sRows[0].propVals[4].AsBinary);
                        recipients.Add(new Recipient(entryId, name, email, addrType == AddressType.EX.ToString() ? AddressType.EX : AddressType.SMTP));
                    }
                }
                return recipients;
            }
        }
        /// <summary>
        /// Gets the count of attachmentsof the message
        /// </summary>
        public int AttachmentCount
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_HASATTACH);
                if (!value.AsBool)
                    return 0;
                int count = 0;
                using (MAPITable tb = GetAttachmentTable())
                {
                    if (tb != null)
                    {
                        count = (int)tb.GetRowCount();
                    }
                }
                return count;
            }
        }

        /// <summary>
        /// Gets/sets a Date indicating the date and time at which the item was created.
        /// </summary>
        public DateTime? CreationTime
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_CREATION_TIME);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
            set
            {
                SetPropertyValue(PropTags.PR_CREATION_TIME, value);
            }
        }

        /// <summary>
        /// Returns a Date indicating the date and time at which the item was received. Read-only
        /// </summary>
        public DateTime? ReceivedTime
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_MESSAGE_DELIVERY_TIME);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
        }

        /// <summary>
        /// Gets/Sets SubmitTime
        /// </summary>
        public DateTime? SubmitTime
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_CLIENT_SUBMIT_TIME);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
            set
            {
                SetPropertyValue(PropTags.PR_CLIENT_SUBMIT_TIME, value);
            }
        }


        /// <summary>
        /// Gets/Sets the sensitivity of the contact
        /// </summary>
        public Sensitivity Sensitivity
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_SENSITIVITY);
                if (value != null)
                    return (Sensitivity)value.AsInt32;
                return Sensitivity.SENSITIVITY_NONE;
            }
            set
            {
                SetPropertyValue(PropTags.PR_SENSITIVITY, value);
            }
        }

        /// <summary>
        /// Gets/sets Importance
        /// </summary>
        public Importance Importance
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_IMPORTANCE);
                if (value != null)
                    return (Importance)value.AsInt32;
                return Importance.IMPORTANCE_NORMAL;
            }
            set
            {
                SetPropertyValue(PropTags.PR_IMPORTANCE, value);
            }
        }

        /// <summary>
        /// Gets/sets Priority
        /// </summary>
        public int Priority
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_PRIORITY);
                return value.AsInt32;
            }
            set
            {
                SetPropertyValue(PropTags.PR_PRIORITY, value);
            }
        }

        /// <summary>
        /// Gets the last modified time
        /// </summary>
        public DateTime? LastModifiedTime
        {
            get
            {
                IPropValue value = GetProperty(PropTags.PR_LAST_MODIFICATION_TIME);
                if (value != null)
                    return value.AsDateTime;
                return null;
            }
            set
            {
                SetPropertyValue(PropTags.PR_LAST_MODIFICATION_TIME, value);
            }
        }

        public MessageFormat MessageFormat
        {
            get
            {
              return messageFormat_;
            }
            set
            {
                if (messageFormat_ != value)
                {
                    messageFormat_ = value;
                    if (value == MessageFormat.RTFHTML && messageFormat_ != MessageFormat.RTF)
                    {
                        SetPropertyValue(PropTags.PR_MSG_EDITOR_FORMAT, (int)MessageFormat.RTF);
                    }
                    else
                    {
                        SetPropertyValue(PropTags.PR_MSG_EDITOR_FORMAT, (int)value);
                    }
                }
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Refresh the cached property values.
        /// </summary>
        public void Refresh()
        {
            List<PropTags> tags = new List<PropTags>();
            if (EntryID == null)
                tags.Add(PropTags.PR_ENTRYID);
            tags.Add(PropTags.PR_MESSAGE_FLAGS);
            tags.Add(PropTags.PR_MESSAGE_CLASS);
            Dictionary<PropTags, IPropValue> values = GetProperties(tags.ToArray());
            if (values.ContainsKey(PropTags.PR_ENTRYID) && values[PropTags.PR_ENTRYID].AsBinary != null)
                Id_ = new EntryID(values[PropTags.PR_ENTRYID].AsBinary);
            if (values.ContainsKey(PropTags.PR_MESSAGE_CLASS))
                messageClass_ = values[PropTags.PR_MESSAGE_CLASS].AsString;
            if (values.ContainsKey(PropTags.PR_MESSAGE_FLAGS))
                messageFlags_ = (uint)values[PropTags.PR_MESSAGE_FLAGS].AsInt32;
        }

        public bool AddRecipient(string strEmail)
        {
            return AddRecipient(strEmail, RecipientType.TO, AddressType.SMTP, null);
        }
        /// <summary>
        /// Add a recipient
        /// </summary>
        /// <param name="strEmail">email address of the recipient</param>
        /// <param name="addressBook">address book object which resolves the recipient </param>
        /// <returns>true, if the recipient was successfully added; otherwise, failed.</returns>
        public bool AddRecipient(string strEmail, MAPIAddressBook addressBook)
        {
            return AddRecipient(strEmail, RecipientType.TO, AddressType.SMTP, addressBook);
        }
        /// <summary>
        /// Add a recipient
        /// </summary>
        /// <param name="strEmail">email address of the recipient</param>
        /// <param name="recipientType">type of the recipient</param>
        /// <param name="addressType">address type of the recipient</param>
        /// <param name="addressBook">address book object which resolves the recipient</param>
        /// <returns>true, if the recipient was successfully added; otherwise, failed.</returns>
        public bool AddRecipient(string strEmail, RecipientType recipientType, AddressType addressType, MAPIAddressBook addressBook)
        {
            HRESULT hResult = HRESULT.E_FAIL;

            pSPropValue val1 = new pSPropValue();
            val1.ulPropTag = (int)PropTags.PR_RECIPIENT_TYPE;
            val1.Value.ul = (uint)recipientType;

            pSPropValue val2 = new pSPropValue();
            val2.ulPropTag = (int)PropTags.PR_DISPLAY_NAME;
            val2.Value.lpszW = Marshal.StringToHGlobalUni(strEmail);

            pSPropValue val3 = new pSPropValue();
            val3.ulPropTag = (int)PropTags.PR_ADDRTYPE;
            val3.Value.lpszW = Marshal.StringToHGlobalUni(addressType.ToString());

            pSPropValue[] values = new pSPropValue[] { val1, val2, val3 };
            IntPtr propValuePtr = values.MarshalToIntPtr();
            bool bResolved = true;
            ADRENTRY aEntry = new ADRENTRY();
            aEntry.cValues = 3;
            aEntry.rgPropVals = propValuePtr;

            IntPtr adrListPtr = new ADRENTRY[] { aEntry }.MarshalToAdrListPtr();
            if (addressBook != null)
            {
                bResolved = addressBook.ResolveName(adrListPtr);
            }
            if (bResolved)
            {
                hResult = Message.ModifyRecipients((uint)ModifyRecipientFlag.ADD, adrListPtr);
            }

            values.MarshalFreeIntPtr(propValuePtr);
            Marshal.FreeHGlobal(adrListPtr);

            return hResult == HRESULT.S_OK;
        }
        /// <summary>
        /// Set sender of the message
        /// </summary>
        /// <param name="name">sender display name</param>
        /// <param name="emailAddress">sender email address</param>
        /// <param name="addressBook">address book object where the specified sender will be created one off.</param>
        /// <returns>true, if the sender was successfully set; otherwise, failed</returns>
        public bool SetSender(string name, string emailAddress, MAPIAddressBook addressBook)
        {
            EntryID entry = addressBook.CreateOneOff(name, RecipientType.TO, emailAddress);
            return SetSender(entry);
        }

        public bool SetSender(EntryID addressEntry)
        {
            return SetPropertyValue((uint)PropTags.PR_SENT_REPRESENTING_ENTRYID, addressEntry.AsByteArray);
        }
        /// <summary>
        /// Set if save a message to the sent folder after sent.
        /// </summary>
        /// <param name="messageStore">Message store object</param>
        /// <param name="bSaveToSentFolder">Boolean flag which controls if save the message to the sent folder after sent.</param>
        /// <returns>true, if successfully; otherwise, failed</returns>
        public bool SaveToSentFolder(MsgStore messageStore, bool bSaveToSentFolder)
        {
            if (bSaveToSentFolder)
            {
                IPropValue value = messageStore.GetProperty(PropTags.PR_IPM_SENTMAIL_ENTRYID);
                if (value.AsBinary != null)
                {
                    return SetPropertyValue(PropTags.PR_SENTMAIL_ENTRYID, value.AsBinary);
                }
            }
            else
                return SetPropertyValue(PropTags.PR_DELETE_AFTER_SUBMIT, 1);
            return false;
        }
        /// <summary>
        /// Send the message
        /// </summary>
        /// <returns>true, if the message was successfully sent; otherwise, failed</returns>
        public bool Send()
        {
            Save();
            try
            {
                HRESULT hResult = Message.SubmitMessage(0);
                if (hResult == HRESULT.S_OK)
                {
                    Close();
                    return true;
                }
            }
            catch
            {
            }
            return false;
        }
        /// <summary>
        /// Save the message to MSG file.
        /// </summary>
        /// <returns>the file path of the MSG file.</returns>
        public string SaveToMsg()
        {
            string file = Subject.Trim();
            if (string.IsNullOrEmpty(file))
            {
                file = "Untitled.msg";
            }
            else
            {
                // Get a list of invalid file characters.
                char[] invalidFileChars = System.IO.Path.GetInvalidFileNameChars();
                foreach (char ch in invalidFileChars)
                {
                    string s = new string(ch, 1);
                    file = file.Replace(s, "");
                }
                file += ".msg";
            }
            file = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), file);
            if (SaveToMsg(file))
                return file;
            return string.Empty;
        }
        /// <summary>
        /// Save the message to a MSG file.
        /// </summary>
        /// <param name="outputFile">the specified output path</param>
        /// <returns>true, if the message was saved to MSG file successfully; otherwise, failed</returns>
        public bool SaveToMsg(string outputFile)
        {
            IMalloc malloc = null;
            IStorage stg = null;
            HRESULT hResult = HRESULT.E_FAIL;
            IntPtr pMsgSess = IntPtr.Zero;
            IMessage msg = null;
            malloc = MAPINative.MAPIGetDefaultMalloc();
            if (malloc != null)
            {
                hResult = Storage.StgCreateDocfile(outputFile, STGM.READWRITE | STGM.TRANSACTED | STGM.CREATE, 0, out stg);
                if (hResult == HRESULT.S_OK)
                    hResult = MAPINative.OpenIMsgSession(malloc, 0, out pMsgSess);
                if (hResult == HRESULT.S_OK)
                    hResult = MAPINative.OpenIMsgOnIStg(pMsgSess, new AllocateBuffer(MAPINative.MAPIAllocateBuffer),
                                                    new AllocateMore(MAPINative.MAPIAllocateMore), new FreeBuffer(MAPINative.MAPIFreeBuffer),
                                                    malloc, IntPtr.Zero, stg, IntPtr.Zero, 0, 0, out msg);
                if (hResult == HRESULT.S_OK)
                    hResult = Storage.WriteClassStg(stg, IID_MailMessage);
                if (hResult == HRESULT.S_OK)
                {
                    hResult = HRESULT.E_FAIL;
                    if (CopyTo(
                        new uint[] { 
                (uint)PropTags.PR_ACCESS, (uint)PropTags.PR_BODY, (uint)PropTags.PR_RTF_SYNC_BODY_COUNT, (uint)PropTags.PR_RTF_SYNC_BODY_CRC, (uint)PropTags.PR_RTF_SYNC_BODY_TAG, (uint)PropTags.PR_RTF_SYNC_PREFIX_COUNT, (uint)PropTags.PR_RTF_SYNC_TRAILING_COUNT},
                        msg
                        ))
                    {
                        hResult = HRESULT.S_OK;
                        msg.SaveChanges((uint)SaveFlags.KEEP_OPEN_READWRITE);
                        stg.Commit((uint)STGC.DEFAULT);
                    }
                }
                if (stg != null)
                    Marshal.ReleaseComObject(stg);
                if (msg != null)
                    Marshal.ReleaseComObject(msg);
                if (pMsgSess != IntPtr.Zero)
                    MAPINative.CloseIMsgSession(pMsgSess);
            }
            return hResult == HRESULT.S_OK;
        }
        /// <summary>
        /// Add a file to an attachment of the message
        /// </summary>
        /// <param name="filePath">the file paty of the original file.</param>
        /// <param name="name">Display name of the new attachment</param>
        /// <returns>true, if the attachment was added successfully; otherwise, failed</returns>
        public bool AddAttachment(string filePath, string name)
        {
            uint attachmentNum;
            IntPtr pAttach;
            Attachment attachment = null;
            bool bResult = false;

            HRESULT hResult = Message.CreateAttach(IntPtr.Zero, 0, out attachmentNum, out pAttach);
            if (hResult == HRESULT.S_OK)
            {
                IAttach obj = Marshal.GetObjectForIUnknown(pAttach) as IAttach;
                if (obj != null)
                    attachment = new Attachment(obj);
            }
            if (attachment != null)
            {
                bResult = attachment.Attach(filePath, name);
                attachment.Dispose();
            }
            return bResult;
        }
        /// <summary>
        /// Add a MAPI Message as an embedded attachment
        /// </summary>
        /// <param name="message">MAPIMessage object</param>
        /// <param name="name">Display name of the new attachment</param>
        /// <returns>true, if the attachment was added successfully; otherwise, failed</returns>
        public bool AddAttachment(MAPIMessage message, string name)
        {
            uint attachmentNum;
            IntPtr pAttach;
            Attachment attachment = null;
            bool bResult = false;

            HRESULT hResult = Message.CreateAttach(IntPtr.Zero, 0, out attachmentNum, out pAttach);
            if (hResult == HRESULT.S_OK)
            {
                IAttach obj = Marshal.GetObjectForIUnknown(pAttach) as IAttach;
                if (obj != null)
                    attachment = new Attachment(obj);
            }
            if (attachment != null)
            {
                bResult = attachment.Attach(message, name);
                attachment.Dispose();
            }
            return bResult;
        }
        /// <summary>
        /// Get an attachment display name
        /// </summary>
        /// <param name="index">index of the attachment</param>
        /// <returns>display name of the attachment</returns>
        public string GetAttachmentName(int index)
        {
            int attachCount = AttachmentCount;
            if (index >= attachCount)
                return null;
            string attachName = null;
            using (MAPITable tb = GetAttachmentTable())
            {
                if (tb != null)
                {
                    tb.SetColumns(new PropTags[] { PropTags.PR_ATTACH_LONG_FILENAME, PropTags.PR_ATTACH_FILENAME });
                    SRow[] sRows;
                    if (tb.QueryRows(attachCount, out sRows) && index < sRows.Length)
                    {
                        attachName = sRows[index].propVals[0].AsString;
                        if (string.IsNullOrEmpty(attachName))
                            attachName = sRows[index].propVals[1].AsString;
                    }
                }
            }
            return attachName;
        }
        /// <summary>
        /// Save attachment to a disk file.
        /// </summary>
        /// <param name="index">index of the attachment</param>
        /// <returns>the path of the output file</returns>
        public string SaveAttachment(int index)
        {
            string folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return SaveAttachment(folder, index, null);
        }
        /// <summary>
        /// Save attachment to a disk file.
        /// </summary>
        /// <param name="folder">output folder</param>
        /// <param name="index">index of the attachment</param>
        /// <returns>the path of the output file</returns>
        public string SaveAttachment(string folder, int index)
        {
            return SaveAttachment(folder, index, null);
        }
        /// <summary>
        /// Save attachment to a disk file.
        /// </summary>
        /// <param name="folder">output folder</param>
        /// <param name="index">index of the attachment</param>
        /// <param name="fileName">output file name</param>
        /// <returns>the path of the output file</returns>
        public string SaveAttachment(string folder, int index, string fileName)
        {
            string path = string.Empty;
            int attachCount = AttachmentCount;
            if (index >= attachCount)
                return path;
            MAPITable tb = GetAttachmentTable();
            if (tb == null)
                return path;
            bool bResult = false;
            tb.SetColumns(new PropTags[] { PropTags.PR_ATTACH_NUM, PropTags.PR_ATTACH_LONG_FILENAME, PropTags.PR_ATTACH_FILENAME });
            SRow[] sRows;
            int i = 0;
            while (tb.QueryRows(1, out sRows))
            {
                if (i++ < index)
                    continue;
                uint attachNum = (uint)sRows[0].propVals[0].AsInt32;
                IntPtr pAttach;
                HRESULT hr = Message.OpenAttach(attachNum, IntPtr.Zero, 0, out pAttach);
                if (hr == HRESULT.S_OK && pAttach != null)
                {
                    Attachment attach = new Attachment(Marshal.GetObjectForIUnknown(pAttach) as IAttach);
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = sRows[0].propVals[1].AsString;
                        if (string.IsNullOrEmpty(fileName))
                            fileName = sRows[0].propVals[1].AsString;
                        if (string.IsNullOrEmpty(fileName))
                            fileName = "Attachment.Dat";

                    }
                    path = System.IO.Path.Combine(folder, fileName);
                    bResult = attach.SaveToDisk(path);
                    attach.Dispose();
                }

            }
            tb.Dispose();
            return path;
        }

        
        #endregion

        #region Private Properties/Methods/Events

        IMessage Message { get { return mapiObj_ as IMessage; } }

        MAPITable GetRecipientTable()
        {
            IntPtr pTable;
            MAPITable mb = null;
            HRESULT hResult = Message.GetRecipientTable((uint)CharacterSet.UNICODE, out pTable);
            if (hResult == HRESULT.S_OK)
            {
                IMAPITable tableObj = Marshal.GetObjectForIUnknown(pTable) as IMAPITable;
                if (tableObj != null)
                {
                    mb = new MAPITable(tableObj);
                }
            }
            return mb;
        }

        private MAPITable GetAttachmentTable()
        {
            IntPtr pTable;
            HRESULT hr = Message.GetAttachmentTable(0, out pTable);
            if (hr == HRESULT.S_OK && pTable != IntPtr.Zero)
                return new MAPITable(Marshal.GetObjectForIUnknown(pTable) as IMAPITable);
            return null;
        }

        private string GetRTFBody()
        {
            IntPtr pObject = OpenProperty((uint)PropTags.PR_RTF_COMPRESSED, Storage.IID_Stream, OpenPropertyMode.READ);
            StringBuilder rtf = new StringBuilder();
            if (pObject != null && pObject != IntPtr.Zero)
            {
                Ole.IStream compressedStream = Marshal.GetObjectForIUnknown(pObject) as Ole.IStream;
                Ole.IStream uncompressedStream;
                if (MAPINative.WrapCompressedRTFStream(compressedStream, 0, out uncompressedStream) == HRESULT.S_OK)
                {
                    byte[] bytes = new byte[BufSize];
                    int cbRead = 0;
                    do
                    {
                        cbRead = Storage.Read(uncompressedStream, bytes);
                        rtf.Append(Encoding.ASCII.GetChars(bytes));
                    } while (cbRead >= BufSize);
                    Marshal.ReleaseComObject(uncompressedStream);
                }
                Marshal.ReleaseComObject(compressedStream);
            }
            return rtf.ToString();
        }

        private void SetRTFHTMLBody(string html)
        {
            int codePage = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ANSICodePage;
            string rtf = "{\\rtf1\\ansi\\ansicpg" + codePage + "\\fromhtml1 {\\*\\htmltag1 ";
            rtf += html + " }}";
            SetRTFBody(rtf);
        }

        private void SetRTFBody(string value)
        {
            IntPtr pObject = OpenProperty((uint)PropTags.PR_RTF_COMPRESSED, Storage.IID_Stream, (uint)(STGM.CREATE | STGM.WRITE), OpenPropertyMode.MODIFY | OpenPropertyMode.CREATE);
            if (pObject != null && pObject != IntPtr.Zero)
            {
                Ole.IStream compressedStream = Marshal.GetObjectForIUnknown(pObject) as Ole.IStream;
                Ole.IStream uncompressedStream;
                if (MAPINative.WrapCompressedRTFStream(compressedStream, 1, out uncompressedStream) == HRESULT.S_OK)
                {
                    Byte [] bytes = Encoding.ASCII.GetBytes(value);
                    Storage.Write(uncompressedStream, bytes);
                    uncompressedStream.Commit((int)STGC.DEFAULT);
                    compressedStream.Commit((int)STGC.DEFAULT);
                    Marshal.ReleaseComObject(uncompressedStream);
                }
                Marshal.ReleaseComObject(compressedStream);
            }
        }

        #endregion

        #region IDisposable Interface

        /// <summary>
        /// Dispose MAPIMessage object
        /// </summary>
        public override void Dispose()
        {
            if (recipientTable_ != null)
                recipientTable_.Dispose();
            base.Dispose();
        }

        #endregion

    }
}
