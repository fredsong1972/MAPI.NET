using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using MAPI.NET.Properties;

namespace MAPI.NET
{
    /// <summary>
    /// Provides access to message store information and to messages and folders.
    /// </summary>
    [
        ComImport, ComVisible(false),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("00020306-0000-0000-C000-000000000046")
    ]
    public interface IMsgStore : IMAPIProp
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
        /// Registers to receive notification of specified events that affect the message store.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the folder or message about which notifications should be generated, or null. If lpEntryID is set to NULL, Advise registers for notifications on the entire message store.</param>
        /// <param name="ulEventMask">A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration.</param>
        /// <param name="pAdviseSink">A pointer to an advise sink object to receive the subsequent notifications.</param>
        /// <param name="lpulConnection">A pointer to a nonzero number that represents the connection between the caller's advise sink object and the session.</param>
        /// <returns>S_OK, if the registration was successful; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT Advise(uint cbEntryID, IntPtr lpEntryID, uint ulEventMask, [In, MarshalAs(UnmanagedType.Interface)] IMAPIAdviseSink pAdviseSink, out uint lpulConnection);
        /// <summary>
        /// Cancels the sending of notifications previously set up with a call to the IMsgStore::Advise method.
        /// </summary>
        /// <param name="ulConnection">The connection number associated with an active notification registration.</param>
        /// <returns>S_OK, if the registration was successfully canceled; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT Unadvise(uint ulConnection);
        /// <summary>
        /// Compares two entry identifiers to determine whether they refer to the same entry in a message store.
        /// </summary>
        /// <param name="cbEntryID1">The byte count in the entry identifier pointed to by the lpEntryID1 parameter.</param>
        /// <param name="lpEntryID1">A pointer to the first entry identifier to be compared.</param>
        /// <param name="cbEntryID2">The byte count in the entry identifier pointed to by the lpEntryID2 parameter.</param>
        /// <param name="lpEntryID2">A pointer to the second entry identifier to be compared.</param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpulResult">true if the two entry identifiers refer to the same object; otherwise, false</param>
        /// <returns>S_OK, if the comparison was successful; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT CompareEntryIDs(uint cbEntryID1, IntPtr lpEntryID1, uint cbEntryID2, IntPtr lpEntryID2, uint ulFlags, out bool lpulResult);
        /// <summary>
        /// Opens a folder or message and returns an interface pointer for further access. 
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the object to open, or NULL. If lpEntryID is set to NULL, OpenEntry opens the root folder for the message store.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the opened object. </param>
        /// <param name="ulFlags">A bitmask of flags that controls how the object is opened.</param>
        /// <param name="lpulObjType">A pointer to the type of the opened object.</param>
        /// <param name="lppUnk">A pointer to a pointer to the opened object.</param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed</returns>
        [PreserveSig]
        HRESULT OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
        /// <summary>
        /// Establishes a folder as the destination for incoming messages of a particular message class.
        /// </summary>
        /// <param name="lpszMessageClass">A message class is associated with the new receive folder</param>
        /// <param name="ulFlags">A bitmask of flags that controls the type of the text in the passed-in strings.</param>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">pointer to the entry identifier of the folder to establish as the receive folder.If the lpEntryID parameter is set to NULL, SetReceiveFolder replaces the current receive folder with the message store's default.</param>
        /// <returns>S_OK, if a receive folder was successfully established; otherwise, failed</returns>
        [PreserveSig]
        HRESULT SetReceiveFolder(string lpszMessageClass, uint ulFlags, uint cbEntryID, IntPtr lpEntryID);
        /// <summary>
        /// Obtains the folder that was established as the destination for incoming messages of a specified message class or as the default receive folder for the message store.
        /// </summary>
        /// <param name="lpszMessageClass">A message class is associated with a receive folder</param>
        /// <param name="ulFlags">A bitmask of flags that controls the type of the passed-in and returned strings. </param>
        /// <param name="cbEntryID">A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</param>
        /// <param name="lppEntryID">A pointer to a pointer to the entry identifier for the requested receive folder.</param>
        /// <param name="lppszExplicitClass">A pointer to a pointer to the message class that explicitly sets as its receive folder the folder pointed to by lppEntryID.</param>
        /// <returns>S_OK, if the receive folder was successfully returned; otherwise, failed</returns>
        [PreserveSig]
        HRESULT GetReceiveFolder([MarshalAs(UnmanagedType.LPWStr)]string lpszMessageClass, uint ulFlags, out uint cbEntryID, out IntPtr lppEntryID, [MarshalAs(UnmanagedType.LPWStr)]StringBuilder lppszExplicitClass);
        /// <summary>
        /// Attempts to remove a message from the outgoing queue.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the message to remove from the outgoing queue. </param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <returns>S_OK, if the message was successfully removed from the outgoing queue; otherwise, failed</returns>
        [PreserveSig]
        HRESULT AbortSubmit(uint cbEntryID, IntPtr lpEntryID, uint ulFlags);
    }

    /// <summary>
    /// IMsgStore .Net Wrapper object
    /// </summary>
    public class MsgStore : MAPIObject
    {
        /// <summary>
        /// New mail event
        /// </summary>
        public event EventHandler<MsgStoreNewMailEventArgs> OnNewMail;
        /// <summary>
        /// Object created event
        /// </summary>
        public event EventHandler<MsgStoreObjectEventArgs> OnObjectCreated;
        /// <summary>
        /// Object deleted event
        /// </summary>
        public event EventHandler<MsgStoreObjectEventArgs> OnObjectDeleted;
        /// <summary>
        /// Object modified event
        /// </summary>
        public event EventHandler<MsgStoreObjectEventArgs> OnObjectModified;
        /// <summary>
        /// Object moved event
        /// </summary>
        public event EventHandler<MsgStoreObjectEventArgs> OnObjectMoved;
        /// <summary>
        /// Object copied event
        /// </summary>
        public event EventHandler<MsgStoreObjectEventArgs> OnObjectCopied;

        static Guid IID_IMsgStore = new Guid("00020306-0000-0000-C000-000000000046");

        OnAdviseCallbackHandler callbackHandler_;
        uint ulConnection_ = 0;
        IMAPIAdviseSink pAdviseSink_ = null;

        /// <summary>
        /// Initializes a new instance of the MsgStore class. 
        /// </summary>
        /// <param name="msgStore">IMsgStore object</param>
        /// <param name="entryID">Entry identification of IMsgStore object</param>
        public MsgStore(MAPISession session, IMsgStore msgStore, EntryID entryID)
            : base(msgStore, entryID)
        {
            ulConnection_ = 0;
            Session = session;
        }

        #region Public Properties
        /// <summary>
        /// Gets entry identificatio of Message Store
        /// </summary>
        public EntryID StoreID { get { return Id_; } }
        /// <summary>
        /// Gets IMsgStore interface GUID
        /// </summary>
        public override Guid InterfaceID
        {
            get
            {
                return IID_IMsgStore;
            }
        }


        public MAPISession Session
        { get; private set; }


        public EntryID InboxEntryID
        {
            get
            {
                uint cb;
                IntPtr lpb;
                EntryID entry = null;
                HRESULT hResult = MAPIStore.GetReceiveFolder("", 0, out cb, out lpb, new StringBuilder());
                if (hResult == HRESULT.S_OK && cb > 0 && lpb != null && lpb != IntPtr.Zero)
                {
                    entry = EntryID.BuildFromPtr(cb, lpb);
                    MAPINative.MAPIFreeBuffer(lpb);
                }
                return entry;
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Open root folder
        /// </summary>
        /// <returns>a MAPIFolder object which represents the root folder</returns>
        public MAPIFolder OpenRootFolder()
        {
            return OpenFolder((uint)PropTags.PR_IPM_SUBTREE_ENTRYID);
        }
        /// <summary>
        /// Open inbox
        /// </summary>
        /// <returns>a MAPIFolder object which represents the inbox</returns>
        public MAPIFolder OpenInbox()
        {
            uint cb;
            IntPtr lpb;
            MAPIFolder mapiFolder = null;
            HRESULT hResult = MAPIStore.GetReceiveFolder("", 0, out cb, out lpb, new StringBuilder());
            if (hResult == HRESULT.S_OK)
            {
                mapiFolder = OpenFolder(cb, lpb);
                MAPINative.MAPIFreeBuffer(lpb);
            }
            return mapiFolder;
        }
        /// <summary>
        /// Open outbox
        /// </summary>
        /// <returns>a MAPIFolder object which represents the outbox</returns>
        public MAPIFolder OpenOutbox()
        {
            return OpenFolder((uint)PropTags.PR_IPM_OUTBOX_ENTRYID, MAPIObject.MailItem);
        }
        /// <summary>
        /// Open draft folder
        /// </summary>
        /// <returns>a MAPIFolder object which represents the draft folder</returns>
        public MAPIFolder OpenDraft()
        {
            return OpenSpecialFolder((uint)PropTags.PR_IPM_DRAFTS_ENTRYID, MAPIObject.MailItem);
        }
        /// <summary>
        /// Open Sent folder
        /// </summary>
        /// <returns>a MAPIFolder object which represents the sent folder</returns>
        public MAPIFolder OpenSentItems()
        {
            return OpenFolder((uint)PropTags.PR_IPM_SENTMAIL_ENTRYID, MAPIObject.MailItem);
        }
        /// <summary>
        /// Open deleted items folder
        /// </summary>
        /// <returns>a MAPIFolder object which represents the deleted items</returns>
        public MAPIFolder OpenDeletedItems()
        {
            return OpenFolder((uint)PropTags.PR_IPM_WASTEBASKET_ENTRYID, MAPIObject.MailItem);
        }

        public MAPIFolder OpenCalendar()
        {
            return OpenSpecialFolder((uint)PropTags.PR_IPM_APPOINTMENT_ENTRYID, MAPIObject.AppointmentItem);
        }

        public MAPIFolder OpenContacts()
        {
            return OpenSpecialFolder((uint)PropTags.PR_IPM_CONTACT_ENTRYID, MAPIObject.ContactItem);
        }


        public MAPIFolder OpenJunkFolder()
        {
            return OpenFolder("Junk E-mail");
        }

        /// <summary>
        /// Open a MAPI folder
        /// </summary>
        /// <param name="entryId">the entry identifier of the folder to open</param>
        /// <returns>a MAPIFolder</returns>
        public MAPIFolder OpenFolder(EntryID entryId)
        {
            MAPIFolder mapiFolder = null;
            SBinary id = SBinary.SBinaryCreate(entryId.AsByteArray);
            try
            {
                mapiFolder = OpenFolder(id.cb, id.lpb);
            }
            finally
            {
                SBinary.SBinaryRelease(ref id);
            }
            return mapiFolder;
        }
        /// <summary>
        /// Registers to receive notification of specified events that affect the message store.
        /// </summary>
        /// <param name="eventMask">A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration. </param>
        /// <returns></returns>
        public bool RegisterEvents(EEventMask eventMask)
        {
            callbackHandler_ = new OnAdviseCallbackHandler(OnNotifyCallback);
            HRESULT hresult = HRESULT.S_OK;
            try
            {
                pAdviseSink_ = new MAPIAdviseSink(IntPtr.Zero, callbackHandler_);
                hresult = MAPIStore.Advise(0, IntPtr.Zero, (uint)eventMask, pAdviseSink_, out ulConnection_);

            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
                return false;
            }
            return hresult == HRESULT.S_OK;
        }

        /// <summary>
        /// Cancels the sending of notifications.
        /// </summary>
        public void UnRegisteEvents()
        {
            if (ulConnection_ != 0)
                MAPIStore.Unadvise(ulConnection_);
        }


        public MAPIMessage CreateMessage()
        {
            using (MAPIFolder folder = OpenOutbox())
            {
                MAPIMessage message = folder.CreateMessage();
                return message;
            }
        }

     
        public Appointment CreateAppointment()
        {
            using (MAPIFolder folder = OpenCalendar())
            {
                Appointment message = folder.CreateMessage() as Appointment;
                return message;
            }
        }

        public Contact CreateContact()
        {
            using (MAPIFolder folder = OpenContacts())
            {
                Contact message = folder.CreateMessage() as Contact;
                return message;
            }
        }

        /// <summary>
        /// Gets MAPI object per the entry identification
        /// </summary>
        /// <param name="id">The entry identification of the object to get</param>
        /// <returns>MAPIObject instance</returns>
        public MAPIObject GetMAPIObject(EntryID id)
        {
            IMAPIProp prop = OpenEntry(id);
            if (prop == null)
                return null;
            if (prop is IMessage)
            {
                IPropValue messageClass = MAPISession.GetObjectProperty(prop, (uint)PropTags.PR_MESSAGE_CLASS);
                if (messageClass == null)
                    return new MAPIMessage(prop as IMessage);
                return GetMAPIMessage(id, 0, messageClass.AsString);
            }
            else if (prop is IMAPIFolder)
                return new MAPIFolder(prop as IMAPIFolder);
            else
                return new MAPIObject(prop);
        }

        /// <summary>
        /// Gets MAPI Message per the entry identification
        /// </summary>
        /// <param name="id">The entry identification of the message to get</param>
        /// <param name="messageFlags">The message flags which is assocaited with the message</param>
        /// <param name="messageClass">The message class which is associated with the message</param>
        /// <returns>MAPI Message instance</returns>
        public MAPIMessage GetMAPIMessage(EntryID id, uint messageFlags, string messageClass)
        {
            IMessage message = OpenEntry(id) as IMessage;
            if (message == null)
                return null;
            switch (messageClass)
            {
                case MAPIObject.AppointmentItem:
                    return new Appointment(message);
                case MAPIObject.ContactItem:
                    return new Contact(message);
            }
            return new MAPIMessage(message, messageFlags, messageClass);
        }

        /// <summary>
        /// Opens an object and returns an interface pointer for additional access.
        /// </summary>
        /// <param name="id">The entry identifier of the object to open.</param>
        /// <returns>The IMAPIProp object which was opened.</returns>
        public IMAPIProp OpenEntry(EntryID id)
        {
            IntPtr pObject;
            uint objectType;
            IMAPIProp prop = null;
            SBinary sb = SBinary.SBinaryCreate(id.AsByteArray);
            try
            {
                HRESULT hResult = MAPIStore.OpenEntry(sb.cb, sb.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objectType, out pObject);
                if (hResult == HRESULT.S_OK && pObject != null)
                {
                    prop = Marshal.GetObjectForIUnknown(pObject) as IMAPIProp;
                }
            }
            catch { }
            finally
            {
                SBinary.SBinaryRelease(ref sb);
            }
            return prop;
        }

        public bool SendMessage(MAPIMessage message, bool saveToSentFolder)
        {
            message.SaveToSentFolder(this, saveToSentFolder);
            return message.Send();
        }

        public MAPIMessage ForwardMessage(MAPIMessage message)
        {
            MAPIMessage forward = CreateMessage();
            if (message.CopyTo(new uint[] 
            {
                (uint)PropTags.PR_ACCESS, (uint)PropTags.PR_RTF_SYNC_BODY_COUNT, (uint)PropTags.PR_RTF_SYNC_BODY_CRC, (uint)PropTags.PR_RTF_SYNC_BODY_TAG, (uint)PropTags.PR_RTF_SYNC_PREFIX_COUNT, (uint)PropTags.PR_RTF_SYNC_TRAILING_COUNT,
                (uint)PropTags.PR_SUBJECT, (uint)PropTags.PR_MESSAGE_RECIPIENTS
            }, forward))
            {

                Dictionary<PropTags, IPropValue> values = message.GetProperties(new PropTags[] { PropTags.PR_SENDER_NAME, PropTags.PR_SENDER_EMAIL_ADDRESS, PropTags.PR_CLIENT_SUBMIT_TIME, PropTags.PR_DISPLAY_TO, PropTags.PR_SUBJECT});
                string body = forward.Body;
                string subject = values.ContainsKey(PropTags.PR_SUBJECT) ? values[PropTags.PR_SUBJECT].AsString : "";
                forward.Subject = "FW: " + subject;
                string to = values.ContainsKey(PropTags.PR_DISPLAY_TO) ? values[PropTags.PR_DISPLAY_TO].AsString : "";
                string sender = values.ContainsKey(PropTags.PR_SENDER_NAME) ? values[PropTags.PR_SENDER_NAME].AsString : "";
                string senderAddress = values.ContainsKey(PropTags.PR_SENDER_EMAIL_ADDRESS) ? values[PropTags.PR_SENDER_EMAIL_ADDRESS].AsString : "";
                string sent = "";
                if (values.ContainsKey(PropTags.PR_CLIENT_SUBMIT_TIME))
                {
                    DateTime? sentTime = values[PropTags.PR_CLIENT_SUBMIT_TIME].AsDateTime;
                   sent = (sentTime.HasValue ? sentTime.Value.ToString("g") : "");
                }
                string forwardHeader = "";
                switch (forward.MessageFormat)
                {
                    case MessageFormat.PLAINTEXT:
                        forwardHeader = "\r\n" + "\r\n" + string.Format(Resource.textForwardHeader, sender, senderAddress, sent, to, subject) + "\r\n";
                        forward.Body = forwardHeader + body;
                        break;
                    case MessageFormat.DONTKNOW:
                    case MessageFormat.RTF:
                        forwardHeader = string.Format(Resource.rtfForwardHeader, sender, senderAddress, sent, to, subject) + "\r\n";
                        forward.Body = forwardHeader + body;
                        break;
                    case MessageFormat.RTFHTML:
                    case MessageFormat.HTML:
                        {
                            forwardHeader = string.Format(Resource.htmlForwardHeader, sender, senderAddress, sent, to, subject);
                            
                            Regex rx = new Regex("<div [\\w\\d=\"]*>\n*");
                            Match match = rx.Match(body);
                            if (match.Success)
                            {
                                body = body.Substring(0, match.Index) + forwardHeader + body.Substring(match.Index + match.Length);
                                forward.Body = body;
                            }
                            break;
                        }
                    //case MessageFormat.RTFHTML:
                    //    {
                    //        forwardHeader = string.Format(Properties.Resources.rtfHtmlForwardHeader, sender, senderAddress, sent, to, subject);
                    //        forwardHeader = forwardHeader.Replace('[', '}');
                    //        forwardHeader = forwardHeader.Replace(']', '}');
                    //        int index = body.IndexOf("<div");
                    //        index = body.IndexOf('\n', index);
                    //        forward.Body = body.Substring(0, index + 1) + forwardHeader + body.Substring(index + 1);
                    //        break;
                    //    }
                }
            }
            return forward;

        }
        #endregion

        #region Private Properties/Methods/Events

        IMsgStore MAPIStore
        {
            get { return mapiObj_ as IMsgStore; }
        }

        void OnNotifyCallback(IntPtr pMAPI, uint cNotification, IntPtr lpNotifications)
        {
            EEventMask eventType = (EEventMask)Marshal.ReadInt32(lpNotifications);
            int intSize = Marshal.SizeOf(typeof(int));
            IntPtr sPtr = lpNotifications + intSize * 2; //ulEventType, ulAlignPad

            switch (eventType)
            {
                case EEventMask.fnevNewMail:
                    {
                        if (this.OnNewMail == null)
                            break;
                        NEWMAIL_NOTIFICATION notification = (NEWMAIL_NOTIFICATION)Marshal.PtrToStructure(sPtr, typeof(NEWMAIL_NOTIFICATION));
                        MsgStoreNewMailEventArgs n = new MsgStoreNewMailEventArgs(StoreID, notification);
                        this.OnNewMail(this, n);
                    }
                    break;
                case EEventMask.fnevObjectCopied:
                case EEventMask.fnevObjectCreated:
                case EEventMask.fnevObjectDeleted:
                case EEventMask.fnevObjectModified:
                case EEventMask.fnevObjectMoved:
                    {
                        EventHandler<MsgStoreObjectEventArgs> handler = null;

                        if (eventType == EEventMask.fnevObjectCopied && OnObjectCopied != null)
                            handler = OnObjectCopied;
                        else if (eventType == EEventMask.fnevObjectCreated && OnObjectCreated != null)
                            handler = OnObjectCreated;
                        else if (eventType == EEventMask.fnevObjectDeleted && this.OnObjectDeleted != null)
                            handler = OnObjectDeleted;
                        else if (eventType == EEventMask.fnevObjectModified && this.OnObjectModified != null)
                            handler = OnObjectModified;
                        else if (eventType == EEventMask.fnevObjectMoved && this.OnObjectMoved != null)
                            handler = OnObjectMoved;

                        if (handler == null)
                            break;

                        OBJECT_NOTIFICATION notification = (OBJECT_NOTIFICATION)Marshal.PtrToStructure(sPtr, typeof(OBJECT_NOTIFICATION));
                        MsgStoreObjectEventArgs o = new MsgStoreObjectEventArgs(StoreID, notification);
                        handler(this, o);
                    }
                    break;
            }
        }

        private MAPIFolder OpenFolder(uint folderID)
        {
            return OpenFolder(folderID, null);
        }

        private MAPIFolder OpenSpecialFolder(uint folderID, string defaultItemType)
        {
            MAPIFolder mapiFolder = null;
            uint cb;
            IntPtr lpb;
            IMAPIFolder inbox = null;
            HRESULT hResult = MAPIStore.GetReceiveFolder("", 0, out cb, out lpb, new StringBuilder());
            if (hResult == HRESULT.S_OK)
            {
                uint objType;
                IntPtr pObj;
                hResult = MAPIStore.OpenEntry(cb, lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pObj);
                if (hResult == HRESULT.S_OK && pObj != null)
                {
                    inbox = Marshal.GetObjectForIUnknown(pObj) as IMAPIFolder;
                }
                MAPINative.MAPIFreeBuffer(lpb);
            }
            if (inbox != null)
            {
                IPropValue prop = MAPISession.GetObjectProperty(inbox, folderID);
                if (prop.AsBinary != null)
                {
                    SBinary entryId = SBinary.SBinaryCreate(prop.AsBinary);
                    uint objType;
                    IntPtr pObj;
                    try
                    {
                        hResult = MAPIStore.OpenEntry(entryId.cb, entryId.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pObj);
                        if (hResult == HRESULT.S_OK && pObj != null)
                        {
                            IMAPIFolder folder = Marshal.GetObjectForIUnknown(pObj) as IMAPIFolder;
                            if (folder != null)
                            {
                                if (string.IsNullOrEmpty(defaultItemType))
                                    mapiFolder = new MAPIFolder(folder);
                                else
                                    mapiFolder = new MAPIFolder(folder, defaultItemType);
                            }
                        }
                    }
                    finally
                    {
                        SBinary.SBinaryRelease(ref entryId);
                    }
                }
                Marshal.ReleaseComObject(inbox);
            }
            return mapiFolder;
        }

        private MAPIFolder OpenFolder(uint folderID, string defaultItemType)
        {
            MAPIFolder mapiFolder = null;
            uint objType;
            IntPtr pObj;
            IPropValue prop = this.GetProperty(folderID);
            if (prop != null && prop.AsBinary != null)
            {
                SBinary entryId = SBinary.SBinaryCreate(prop.AsBinary);
                try
                {
                    HRESULT hResult = MAPIStore.OpenEntry(entryId.cb, entryId.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pObj);
                    if (hResult == HRESULT.S_OK && pObj != null)
                    {
                        IMAPIFolder folder = Marshal.GetObjectForIUnknown(pObj) as IMAPIFolder;
                        if (folder != null)
                        {
                            if (string.IsNullOrEmpty(defaultItemType))
                                mapiFolder = new MAPIFolder(folder);
                            else
                                mapiFolder = new MAPIFolder(folder, defaultItemType);
                        }
                    }
                }
                finally
                {
                    SBinary.SBinaryRelease(ref entryId);
                }
            }
            return mapiFolder;
        }
        
        private MAPIFolder OpenFolder(uint cb, IntPtr lpb)
        {
            MAPIFolder mapiFolder = null;
            uint objType;
            IntPtr pObj;
            HRESULT hResult = MAPIStore.OpenEntry(cb, lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pObj);
            if (hResult == HRESULT.S_OK && pObj != null)
            {
                IMAPIFolder folder = Marshal.GetObjectForIUnknown(pObj) as IMAPIFolder;
                if (folder != null)
                    mapiFolder = new MAPIFolder(folder);
            }
            return mapiFolder;
        }

        private MAPIFolder OpenFolder(string name)
        {
            MAPIFolder root = OpenRootFolder();
            if (root != null)
            {
                MAPIFolder folder = root.OpenSubFolder(name);
                root.Dispose();
                return folder;
            }
            return null;
         }

        #endregion

        #region IDisposable Interface
        /// <summary>
        /// Dispose Msgstore object
        /// </summary>
        public override void Dispose()
        {
            UnRegisteEvents();
            Session = null;
            if (mapiObj_ != null)
            {
                Marshal.ReleaseComObject(mapiObj_);
                mapiObj_ = null;
            }
        }

        #endregion


    }


}
