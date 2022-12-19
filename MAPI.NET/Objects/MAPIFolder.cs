using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MAPI.NET
{
    /// <summary>
    /// Manages high-level operations on container objects such as address books, distribution lists, and folders.
    /// </summary>
    [ComVisible(false)]
    [ComImport()]
    [Guid("0002030B-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMAPIContainer : IMAPIProp
    {
        /// <summary>
        /// Returns a pointer to the container's contents table.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls how the contents table is returned.</param>
        /// <param name="pTable">A pointer to a pointer to the contents table.</param>
        /// <returns>S_OK, if the contents table was successfully retrieved; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT GetContentsTable(uint ulFlags, out IntPtr pTable);
        /// <summary>
        /// Returns a pointer to the container's hierarchy table.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls how information is returned in the table. </param>
        /// <param name="pTable">A pointer to a pointer to the hierarchy table.</param>
        /// <returns>S_OK, if the hierarchy table was successfully retrieved; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT GetHierarchyTable(uint ulFlags, out IntPtr pTable);
        /// <summary>
        /// Opens an object in the container, returning an interface pointer for further access.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the object to open. If lpEntryID is set to NULL, the top-level container in the container's hierarchy is opened.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the object. </param>
        /// <param name="ulFlags">A bitmask of flags that controls how the object is opened.</param>
        /// <param name="lpulObjType">A pointer to the opened object's type.</param>
        /// <param name="lppUnk">A pointer to a pointer to the interface implementation to use to access the open object.</param>
        /// <returns>S_OK, if the object was successfully opened; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
        /// <summary>
        /// Establishes search criteria for the container.
        /// </summary>
        /// <param name="pRestriction">A pointer to an SRestriction structure that defines the search criteria.</param>
        /// <param name="pContainerList">A pointer to an array of entry identifiers that represent containers to be included in the search.</param>
        /// <param name="ulSearchFlags">A bitmask of flags that control how the search is performed.</param>
        /// <returns>S_OK, if the search criteria was successfully set; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT SetSearchCriteria(IntPtr pRestriction, IntPtr pContainerList, uint ulSearchFlags);
        /// <summary>
        /// Obtains the search criteria for the container.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls the type of the passed-in strings.</param>
        /// <param name="pRestriction">A pointer to a pointer to an SRestriction structure that defines the search criteria.</param>
        /// <param name="pContainerList"> A pointer to a pointer to an array of entry identifiers that represent containers to be included in the search.</param>
        /// <param name="searchState">A pointer to a bitmask of flags used to indicate the current state of the search</param>
        /// <returns>S_OK, if the search criteria was successfully obtained; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT GetSearchCriteria(uint ulFlags, out IntPtr pRestriction, out IntPtr pContainerList, out uint searchState);

    }

    /// <summary>
    /// PerPerforms operations on the messages and subfolders in a folder.forms operations on the messages and subfolders in a folder.
    /// </summary>
    [
       ComImport, ComVisible(false),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
       Guid("0002030C-0000-0000-C000-000000000046")
    ]
    public interface IMAPIFolder : IMAPIContainer
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
        /// <exclude/>
        new HRESULT GetContentsTable(uint ulFlags, out IntPtr pTable);
        /// <exclude/>
        new HRESULT GetHierarchyTable(uint ulFlags, out IntPtr pTable);
        /// <exclude/>
        new HRESULT OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
        /// <exclude/>
        new HRESULT SetSearchCriteria(IntPtr pRestriction, IntPtr pContainerList, uint ulSearchFlags);
        /// <exclude/>
        new HRESULT GetSearchCriteria(uint ulFlags, out IntPtr pRestriction, out IntPtr pContainerList, out uint searchState);
        /// <summary>
        /// Creates a new message.
        /// </summary>
        /// <param name="lpInterface"> A pointer to the interface identifier (IID) that represents the interface to be used to access the new message.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the message is created.</param>
        /// <param name="lppMessage">A pointer to a pointer to the newly created message.</param>
        /// <returns>S_OK, if the message was successfully created; otherwise, failed.</returns>
        HRESULT CreateMessage(IntPtr lpInterface, uint ulFlags, out IntPtr lppMessage);
        /// <summary>
        /// Copies or moves one or more messages.
        /// </summary>
        /// <param name="lpMsgList">A pointer to an array of ENTRYLIST structures that identify the message or messages to copy or move.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the destination folder pointed to by the lpDestFolder parameter. </param>
        /// <param name="lpDestFolder"> A pointer to the open folder to receive the copied or moved messages.</param>
        /// <param name="ulUIParam"> A handle to the parent window of any dialog boxes or windows this method displays. </param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator. </param>
        /// <param name="ulFlags"> A bitmask of flags that controls how the copy or move operation is accomplished. </param>
        /// <returns>S_OK, if the message or messages have been successfully copied or moved; otherwise, failed.</returns>
        HRESULT CopyMessages(IntPtr lpMsgList, IntPtr lpInterface, IntPtr lpDestFolder, uint ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Deletes one or more messages.
        /// </summary>
        /// <param name="lpMsgList">A pointer to an ENTRYLIST structure that contains the number of messages to delete and an array of ENTRYID structures that identify the messages.</param>
        /// <param name="ulUIParam">A handle to the parent window of the progress indicator. The ulUIParam parameter is ignored unless the MESSAGE_DIALOG flag is set in the ulFlags parameter.</param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the messages are deleted.</param>
        /// <returns>S_OK, if the specified message or messages were successfully deleted; otherwise, failed.</returns>
        HRESULT DeleteMessages(IntPtr lpMsgList, uint ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Creates a new subfolder.
        /// </summary>
        /// <param name="ulFolderType">The type of folder to create.</param>
        /// <param name="lpszFolderName">The name for the new folder.</param>
        /// <param name="lpszFolderComment">A comment associated with the new folder.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the new folder. Passing NULL causes the message store provider to return the standard folder interface, IMAPIFolder : IMAPIContainer.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls how the folder is created.</param>
        /// <param name="lppFolder">A pointer to a pointer to the newly created folder.</param>
        /// <returns>S_OK, if the new folder has been successfully created or opened; otherwise, failed.</returns>
        HRESULT CreateFolder(uint ulFolderType, [MarshalAs(UnmanagedType.LPWStr)] string lpszFolderName, [MarshalAs(UnmanagedType.LPWStr)] string lpszFolderComment, IntPtr lpInterface, uint ulFlags, out IntPtr lppFolder);
        /// <summary>
        /// Copies or moves a subfolder.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the subfolder to copy or move.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the folder that the lpDestFolder parameter points to. </param>
        /// <param name="lpDestFolder">A pointer to the open folder to receive the copied or moved subfolder.</param>
        /// <param name="lpszNewFolderName"> The name of the copied or moved folder in its new destination.</param>
        /// <param name="ulUIParam">A handle to the parent window of the progress indicator. </param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the copy or move operation. </param>
        /// <returns>S_OK, if the specified folder has been successfully copied or moved; otherwise, failed.</returns>
        HRESULT CopyFolder(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, IntPtr lpDestFolder, [MarshalAs(UnmanagedType.LPWStr)] string lpszNewFolderName, IntPtr ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Deletes a subfolder.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the subfolder to delete.</param>
        /// <param name="ulUIParam">A handle to the parent window of the progress indicator.</param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator. </param>
        /// <param name="ulFlags">A bitmask of flags that controls the deletion of the subfolder.</param>
        /// <returns>S_OK, if the specified folder has been successfully deleted; otherwise, failed.</returns>
        HRESULT DeleteFolder(uint cbEntryID, IntPtr lpEntryID, uint ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of one or more of the folder's messages, and manages the sending of read reports.
        /// </summary>
        /// <param name="lpMsgList">A pointer to an array of ENTRYLIST structures that identify the message or messages that have read flags to set or clear. </param>
        /// <param name="ulUIParam">A handle to the parent window of the progress indicator. </param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls the setting of a message's read flag and the processing of read reports. </param>
        /// <returns>S_OK, if the read flag for the specified message or messages was successfully set or cleared; otherwise, failed.</returns>
        HRESULT SetReadFlags(IntPtr lpMsgList, uint ulUIParam, IntPtr lpProgress, uint ulFlags);
        /// <summary>
        /// Obtains the status associated with a message in a particular folder.
        /// </summary>
        /// <param name="cbEntryID"> The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID"> A pointer to the entry identifier for the message whose status is obtained.</param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpulMessageStatus">A pointer to a pointer to a bitmask of flags that indicate the message's status. </param>
        /// <returns>S_OK, if the message status was successfully retrieved; otherwise, failed.</returns>
        HRESULT GetMessageStatus(uint cbEntryID, IntPtr lpEntryID, uint ulFlags, out uint lpulMessageStatus);
        /// <summary>
        /// Sets the status associated with a message.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier for the message whose status is set.</param>
        /// <param name="ulNewStatus">The new status to be assigned.</param>
        /// <param name="ulNewStatusMask">A bitmask of flags that is applied to the new status and indicates the flags to be set.</param>
        /// <param name="lpulOldStatus">A pointer to the previous status of the message.</param>
        /// <returns>S_OK, if the message status was successfully set; otherwise, failed.</returns>
        HRESULT SetMessageStatus(uint cbEntryID, IntPtr lpEntryID, uint ulNewStatus, uint ulNewStatusMask, out uint lpulOldStatus);
        /// <summary>
        /// Sets the default sort order for a folder's contents table.
        /// </summary>
        /// <param name="lpSortCriteria">] A pointer to an SSortOrderSet structure that contains the default sort order.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls how the default sort order is set. </param>
        /// <returns>S_OK, if the sort order was successfully saved; otherwise, failed.</returns>
        HRESULT SaveContentsSort(IntPtr lpSortCriteria, uint ulFlags);
        /// <summary>
        /// Deletes all messages and subfolders from a folder without deleting the folder itself.
        /// </summary>
        /// <param name="ulUIParam"> A handle to the parent window of the progress indicator. </param>
        /// <param name="lpProgress">A pointer to a progress object that displays a progress indicator.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the folder is emptied.</param>
        /// <returns>S_OK, if the folder was successfully emptied; otherwise, failed.</returns>
        HRESULT EmptyFolder(uint ulUIParam, IntPtr lpProgress, uint ulFlags);
    }
    /// <summary>
    /// IMAPIFolder .Net Wrapper object
    /// </summary>
    public class MAPIFolder : MAPIObject
    {
        static Guid IID_IMAPIFolder = new Guid("0002030C-0000-0000-C000-000000000046");
        const string MailContainer = "IPF.Note";
        const string ContactContainer = "IPF.Contact";
        const string AppointmentContainer = "IPF.Appointment";

        string name_ = string.Empty;
        string defaultItemType_ = MAPIObject.MailItem;
        MAPITable hierarchy_;
        
        /// <summary>
        /// Initializes a new instance of the MAPIFolder class. 
        /// </summary>
        /// <param name="folder">IMAPIFolder object</param>
        public MAPIFolder(IMAPIFolder folder)
            : base(folder)
        {
            Dictionary<PropTags, IPropValue> values = GetProperties(new PropTags[] { PropTags.PR_DISPLAY_NAME, PropTags.PR_CONTAINER_CLASS });
            if (values.ContainsKey(PropTags.PR_DISPLAY_NAME))
                name_ = values[PropTags.PR_DISPLAY_NAME].AsString;
            if (values.ContainsKey(PropTags.PR_CONTAINER_CLASS))
            {
                string containerClass = values[PropTags.PR_CONTAINER_CLASS].AsString;
                switch (containerClass)
                {
                    case MailContainer:
                        defaultItemType_ = MAPIObject.MailItem;
                        break;
                    case AppointmentContainer:
                        defaultItemType_ = MAPIObject.AppointmentItem;
                        break;
                    case ContactContainer:
                        defaultItemType_ = MAPIObject.ContactItem;
                        break;
                }
             }
          
        }

        public MAPIFolder(IMAPIFolder folder, string defaultItemType) : base(folder)
        {
            IPropValue prop = GetProperty(PropTags.PR_DISPLAY_NAME);
            name_ = prop.AsString;
            defaultItemType_ = defaultItemType;
        }


        #region Public Properties
        /// <summary>
        /// Gets or sets folder display name. 
        /// </summary>
        public string Name
        {
            get
            {
                if (name_ == null)
                {
                    IPropValue val = GetProperty(PropTags.PR_DISPLAY_NAME);
                    if (val != null)
                        name_ = val.AsString;
                }
                return name_;
            }
            set
            {
                if (name_ != value)
                {
                    SetPropertyValue((uint)PropTags.PR_DISPLAY_NAME, value);
                    name_ = value;
                }
            }
        }

        /// <summary>
        /// Gets IMAPIFolder Interface Guid
        /// </summary>
        public override Guid InterfaceID
        {
            get
            {
                return IID_IMAPIFolder;
            }
        }

        public string DefaultItemType
        {
            get
            {
                return defaultItemType_;
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Creates a new message.
        /// </summary>
        /// <returns>A MAPI Message object for the message created.</returns>
        public MAPIMessage CreateMessage()
        {
            IntPtr pMessage;
            try
            {
                HRESULT hResult = Folder.CreateMessage(IntPtr.Zero, 0, out pMessage);
                if (hResult == HRESULT.S_OK && pMessage != IntPtr.Zero)
                {
                    IMessage p = Marshal.GetObjectForIUnknown(pMessage) as IMessage;
                    if (p != null)
                    {
                        MAPIMessage message;
                        switch (DefaultItemType)
                        {
                            case MAPIObject.AppointmentItem:
                                message = new Appointment(p);
                                break;
                            case MAPIObject.ContactItem:
                                 message = new Contact(p);
                                 break;
                            default:
                                message = new MAPIMessage(p);
                                break;
                        }
                        return message;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex);
            }
            return null;
        }

        public bool CopyMessages(MAPIMessage message, MAPIFolder destFolder)
        {
            return CopyMessages(new List<MAPIMessage>(new MAPIMessage[] { message }), destFolder);

        }

        public bool CopyMessages(List<MAPIMessage> messages, MAPIFolder destFolder)
        {
            List<EntryID> msgEntryIds = new List<EntryID>();
            foreach (MAPIMessage message in messages)
            {
                msgEntryIds.Add(message.EntryID);
            }
            return CopyMessages(msgEntryIds, destFolder);

        }

        /// <summary>
        /// Deletes one message.
        /// </summary>
        /// <param name="message">MAPI message to delete</param>
        /// <returns>true, if the specified message was successfully deleted; otherwise, failed.</returns>
        public bool DeleteMessage(MAPIMessage message)
        {
            return DeleteMessages(new List<MAPIMessage>(new MAPIMessage[] { message }));
        }
        /// <summary>
        /// Delete messages
        /// </summary>
        /// <param name="messages">MAPI messages to delete</param>
        /// <returns>true, if the specified messages were successfully deleted; otherwise, failed.</returns>
        public bool DeleteMessages(List<MAPIMessage> messages)
        {
            List<EntryID> msgEntryIds = new List<EntryID>();
            foreach (MAPIMessage message in messages)
            {
                msgEntryIds.Add(message.EntryID);
            }
            return DeleteMessages(msgEntryIds);
      
        }

        public MAPIFolder OpenSubFolder(string name)
        {
            if (Hierarchy == null)
                return null;
            SRow[] sRows;
            EntryID entryID = null;
            while (Hierarchy.QueryRows(1, out sRows))
            {
                if (sRows.Length != 1)
                    break;
                if (sRows[0].propVals[0].AsString != null && sRows[0].propVals[0].AsString == name)
                {
                    if (sRows[0].propVals[1].AsBinary != null)
                    {
                        entryID = new EntryID(sRows[0].propVals[1].AsBinary);
                        break;
                    }
                }
            }
            if (entryID != null)
            {
                SBinary sb = SBinary.SBinaryCreate(entryID.AsByteArray);
                uint objectType;
                IntPtr pSubFolder;
                HRESULT hResult = Folder.OpenEntry(sb.cb, sb.lpb, IntPtr.Zero, (uint)OpenPropertyMode.MODIFY, out objectType, out pSubFolder);
                if (hResult == HRESULT.S_OK && pSubFolder != null && pSubFolder != IntPtr.Zero && objectType == (uint)ObjectType.MAPI_FOLDER)
                {
                    IMAPIFolder mapiObj = Marshal.GetObjectForIUnknown(pSubFolder) as IMAPIFolder;
                    if (mapiObj != null)
                        return new MAPIFolder(mapiObj);
                }
                SBinary.SBinaryRelease(ref sb);
            }
            return null;
        }

        #endregion

        #region Private Properties/Methods/Events

        MAPITable Hierarchy
        {
            get
            {
                if (hierarchy_ == null)
                {
                    IntPtr pTable;
                    HRESULT hResult = Folder.GetHierarchyTable(0, out pTable);
                    if (hResult == HRESULT.S_OK && pTable != null && pTable != IntPtr.Zero)
                    {
                        IMAPITable t = Marshal.GetObjectForIUnknown(pTable) as IMAPITable;
                        if (t != null)
                        {
                            hierarchy_ = new MAPITable(t);
                            hierarchy_.SetColumns(new PropTags[] { PropTags.PR_DISPLAY_NAME, PropTags.PR_ENTRYID });
                        }
                    }
                }
                return hierarchy_;
            }
        }

        IMAPIFolder Folder { get { return mapiObj_ as IMAPIFolder; } }

        bool DeleteMessages(List<EntryID> msgEntryIds)
        {

            int SBSize = Marshal.SizeOf(typeof(SBinary));
            IntPtr pMsgs = Marshal.AllocHGlobal(SBSize);
            IntPtr pArray = Marshal.AllocHGlobal(SBSize * msgEntryIds.Count);
            SBinary sb = new SBinary();
            sb.cb = (uint)msgEntryIds.Count;
            sb.lpb = pArray;
            Marshal.StructureToPtr(sb, pMsgs, true);
            List<SBinary> sBinaryList = new List<SBinary>();

            for (int i = 0; i < msgEntryIds.Count; i++ )
            {
                EntryID entry = msgEntryIds[i];
                sb = SBinary.SBinaryCreate(entry.AsByteArray);
                sBinaryList.Add(sb);
                Marshal.StructureToPtr(sb, pArray + i * SBSize, true);
            }
            
            HRESULT hr = HRESULT.E_FAIL;
            try
            {
                hr = Folder.DeleteMessages(pMsgs, 0, IntPtr.Zero, 0);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }

            for (int i = sBinaryList.Count - 1; i >= 0; i--)
            {
                SBinary s = sBinaryList[i];
                SBinary.SBinaryRelease(ref s);
            }
            sBinaryList.Clear();
            Marshal.FreeHGlobal(pArray);
            Marshal.FreeHGlobal(pMsgs);
            return hr == HRESULT.S_OK;
        }

        bool CopyMessages(List<EntryID> msgEntryIds, MAPIFolder destFolder)
        {

            int SBSize = Marshal.SizeOf(typeof(SBinary));
            IntPtr pMsgs = Marshal.AllocHGlobal(SBSize);
            IntPtr pArray = Marshal.AllocHGlobal(SBSize * msgEntryIds.Count);
            SBinary sb = new SBinary();
            sb.cb = (uint)msgEntryIds.Count;
            sb.lpb = pArray;
            Marshal.StructureToPtr(sb, pMsgs, true);
            List<SBinary> sBinaryList = new List<SBinary>();

            for (int i = 0; i < msgEntryIds.Count; i++)
            {
                EntryID entry = msgEntryIds[i];
                sb = SBinary.SBinaryCreate(entry.AsByteArray);
                sBinaryList.Add(sb);
                Marshal.StructureToPtr(sb, pArray + i * SBSize, true);
            }

            HRESULT hr = HRESULT.E_FAIL;
            try
            {
                hr = Folder.CopyMessages(pMsgs, IntPtr.Zero, (IntPtr)destFolder, 0, IntPtr.Zero, 0);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.Message);
            }

            for (int i = sBinaryList.Count - 1; i >= 0; i--)
            {
                SBinary s = sBinaryList[i];
                SBinary.SBinaryRelease(ref s);
            }
            sBinaryList.Clear();
            Marshal.FreeHGlobal(pArray);
            Marshal.FreeHGlobal(pMsgs);
            return hr == HRESULT.S_OK;
        }
        #endregion

        #region IDisposable Interface

        /// <summary>
        /// Dispose MAPI folder.
        /// </summary>
        public override void Dispose()
        {
            if (hierarchy_ != null)
            {
                hierarchy_.Dispose();
            }
            base.Dispose();
        }

        #endregion
    }
}


