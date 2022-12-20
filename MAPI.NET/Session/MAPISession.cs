using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Manages objects associated with a MAPI logon session.
    /// </summary>
    [
      ComImport, ComVisible(false),
      InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
      Guid("00020300-0000-0000-C000-000000000046")
   ]

    public interface IMAPISession
    {
        /// <summary>
        /// Returns a MAPIERROR structure that contains information about the previous session error.
        /// </summary>
        /// <param name="hResult">A handle to the error value generated in the previous method call.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the type of strings returned.</param>
        /// <param name="lppMAPIError">A pointer to a pointer to a MAPIERROR structure that contains version, component, and context information for the error. </param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.</returns>
        HRESULT GetLastError(int hResult, uint ulFlags, out IntPtr lppMAPIError);
        /// <summary>
        /// Provides access to the message store table that contains information about all the message stores in the session profile.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that determines the format for columns that are character strings.</param>
        /// <param name="lppTable">A pointer to a pointer to the message store table.</param>
        /// <returns>S_OK, if the table was successfully returned; otherwise, failed.</returns>
        HRESULT GetMsgStoresTable(uint ulFlags, out IntPtr lppTable);
        /// <summary>
        /// Opens a message store and returns an IMsgStore pointer for further access.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of the common address dialog box and other related displays.</param>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the message store to be opened. The lpEntryID parameter must not be NULL.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the message store.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the object is opened.</param>
        /// <param name="lppMDB">Pointer to a pointer of the message store.</param>
        /// <returns>S_OK, if the message store was successfully opened; otherwise, failed.</returns>
        HRESULT OpenMsgStore(uint ulUIParam, uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out IntPtr lppMDB);
        /// <summary>
        /// Opens the MAPI integrated address book, returning an IAddrBook pointer for further access.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of the common address dialog box and other related displays.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the address book.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the opening of the address book.</param>
        /// <param name="lppAdrBook">A pointer to a pointer to the address book.</param>
        /// <returns>S_OK, if the address book was successfully opened; otherwise, failed.</returns>
        HRESULT OpenAddressBook(uint ulUIParam, IntPtr lpInterface, uint ulFlags, out IntPtr lppAdrBook);
        /// <summary>
        /// Opens a section of the current profile and returns an IProfSect pointer for further access.
        /// </summary>
        /// <param name="lpUID">pointer to the MAPIUID structure that identifies the profile section.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the profile section. </param>
        /// <param name="ulFlags">A bitmask of flags that controls access to the profile section. </param>
        /// <param name="lppProfSect">A pointer to a pointer to the profile section.</param>
        /// <returns>S_OK, if the profile section was successfully opened; otherwise, failed.</returns>
        HRESULT OpenProfileSection(ref Guid lpUID, ref Guid lpInterface, uint ulFlags, out IntPtr lppProfSect);
        /// <summary>
        /// Provides access to the status table, a table that contains information about all the MAPI resources in the session.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that determines the format for columns that are character strings.</param>
        /// <param name="lppTable">A pointer to a pointer to the status table.</param>
        /// <returns>S_OK, if the table was successfully returned; otherwise, failed.</returns>
        HRESULT GetStatusTable(uint ulFlags, out IntPtr lppTable);
        /// <summary>
        /// Opens an object and returns an interface pointer for additional access.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the object to open.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the opened object.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the object is opened. </param>
        /// <param name="lpulObjType">The type of the opened object</param>
        /// <param name="lppUnk">A pointer to a pointer to the opened object.</param>
        /// <returns>S_OK, if the object was opened successfully; otherwise, failed</returns>
        HRESULT OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
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
        HRESULT CompareEntryIDs(uint cbEntryID1, IntPtr lpEntryID1, uint cbEntryID2, IntPtr lpEntryID2, uint ulFlags, out bool lpulResult);
        /// <summary>
        /// Registers to receive notification of specified events that affect the session.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the address book or message store object about which notifications should be generated, or NULL, which indicates that the client is registering to receive notifications about events that affect only the session. </param>
        /// <param name="ulEventMask">A mask of values that indicate the types of notification events that the client is interested in and should be included in the registration.</param>
        /// <param name="lpAdviseSink">A pointer to an advise sink object to receive the subsequent notifications.</param>
        /// <param name="lpulConnection">A pointer to a nonzero number that represents the connection between the caller's advise sink object and the session.</param>
        /// <returns>S_OK, if the registration was successful; otherwise, failed.</returns>
        HRESULT Advise(uint cbEntryID, IntPtr lpEntryID, uint ulEventMask, IntPtr lpAdviseSink, out uint lpulConnection);
        /// <summary>
        /// Cancels the sending of notifications previously set up with a call to the IMAPISession::Advise method.
        /// </summary>
        /// <param name="ulConnection">A connection number associated with an active notification registration.</param>
        /// <returns>S_OK, if the registration was successfully canceled; otherwise, failed.</returns>
        HRESULT Unadvise(uint ulConnection);
        /// <exclude/>
        HRESULT MessageOptions(uint ulUIParam, uint ulFlags, [MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, IntPtr lpMessage);
        /// <exclude/>
        HRESULT QueryDefaultMessageOpt([MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, uint ulFlags, out uint lpcValues, out IntPtr lppOptions);
        /// <exclude/>
        HRESULT EnumAdrTypes(uint ulFlags, out uint lpcAdrTypes, out IntPtr lpppszAdrTypes);
        /// <summary>
        /// Returns the entry identifier of the object that provides the primary identity for the session.
        /// </summary>
        /// <param name="lpcbEntryID">A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</param>
        /// <param name="lppEntryID">A pointer to a pointer to the entry identifier of the object that provides the primary identity.</param>
        /// <returns>S_OK, if the primary identity was successfully returned; otherwise, failed.</returns>
        HRESULT QueryIdentity(out uint lpcbEntryID, out IntPtr lppEntryID);
        /// <summary>
        /// Ends a MAPI session.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of any dialog boxes or windows to be displayed. </param>
        /// <param name="ulFlags">A bitmask of flags that control the logoff operation. </param>
        /// <param name="ulReserved">Reserved; must be zero.</param>
        /// <returns>S_OK, if the logoff operation was successful; otherwise, failed.</returns>
        HRESULT Logoff(uint ulUIParam, uint ulFlags, uint ulReserved);
        /// <summary>
        /// Establishes a message store as the default message store for the session.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls the setting of the default message store.</param>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the message store that is intended as the default.</param>
        /// <returns>S_OK, if the call succeeded and returned the expected value or values; otherwise, failed.</returns>
        HRESULT SetDefaultStore(uint ulFlags, uint cbEntryID, IntPtr lpEntryID);
        /// <summary>
        /// Returns an IMsgServiceAdmin pointer for making changes to message services. 
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lppServiceAdmin">A pointer to a pointer to a message service administration object.</param>
        /// <returns>S_OK, if a pointer to a message service administration object was successfully returned; otherwise, failed.</returns>
        HRESULT AdminServices(uint ulFlags, out IntPtr lppServiceAdmin);
        /// <summary>
        /// Displays a form.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of the form.</param>
        /// <param name="lpMsgStore">A pointer to the message store that contains the folder pointed to by the lpParentFolder parameter. </param>
        /// <param name="lpParentFolder">A pointer to the folder in which the message associated with the ulMessageToken parameter was created. </param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the message that is displayed in the form.</param>
        /// <param name="ulMessageToken">The token that is associated with the message to be displayed in the form. </param>
        /// <param name="lpMessageSent">Reserved; must be NULL.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how and whether the message is saved.</param>
        /// <param name="ulMessageStatus">A bitmask of flags copied from the PR_MSG_STATUS property of the message associated with the token in the ulMessageToken parameter.</param>
        /// <param name="ulMessageFlags">A bitmask of flags copied from the PR_MESSAGE_FLAGS property of the message associated with the token in the ulMessageToken parameter. </param>
        /// <param name="ulAccess">A flag that indicates the permission level for the message that is displayed in the form. </param>
        /// <param name="lpszMessageClass">The message class of the message being displayed in the form.</param>
        /// <returns>S_OK, if the form was successfully displayed; otherwise, failed.</returns>
        HRESULT ShowForm(uint ulUIParam, IntPtr lpMsgStore, IntPtr lpParentFolder, ref Guid lpInterface, uint ulMessageToken,
        IntPtr lpMessageSent, uint ulFlags, uint ulMessageStatus, uint ulMessageFlags, uint ulAccess, [MarshalAs(UnmanagedType.LPWStr)] string lpszMessageClass);
        /// <summary>
        /// Creates a numeric token that the IMAPISession::ShowForm method uses to access a message.
        /// </summary>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) that represents the interface to be used to access the message.</param>
        /// <param name="lpMessage">A pointer to the message to be displayed in the form.</param>
        /// <param name="lpulMessageToken">A pointer to a message token, which is used by the IMAPISession::ShowForm method to access the message pointed to by lpMessage.</param>
        /// <returns>S_OK, if the form preparation was successful; otherwise, failed.</returns>
        HRESULT PrepareForm(ref Guid lpInterface, IntPtr lpMessage, out uint lpulMessageToken);
    }
    /// <summary>
    /// IMAPISession .Net wrapper object
    /// </summary>
    public partial class MAPISession : Component
    {
        static Guid IID_IMAPISession = new Guid("00020300-0000-0000-C000-000000000046");

        private IMAPISession session_ = null;
        private MAPIStatus mapiSpooler_ = null;
        private EntryID primaryIdentity_ = null;
        private string profileEmail_;

        /// <summary>
        /// Initializes a new instance of the MAPISession class. 
        /// </summary>
        public MAPISession()
        {
            InitializeComponent();
        }
        /// <summary>
        /// Initializes a new instance of the MAPISession class. 
        /// </summary>
        /// <param name="container">IContainer object</param>
        public MAPISession(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        #region Public Properties
        /// <summary>
        /// Gets current message store.
        /// </summary>
        public MsgStore MsgStore { get; private set; }

        /// <summary>
        /// Gets current spooler.
        /// </summary>
        public MAPIStatus Spooler
        {
            get
            {
                if (mapiSpooler_ == null)
                {
                    IntPtr pTable = IntPtr.Zero;
                    session_.GetStatusTable(0, out pTable);
                    if (pTable == IntPtr.Zero)
                        return null;
                    object tableObj = null;
                    MAPITable mb = null;
                    try
                    {
                        tableObj = Marshal.GetObjectForIUnknown(pTable);
                        mb = new MAPITable(tableObj as IMAPITable);
                        if (mb != null)
                        {
                            uint[] columns = new uint[] { 2, (uint)PropTags.PR_RESOURCE_TYPE, (uint)PropTags.PR_ENTRYID };
                            if (mb.SetColumns(new PropTags[] { PropTags.PR_RESOURCE_TYPE, PropTags.PR_ENTRYID }))
                            {
                                SRow[] sRows;

                                while (mb.QueryRows(1, out sRows))
                                {
                                    if (sRows.Length != 1)
                                        break;
                                    if (sRows[0].propVals[0].Tag == (uint)PropTags.PR_RESOURCE_TYPE && sRows[0].propVals[0].AsInt32 == (uint)MAPIFlag.SPOOLER)
                                    {
                                        if (sRows[0].propVals[1].Tag != (uint)PropTags.PR_ENTRYID)
                                            break;
                                        SBinary sb = SBinary.SBinaryCreate(sRows[0].propVals[1].AsBinary);
                                        IntPtr pStatus;
                                        uint objType;
                                        HRESULT hResult = session_.OpenEntry(sb.cb, sb.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pStatus);
                                        if (hResult == HRESULT.S_OK && pStatus != IntPtr.Zero)
                                        {
                                            IMAPIStatus s = Marshal.GetObjectForIUnknown(pStatus) as IMAPIStatus;
                                            if (s != null)
                                                mapiSpooler_ = new MAPIStatus(s, new EntryID(sb.AsBytes));
                                        }
                                        SBinary.SBinaryRelease(ref sb);
                                        break;
                                    }

                                }
                            }
                            mb.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Trace.WriteLine(ex.Message);
                    }

                }
                return mapiSpooler_;
            }
        }

        /// <summary>
        /// Get primary identity
        /// </summary>
        public EntryID PrimaryIdentity
        {
            get
            {
                if (primaryIdentity_ == null)
                {
                    uint cb;
                    IntPtr lpb;
                    HRESULT hr = session_.QueryIdentity(out cb, out lpb);
                    if (hr == HRESULT.S_OK)
                        primaryIdentity_ = EntryID.BuildFromPtr(cb, lpb);
                }
                return primaryIdentity_;
            }
        }

        /// <summary>
        /// Get profile email address
        /// </summary>
        public string ProfileEmailAddress
        {
            get
            {
                if (string.IsNullOrEmpty(profileEmail_))
                {
                    MAPIAddressBook addressBook = OpendAddressBook();
                    profileEmail_ = addressBook.GetEmailAddress(PrimaryIdentity);
                    addressBook.Dispose();
                }
                return profileEmail_;
            }

        }

        #endregion

        #region Public Methods
        /// <summary>
        /// MAPI session intialization and logon.
        /// </summary>
        /// <returns>true if successful; otherwise, false</returns>
        public bool Initialize()
        {
            IntPtr pSession = IntPtr.Zero;
            if (MAPINative.MAPIInitialize(IntPtr.Zero) == HRESULT.S_OK)
            {
                MAPINative.MAPILogonEx(0, null, null, (uint)(MAPIFlag.EXTENDED | MAPIFlag.USE_DEFAULT), out pSession);
                if (pSession == IntPtr.Zero)
                    MAPINative.MAPILogonEx(0, null, null, (uint)(MAPIFlag.EXTENDED | MAPIFlag.NEW_SESSION | MAPIFlag.USE_DEFAULT), out pSession);
            }

            if (pSession != IntPtr.Zero)
            {
                object sessionObj = null;
                try
                {
                    sessionObj = Marshal.GetObjectForIUnknown(pSession);
                    session_ = sessionObj as IMAPISession;
                }
                catch { }
            }

            return session_ != null;
        }

        /// <summary>
        /// Opens a message store.
        /// </summary>
        /// <param name="storeName">store name</param>
        /// <returns>true, if the message store was successfully opened; otherwise, failed.</returns>
        public bool OpenMessageStore(string storeName)
        {
            if (session_ == null)
                return false;
            IntPtr pTable = IntPtr.Zero;
            session_.GetMsgStoresTable(0, out pTable);
            if (pTable == IntPtr.Zero)
                return false;
            bool bResult = false;
            object tableObj = null;
            MAPITable mb = null;
            try
            {
                tableObj = Marshal.GetObjectForIUnknown(pTable);
                mb = new MAPITable(tableObj as IMAPITable);
                if (mb != null)
                {
                    if (mb.SetColumns(new PropTags[] { PropTags.PR_DISPLAY_NAME, PropTags.PR_ENTRYID, PropTags.PR_DEFAULT_STORE }))
                    {
                        SRow[] sRows;

                        while (mb.QueryRows(1, out sRows))
                        {
                            if (sRows.Length != 1)
                                break;
                            if (string.IsNullOrEmpty(storeName))
                            {
                                if (sRows[0].propVals[2].AsBool)
                                    bResult = true;
                            }
                            else if (sRows[0].propVals[0].AsString.IndexOf(storeName) > -1)
                                bResult = true;
                            if (bResult)
                                break;
                        }
                        if (bResult)
                        {
                            if (MsgStore != null)
                            {
                                MsgStore.Dispose();
                                MsgStore = null;
                            }
                            SBinary entryId = SBinary.SBinaryCreate(sRows[0].propVals[1].AsBinary);
                            IntPtr pStore = IntPtr.Zero;
                            if (session_.OpenMsgStore(0, entryId.cb, entryId.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out pStore) == HRESULT.S_OK)
                            {
                                if (pStore != IntPtr.Zero)
                                {
                                    //Guid storeID = new Guid("00020306-0000-0000-C000-000000000046");
                                    //IntPtr pStore2 = IntPtr.Zero;
                                    //int hr = Marshal.QueryInterface(pStore, ref storeID, out pStore2);
                                    IMsgStore msgStore = Marshal.GetObjectForIUnknown(pStore) as IMsgStore;
                                    MsgStore = new MsgStore(this, msgStore, new EntryID(entryId.AsBytes));
                                }
                            }
                            SBinary.SBinaryRelease(ref entryId);
                        }
                    }
                    mb.Dispose();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }
            return bResult;
        }
        /// <summary>
        /// Opens the MAPI integrated address book.
        /// </summary>
        /// <returns>The MAPI address book object which was opened</returns>
        public MAPIAddressBook OpendAddressBook()
        {
            IntPtr p;
            HRESULT hResult = session_.OpenAddressBook(0, IntPtr.Zero, (uint)MAPIFlag.AB_NO_DIALOG, out p);
            if (hResult == HRESULT.S_OK && p != IntPtr.Zero)
            {
                IAddrBook addrBook = Marshal.GetObjectForIUnknown(p) as IAddrBook;
                return new MAPIAddressBook(addrBook);
            }
            return null;
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
                HRESULT hResult = session_.OpenEntry(sb.cb, sb.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objectType, out pObject);
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

        /// <summary>
        /// Flush all outgoing messages. 
        /// </summary>
        /// <returns>true, if successful; otherwise, failed.</returns>
        public bool SendAll()
        {
            bool bResult = false;
            MAPIStatus status = Spooler;
            if (status != null)
            {
                status.SendAll();
            }

            return bResult;
        }

        /// <exclude/>
        public static IPropValue GetObjectProperty(IMAPIProp obj, uint tag)
        {
            Dictionary<uint, IPropValue> props = GetObjectProperties(obj, new uint[] { tag });
            return props.ContainsKey(tag) ? props[tag] : null;
        }

        /// <exclude/>
        public static Dictionary<uint, IPropValue> GetObjectProperties(IMAPIProp obj, uint[] tags)
        {
            Dictionary<uint, IPropValue> props = new Dictionary<uint, IPropValue>();
            IntPtr propVals = IntPtr.Zero;
            uint ivalues = 0;
            uint iprops = (uint)tags.Length;
            uint[] p = new uint[tags.Length + 1];

            p[0] = (uint)tags.Length;
            for (int i = 0; i < tags.Length; i++)
                p[i + 1] = tags[i];

            HRESULT hResult = HRESULT.E_FAIL;
            try
            {
                hResult = obj.GetProps(p, 0, out ivalues, ref propVals);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
            }

            if (propVals != IntPtr.Zero)
            {

                uint pProps = (uint)propVals;
                for (int i = 0; i < ivalues; i++)
                {
                    pSPropValue lpProp = (pSPropValue)Marshal.PtrToStructure((IntPtr)(pProps + i * Marshal.SizeOf(typeof(pSPropValue))), typeof(pSPropValue));
                    if (lpProp.ulPropTag == (uint)PT.PT_ERROR)
                        continue;
                    if ((lpProp.ulPropTag == tags[i]) || ((lpProp.ulPropTag & 0xFFFF0000) == tags[i]))
                    {
                        MAPIProp prop = new MAPIProp(lpProp);
                        props[tags[i]] = prop;
                    }
                }
                MAPINative.MAPIFreeBuffer(propVals);
            }
            return props;
        }

        #endregion

   }
}
