using System;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Supports access to the MAPI address book and includes operations such as displaying common dialog boxes; opening containers, messaging users, and distribution lists; and performing name resolution.
    /// </summary>
    [
        ComImport, ComVisible(false),
        InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
        Guid("00020309-0000-0000-C000-000000000046")
    ]
    public interface IAddrBook : IMAPIProp
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
        /// Opens an address book entry and returns a pointer to an interface that can be used to access the entry.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID"> A pointer to the entry identifier that represents the address book entry to open.</param>
        /// <param name="lpInterface">A pointer to the interface identifier (IID) of the interface to be used to access the open entry.</param>
        /// <param name="ulFlags">A bitmask of flags that controls how the entry is opened. The following flags can be set.</param>
        /// <param name="lpulObjType">A pointer to the type of the opened entry.</param>
        /// <param name="lppUnk"> A pointer to a pointer to the opened entry.</param>
        /// <returns>S_OK, if the entry was successfully opened;otherwise, failed.</returns>
        HRESULT OpenEntry(uint cbEntryID, IntPtr lpEntryID, IntPtr lpInterface, uint ulFlags, out uint lpulObjType, out IntPtr lppUnk);
        /// <summary>
        /// Compares two entry identifiers that belong to a particular address book provider to determine whether they refer to the same address book object.
        /// </summary>
        /// <param name="cbEntryID1">The byte count in the entry identifier pointed to by the lpEntryID1 parameter.</param>
        /// <param name="lpEntryID1"> A pointer to the first entry identifier to be compared.</param>
        /// <param name="cbEntryID2">The byte count in the entry identifier pointed to by the lpEntryID2 parameter.</param>
        /// <param name="lpEntryID2">A pointer to the second entry identifier to be compared.</param>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpulResult">A pointer to the result of the comparison. The contents of lpulResult are set to TRUE if the two entry identifiers refer to the same object; otherwise, the contents are set to FALSE.</param>
        /// <returns>S_OK, if the call succeeded and has returned the expected value or values;otherwise, failed.</returns>
        HRESULT CompareEntryIDs(uint cbEntryID1, IntPtr lpEntryID1, uint cbEntryID2, IntPtr lpEntryID2, uint ulFlags, out uint lpulResult);
        /// <summary>
        /// Registers a client or service provider to receive notifications about changes to one or more entries in the address book.
        /// </summary>
        /// <param name="cbEntryID"> The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier of the address book container, messaging user, or distribution list that will generate a notification when a change occurs of the type or types described in the ulEventMask parameter.</param>
        /// <param name="ulEventMask">One or more notification events that the caller is registering to receive. Each event is associated with a particular notification structure that contains information about the change that occurred.</param>
        /// <param name="lpAdviseSink">A pointer to the advise sink object to be called when the event for which notification has been requested occurs.</param>
        /// <param name="lpulConnection"> A pointer to a nonzero connection number that represents the notification registration.</param>
        /// <returns>S_OK, if the notification registration was successful; otherwise, failed.</returns>
        HRESULT Advise(uint cbEntryID, IntPtr lpEntryID, uint ulEventMask, IntPtr lpAdviseSink, out uint lpulConnection);
        /// <summary>
        /// Cancels a notification registration previously established for an address book entry.
        /// </summary>
        /// <param name="ulConnection">A connection number that represents the registration to be canceled. The ulConnection parameter should contain a value returned by a prior call to the IAddrBook::Advise method.</param>
        /// <returns>S_OK, if the registration was successfully canceled; otherwise, failed.</returns>
        HRESULT Unadvise(uint ulConnection);
        /// <summary>
        /// Creates an entry identifier for a one-off address.
        /// </summary>
        /// <param name="lpszName">The recipient’s display name.</param>
        /// <param name="lpszAdrType">The address type of the recipient</param>
        /// <param name="lpszAddress"> The address of the recipient</param>
        /// <param name="ulFlags">A bitmask of flags that affects the one-off recipient.</param>
        /// <param name="lpcbEntryID">A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</param>
        /// <param name="lppEntryID">A pointer to a pointer to the entry identifier for the one-off recipient.</param>
        /// <returns>S_OK, if the one-off entry identifier was created successfully; otherwise, failed.</returns>
        HRESULT CreateOneOff([MarshalAs(UnmanagedType.LPWStr)] string lpszName, [MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType,
            [MarshalAs(UnmanagedType.LPWStr)] string lpszAddress, uint ulFlags, out uint lpcbEntryID, out IntPtr lppEntryID);
        /// <summary>
        /// Adds a new recipient to an address book container or to the recipient list of an outgoing message.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window for the dialog box.</param>
        /// <param name="ulFlags">A bitmask of flags that controls the type of the text that is used.</param>
        /// <param name="cbEIDContainer">The byte count in the entry identifier pointed to by the lpEIDContainer parameter.</param>
        /// <param name="lpEIDContainer">A pointer to the entry identifier of the container where the new recipient is to be added.</param>
        /// <param name="cbEIDNewEntryTpl">The byte count in the entry identifier pointed to by the lpEIDNewEntryTpl parameter.</param>
        /// <param name="lpEIDNewEntryTpl">A pointer to a one-off template that will be used to create the new recipient.</param>
        /// <param name="lpcbEIDNewEntry"> A pointer to the byte count in the entry identifier pointed to by the lppEIDNewEntry parameter.</param>
        /// <param name="lppEIDNewEntry"> A pointer to a pointer to the new recipient's entry identifier.</param>
        /// <returns>S_OK, if the new address book entry was successfully created; otherwise, failed.</returns>
        HRESULT NewEntry(uint ulUIParam, uint ulFlags, uint cbEIDContainer, IntPtr lpEIDContainer, uint cbEIDNewEntryTpl, IntPtr lpEIDNewEntryTpl, out uint lpcbEIDNewEntry, out IntPtr lppEIDNewEntry);
        /// <summary>
        /// Performs name resolution, assigning entry identifiers to recipients in a recipient list.
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of a dialog box that is shown, if specified, to prompt the user to resolve ambiguity.</param>
        /// <param name="ulFlags">A bitmask of flags that control various aspects of the resolution process. </param>
        /// <param name="lpszNewEntryTitle">A pointer to text for the title of the control in the dialog box that prompts the user to enter a recipient. </param>
        /// <param name="lpAdrList">A pointer to an ADRLIST structure that contains the list of recipient names to be resolved. </param>
        /// <returns>S_OK, if the name resolution process succeeded; otherwise, failed.</returns>
        HRESULT ResolveName(uint ulUIParam, uint ulFlags, [MarshalAs(UnmanagedType.LPWStr)] string lpszNewEntryTitle, IntPtr lpAdrList);
        /// <summary>
        /// Displays the Outlook address book dialog box.
        /// </summary>
        /// <param name="lpulUIParam"> A pointer to a handle of the parent window of the dialog box.</param>
        /// <param name="lpAdrParms"> A pointer to an ADRPARM structure that controls the presentation and behavior of the address dialog box.</param>
        /// <param name="lppAdrList"> pointer to a pointer to an ADRLIST structure that contains recipient information.</param>
        /// <returns>S_OK, if the common address dialog box was successfully displayed; otherwise, failed.</returns>
        HRESULT Address(out uint lpulUIParam, IntPtr lpAdrParms, out IntPtr lppAdrList);
        /// <summary>
        /// Displays a dialog box that shows details about a particular address book entry.
        /// </summary>
        /// <param name="lpulUIParam">A pointer to a handle of the parent window for the dialog box.</param>
        /// <param name="lpfnDismiss">A pointer to a function based on the DISMISSMODELESS prototype, or NULL.</param>
        /// <param name="lpvDismissContext">A pointer to context information to pass to the DISMISSMODELESS function pointed to by the lpfnDismiss parameter.</param>
        /// <param name="cbEntryID"> The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID">A pointer to the entry identifier for the entry for which details are displayed.</param>
        /// <param name="lpfButtonCallback">A pointer to a function based on the LPFNBUTTON function prototype</param>
        /// <param name="lpvButtonContext">A pointer to data that was used as a parameter for the function specified by the lpfButtonCallback parameter.</param>
        /// <param name="lpszButtonText">A pointer to a string that contains text to be applied to the added button, if that button is extensible. </param>
        /// <param name="ulFlags"> A bitmask of flags that controls the type of the text for the lpszButtonText parameter.</param>
        /// <returns>S_OK, if the details dialog box was successfully displayed for the address book entry; otherwise, failed.</returns>
        HRESULT Details(out uint lpulUIParam, IntPtr lpfnDismiss, IntPtr lpvDismissContext, uint cbEntryID, IntPtr lpEntryID,
            IntPtr lpfButtonCallback, IntPtr lpvButtonContext, [MarshalAs(UnmanagedType.LPWStr)] string lpszButtonText, uint ulFlags);
        /// <exclude/>
        HRESULT RecipOptions(uint ulUIParam, uint ulFlags, IntPtr lpRecip);
        /// <exclude/>
        HRESULT QueryDefaultRecipOpt([MarshalAs(UnmanagedType.LPWStr)] string lpszAdrType, uint ulFlags, out uint lpcValues, out IntPtr lppOptions);
        /// <summary>
        /// Returns the entry identifier of the container that is designated as the personal address book (PAB).
        /// </summary>
        /// <param name="lpcbEntryID">A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</param>
        /// <param name="lppEntryID">A pointer to a pointer to the entry identifier of the PAB.</param>
        /// <returns>S_OK, if the entry identifier of the PAB was successfully returned; otherwise, failed.</returns>
        HRESULT GetPAB(out uint lpcbEntryID, out IntPtr lppEntryID);
        /// <summary>
        /// Designates a particular container as the personal address book (PAB).
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID"> A pointer to the entry identifier of the container to be designated as the PAB. The lpEntryID parameter cannot be NULL.</param>
        /// <returns>S_OK, if the specified container has been established as the PAB; otherwise, failed.</returns>
        HRESULT SetPAB(uint cbEntryID, IntPtr lpEntryID);
        /// <summary>
        /// Returns the entry identifier for the initial address book container.
        /// </summary>
        /// <param name="lpcbEntryID">A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</param>
        /// <param name="lppEntryID">A pointer to a pointer to the entry identifier of the default container.</param>
        /// <returns>S_OK, if the entry identifier of the default container was successfully returned; otherwise, failed.</returns>
        HRESULT GetDefaultDir(out uint lpcbEntryID, out IntPtr lppEntryID);
        /// <summary>
        /// Establishes the specified container as the default address book container.
        /// </summary>
        /// <param name="cbEntryID">The byte count in the entry identifier pointed to by the lpEntryID parameter.</param>
        /// <param name="lpEntryID"> A pointer to the entry identifier of the default address book container.</param>
        /// <returns>S_OK, if the default address book container was successfully set; otherwise, failed.</returns>
        HRESULT SetDefaultDir(uint cbEntryID, IntPtr lpEntryID);
        /// <summary>
        /// Returns an ordered list of entry identifiers of the containers to be included in the name resolution process initiated by the IAddrBook::ResolveName method.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls the type of the strings returned in the search path. </param>
        /// <param name="lppSearchPath"> A pointer to a pointer to an ordered list of container entry identifiers.</param>
        /// <returns>S_OK, if the search path was successfully retrieved; otherwise, failed.</returns>
        HRESULT GetSearchPath(uint ulFlags, out IntPtr lppSearchPath);
        /// <summary>
        /// Sets a new search path in the profile that is used for the name resolution process.
        /// </summary>
        /// <param name="ulFlags">Reserved; must be zero.</param>
        /// <param name="lpSearchPath">A pointer to the SRowSet structure used to hold the search path. The first property for each aRow member in SRowSet must be PR_ENTRYID.</param>
        /// <returns>S_OK, if The search path was successfully set; otherwise, failed.</returns>
        HRESULT SetSearchPath(uint ulFlags, IntPtr lpSearchPath);
        /// <summary>
        /// Prepares a recipient list for later use by the messaging system.
        /// </summary>
        /// <param name="ulFlags">A bitmask of flags that controls how the entry is opened.</param>
        /// <param name="lpSPropTagArray">A pointer to an SPropTagArray structure that contains an array of property tags that indicate the properties, if any, that require updating. The lpSPropTagArray parameter can be NULL.</param>
        /// <param name="lpRecipList">A pointer to an ADRLIST structure that contains the list of recipients.</param>
        /// <returns>S_OK, if the recipient list was successfully prepared; otherwise, failed.</returns>
        HRESULT PrepareRecips(uint ulFlags, IntPtr lpSPropTagArray, IntPtr lpRecipList);
    }

    /// <summary>
    /// IAddrBook .Net Wrapper object
    /// </summary>
    public class MAPIAddressBook : IDisposable
    {
        internal const int SizeOfADREntry = 16;
        static Guid IID_IAddrBook = new Guid("00020309-0000-0000-C000-000000000046");

        /// <summary>
        /// Initializes a new instance of the MAPIAddressBook class. 
        /// </summary>
        /// <param name="addrBook">IAddrBook object</param>
        public MAPIAddressBook(IAddrBook addrBook) 
        {
            AddrBook = addrBook;
        }


        #region Public Methods

        /// <summary>
        /// Performs name resolution, assigning entry identifiers to recipients in a recipient list.
        /// </summary>
        /// <param name="pAdrList">A pointer to an ADRLIST structure that contains the list of recipient names to be resolved.</param>
        /// <returns>true if successful; otherwise, false</returns>
        public bool ResolveName(IntPtr pAdrList)
        {
            try
            {
                HRESULT hResult = AddrBook.ResolveName(0, (uint)CharacterSet.UNICODE, null, pAdrList);
                return hResult == HRESULT.S_OK;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine(ex.Message);
                return false;
            }
        }
        /// <summary>
        /// Creates an entry identifier for a one-off address.
        /// </summary>
        /// <param name="name">The recipient’s display name.</param>
        /// <param name="recipientType">The address type of the recipient.</param>
        /// <param name="emailAddress">The address of the recipient</param>
        /// <returns>true if successful; otherwise, false</returns>
        public EntryID CreateOneOff(string name, RecipientType recipientType, string emailAddress)
        {
            uint cb;
            IntPtr lpb;
             HRESULT hResult = AddrBook.CreateOneOff(name, recipientType.ToString(), emailAddress, (uint)CharacterSet.UNICODE, out cb, out lpb);
            if (hResult == HRESULT.S_OK)
            {
                return EntryID.BuildFromPtr(cb, lpb);
            }
            return null;
        }
        /// <summary>
        /// Get email address for a recipient
        /// </summary>
        /// <param name="entryId">Entry identifier of the recipient</param>
        /// <returns>Email address of the recipient</returns>
        public string GetEmailAddress(EntryID entryId)
        {
            SBinary sb = SBinary.SBinaryCreate(entryId.AsByteArray);
            IntPtr pObj;
            uint objType;
            HRESULT hResult = AddrBook.OpenEntry(sb.cb, sb.lpb, IntPtr.Zero, (uint)MAPIFlag.BEST_ACCESS, out objType, out pObj);
            if (hResult == HRESULT.S_OK && pObj != IntPtr.Zero)
            {
                if (objType == (uint)ObjectType.MAPI_DISTLIST || objType == (uint)ObjectType.MAPI_MAILUSER)
                {
                    MAPIObject mapiObj = null;
                    IMAPIProp prop = Marshal.GetObjectForIUnknown(pObj) as IMAPIProp;
                    if (prop != null)
                        mapiObj = new MAPIObject(prop, entryId);
                    if (mapiObj != null)
                    {
                        IPropValue value = mapiObj.GetProperty(PropTags.PR_SMTP_ADDRESS);
                        if (value == null || string.IsNullOrEmpty(value.AsString))
                        {
                            value = mapiObj.GetProperty(PropTags.PR_EMAIL_ADDRESS);
                        }
                        mapiObj.Dispose();
                        return value == null ? null : value.AsString;
                    }
                }
                Marshal.Release(pObj);
            }
            return string.Empty;
        }
        #endregion

        #region Private Properties/Methods/Events

        IAddrBook AddrBook { get; set; }

        #endregion

        #region IDisposable Interface
        /// <summary>
        /// Dispose MAPIAddressBook object
        /// </summary>
        public virtual void Dispose()
        {
            if (AddrBook != null)
            {
                Marshal.ReleaseComObject(AddrBook);
                AddrBook = null;
            }
        }

        #endregion

    }
}
