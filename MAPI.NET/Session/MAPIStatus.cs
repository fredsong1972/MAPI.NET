using System;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Provides status information about the MAPI subsystem, the integrated address book, and the MAPI spooler. 
    /// </summary>
    [
       ComImport, ComVisible(false),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
       Guid("00020305-0000-0000-C000-000000000046")
   ]
    public interface IMAPIStatus : IMAPIProp
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
        /// Confirms the external status information available for the MAPI resource or the service provider. 
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of any dialog boxes or windows that this method displays.</param>
        /// <param name="ulFlags"> A bitmask of flags that controls the validation.</param>
        /// <returns>S_OK, if the validation was successful; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT ValidateState(IntPtr ulUIParam, uint ulFlags);
        /// <exclude/>
        [PreserveSig]
        HRESULT SettingsDialog(IntPtr ulUIParam, uint ulFlags);
        /// <summary>
        /// Modifies a service provider's password without displaying a user interface.
        /// </summary>
        /// <param name="lpOldPass">The old password</param>
        /// <param name="lpNewPass">The new password</param>
        /// <param name="ulFlags">A bitmask of flags that controls the format of the passwords.</param>
        /// <returns>S_OK, if the password modification was successful; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT ChangePassword(string lpOldPass, string lpNewPass, uint ulFlags);
        /// <summary>
        /// Forces all messages waiting to be sent or received to be immediately uploaded or downloaded. 
        /// </summary>
        /// <param name="ulUIParam">A handle to the parent window of any dialog boxes or windows that this method displays.</param>
        /// <param name="cbTargetTransport">The byte count in the entry identifier pointed to by the lpTargetTransport parameter.</param>
        /// <param name="lpTargetTransport">A pointer to the entry identifier of the transport provider that is to flush its message queues. </param>
        /// <param name="ulFlags">A bitmask of flags that controls the flush operation.</param>
        /// <returns>S_OK, if the flush operation was successful; otherwise, failed.</returns>
        [PreserveSig]
        HRESULT FlushQueues(IntPtr ulUIParam, uint cbTargetTransport, IntPtr lpTargetTransport, uint ulFlags);
    }

    /// <summary>
    /// IMAPIStatus .Net wrapper object
    /// </summary>
    public class MAPIStatus : MAPIObject
    {
        Guid IID_IMAPIStatus = new Guid("00020305-0000-0000-C000-000000000046");

        enum TransportStatus : uint
        {
            INBOUND_ENABLED = 0x00010000,
            INBOUND_ACTIVE = 0x00020000,
            INBOUND_FLUSH = 0x00040000,
            OUTBOUND_ENABLED = 0x00100000,
            OUTBOUND_ACTIVE = 0x00200000,
            OUTBOUND_FLUSH = 0x00400000,
            REMOTE_ACCESS = 0x00800000,
        }

        const int timeOut_ = 300; //seconds

        /// <summary>
        /// Initializes a new instance of the MAPIStatus class. 
        /// </summary>
        /// <param name="status">IMAPIStatus object</param>
        /// <param name="id">Entry identification of IMAPIStatus object</param>
        public MAPIStatus(IMAPIStatus status, EntryID id)
            : base(status, id)
        {
        }

        #region Public Properties
        /// <summary>
        /// IMAPIStatus interface GUID
        /// </summary>
        public override Guid InterfaceID
        {
            get
            {
                return IID_IMAPIStatus;
            }
        }

        #endregion

        #region Public Methods
        /// <summary>
        /// Forces all messages waiting to be sent or received to be immediately uploaded or downloaded. 
        /// </summary>
        /// <param name="flushFlags">A bitmask of flags that controls the flush operation.</param>
        /// <returns>true, if the flush operation was successful; otherwise, failed.</returns>
        public bool FlushQueues(FlushFlag flushFlags)
        {
            HRESULT hResult = Status.FlushQueues(IntPtr.Zero, 0, IntPtr.Zero, (uint)flushFlags);
            return (hResult == HRESULT.S_OK);
        }
        /// <summary>
        /// Flush all outgoing messages
        /// </summary>
        public void SendAll()
        {
            FlushQueues(FlushFlag.UPLOAD | FlushFlag.FORCE);
        }
        /// <summary>
        /// Flush all incoming messages
        /// </summary>
        public void ReceiveAll()
        {
            FlushQueues(FlushFlag.DOWNLOAD | FlushFlag.FORCE);
        }
        /// <summary>
        /// Flush all outgoing and incoming messages
        /// </summary>
        public void SendReceiveAll()
        {
            FlushQueues(FlushFlag.DOWNLOAD | FlushFlag.UPLOAD | FlushFlag.FORCE);
        }

        #endregion

        #region Private Properties/Methods/Events

        IMAPIStatus Status { get { return mapiObj_ as IMAPIStatus; } }

        #endregion
    }
}
