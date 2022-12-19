using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// The IMAPIAdviseSink interface is used to implement an Advise Sink object for handling notifications.
    /// </summary>
    [
       ComImport, ComVisible(false),
       InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
       Guid("00020302-0000-0000-C000-000000000046")
    ]
    public interface IMAPIAdviseSink
    {
        /// <summary>
        /// The OnNotify method responds to a notification by performing one or more tasks, which depend on the object generating the notification, and type of event.
        /// </summary>
        /// <param name="cNotify">Ignored</param>
        /// <param name="lpNotifications">Reference to one NOTIFICATION structure that provides information about the events that have occurred.</param>
        /// <returns>S_OK, if the notification was processed successfully; otherwise, failed.</returns>
        HRESULT OnNotify(uint cNotify, IntPtr lpNotifications);
    }
    /// <summary>
    /// Describes information that relate to the arrival of a new message.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct NEWMAIL_NOTIFICATION
    {
        /// <summary>
        /// Count of bytes in the entry identifier pointed to by the lpEntryID member.
        /// </summary>
        public uint cbEntryID;
        /// <summary>
        /// Pointer to the entry identifier of the newly arrived message.
        /// </summary>
        public IntPtr pEntryID;
        /// <summary>
        /// Count of bytes in the entry identifier pointed to by the lpParentID member.
        /// </summary>
        public uint cbParentID;
        /// <summary>
        /// Pointer to the entry identifier of the receive folder for the newly arrived message.
        /// </summary>
        public IntPtr pParentID;
        /// <summary>
        /// Bitmask of flags used to describe the format of the string properties included with the message.
        /// </summary>
        public uint Flags;
        /// <summary>
        /// The message class of the newly arrived message.
        /// </summary>
        public IntPtr MessageClass;
        /// <summary>
        /// Bitmask of flags that describes the current state of the newly arrived message.
        /// </summary>
        public uint MessageFlags;
    }

    /// <summary>
    /// Contains information about an object that has undergone a change, such as being copied or modified.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct OBJECT_NOTIFICATION
    {
        /// <summary>
        /// Count of bytes in the entry identifier.
        /// </summary>
        public uint cbEntryID;
        /// <summary>
        /// Pointer to the entry identifier of the affected object.
        /// </summary>
        public IntPtr pEntryID;
        /// <summary>
        /// Type of object affected.
        /// </summary>
        public uint ObjType;
        /// <summary>
        /// Count of bytes in the entry identifier pointed to by the pParentID member.
        /// </summary>
        public uint cbParentID;
        /// <summary>
        /// Pointer to the entry identifier of the parent of the affected object.
        /// </summary>
        public IntPtr pParentID;
        /// <summary>
        /// Count of bytes in the entry identifier pointed to by the pOldID member.
        /// </summary>
        public uint cbOldID;
        /// <summary>
        /// Pointer to the entry identifier of the original object. This pointer can be NULL if the event does not require an original object.
        /// </summary>
        public IntPtr pOldID;
        /// <summary>
        /// Count of bytes in the entry identifier pointed to by the pOldParentID member.
        /// </summary>
        public uint cbOldParentID;
        /// <summary>
        /// Pointer to the entry identifier of the parent of the original object. This pointer can be NULL if the event does not require an original object.
        /// </summary>
        public IntPtr pOldParentID;
        /// <summary>
        /// Pointer to an SPropTagArray structure that contains the property tags identifying properties affected by the event.
        /// </summary>
        public IntPtr PropTagArray;
    }
    /// <summary>
    /// Describes information that relate to a critical error. This causes an error notification to be generated.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    struct ERROR_NOTIFICATION
    {
        /// <summary>
        /// Count of bytes in the entry identifier.
        /// </summary>
        uint cbEntryID;
        /// <summary>
        /// Pointer to the entry identifier of the object that causes the error.
        /// </summary>
        IntPtr lpEntryID;
        /// <summary>
        /// Error value for the critical error.
        /// </summary>
        uint scode;
        /// <summary>
        /// Bitmask of flags used to designate the format of the error text.
        /// </summary>
        uint ulFlags;
        /// <summary>
        /// Pointer to a MAPIERROR structure describing the error.
        /// </summary>
        IntPtr lpMAPIError;
    }
    /// <summary>
    /// Provides detailed information about an error, typically generated by the operating system, MAPI, or a service provider.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    struct MAPIERROR
    {
        /// <summary>
        /// Version number of the structure.
        /// </summary>
        uint ulVersion;
        /// <summary>
        /// Pointer to a string that describes the error. 
        /// </summary>
        IntPtr lpszError;
        /// <summary>
        /// Pointer to a string that describes the component that generated the error
        /// </summary>
        IntPtr lpszComponent;
        /// <summary>
        /// Low-level error value that is used only when the error to be returned is low-level.
        /// </summary>
        uint ulLowLevelError;
        /// <summary>
        /// 
        /// </summary>Value that represents the location in the component pointed to by the lpszComponent member that identifies where the error occurred.
        uint ulContext;
    }

    /// <summary>
    /// Describes a row in a table that has been affected by some type of event, such as a change or an error. This causes a table notification to be generated.
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    struct TABLE_NOTIFICATION
    {
        /// <summary>
        /// Bitmask of flags used to represent the table event type.
        /// </summary>
        uint ulTableEvent;
        /// <summary>
        /// HRESULT value for the error that has occurred, if the ulTableEvent member is set to TABLE_ERROR.
        /// </summary>
        HRESULT hResult;
        /// <summary>
        /// SPropValue structure for the PR_INSTANCE_KEY property of the affected row
        /// </summary>
        pSPropValue propIndex;
        /// <summary>
        /// SPropValue structure for the PR_INSTANCE_KEY property of the row before the affected one.
        /// </summary>
        pSPropValue propPrior;
        /// <summary>
        /// SRow structure describing the affected row.
        /// </summary>
        SRow row;
    }

    /// <summary>
    /// .Net Object implement IMAPIAdviseSink interface
    /// </summary>
    class MAPIAdviseSink : IMAPIAdviseSink
    {
        OnAdviseCallbackHandler callbackHandler_;
        IntPtr pContext_;

        /// <summary>
        /// Initializes a new instance of the MAPIAdviseSink class. 
        /// </summary>
        /// <param name="pContext">object pointer</param>
        /// <param name="callbackHandler">callback delegate</param>
        public MAPIAdviseSink(IntPtr pContext, OnAdviseCallbackHandler callbackHandler)
        {
            pContext_ = pContext;
            callbackHandler_ = callbackHandler;
        }
        /// <summary>
        /// Responds to a notification by performing one or more tasks.
        /// </summary>
        /// <param name="cNotify">The count of NOTIFICATION structures pointed to by the lpNotifications parameter.</param>
        /// <param name="lpNotifications">A pointer to one or more NOTIFICATION structures that provide information about the events that have occurred.</param>
        /// <returns>S_OK, if the notification was processed successfully; otherwise, failed.</returns>
        public HRESULT OnNotify(uint cNotify, IntPtr lpNotifications)
        {
            if (callbackHandler_ != null)
                callbackHandler_(pContext_, cNotify, lpNotifications);
            return HRESULT.S_OK;
        }
    }
}
