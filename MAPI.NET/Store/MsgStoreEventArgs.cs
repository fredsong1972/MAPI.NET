using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace MAPI.NET
{
    /// <summary>
    /// Provides data for an object that has undergone a change, such as being copied or modified.
    /// </summary>
    public class MsgStoreObjectEventArgs : EventArgs
    {
        /// <summary>
        /// Get entry identification of message store.
        /// </summary>
        public EntryID StoreID { get; private set; }
        /// <summary>
        /// Get entry identification of the object.
        /// </summary>
        public EntryID EntryID { get; private set; }
        /// <summary>
        /// Get entry identification of the parent object.
        /// </summary>
        public EntryID ParentID { get; private set; }
        /// <summary>
        /// Get entry identification of the original object.
        /// </summary>
        public EntryID OldID { get; private set; }
        /// <summary>
        /// Get entry identification of the original parent object.
        /// </summary>
        public EntryID OldParentID { get; private set; }
        /// <summary>
        /// Type of object affected.
        /// </summary>
        public ObjectType ObjectType { get; private set; }

        /// <summary>
        /// Initializes a new instance of the MsgStoreObjectEventArgs class. 
        /// </summary>
        /// <param name="storeID">entry identification of message store</param>
        /// <param name="notification">Object notification structure</param>
        public MsgStoreObjectEventArgs(EntryID storeID, OBJECT_NOTIFICATION notification)
        {
            StoreID = storeID;
            SBinary sbEntry = new SBinary() { cb = notification.cbEntryID, lpb = notification.pEntryID };
            SBinary sbParent = new SBinary() { cb = notification.cbParentID, lpb = notification.pParentID };
            SBinary sbOld = new SBinary() { cb = notification.cbOldID, lpb = notification.pOldID };
            SBinary sbOldParent = new SBinary() { cb = notification.cbOldParentID, lpb = notification.pOldParentID };
            EntryID = sbEntry.cb > 0 ? new EntryID(sbEntry.AsBytes) : null;
            ParentID = sbParent.cb > 0 ? new EntryID(sbParent.AsBytes) : null;
            OldID = sbOld.cb > 0 ? new EntryID(sbOld.AsBytes) : null;
            OldParentID = sbOldParent.cb > 0 ? new EntryID(sbOldParent.AsBytes) : null;
            ObjectType = (ObjectType)notification.ObjType;
        }
    }
    /// <summary>
    /// Provides data for the arrival of a new message.
    /// </summary>
    public class MsgStoreNewMailEventArgs : EventArgs
    {
        /// <summary>
        /// Entry identification of message store.
        /// </summary>
        public EntryID StoreID { get; private set; }
        /// <summary>
        /// Entry identification of the newly arrived message.
        /// </summary>
        public EntryID EntryID { get; private set; }
        /// <summary>
        /// The entry identifier of the receive folder for the newly arrived messag.
        /// </summary>
        public EntryID ParentID { get; private set; }
        /// <summary>
        /// Bitmask of flags that describes the current state of the newly arrived message.
        /// </summary>
        public int MessageFlags { get; private set; }
        /// <summary>
        /// The message class of the newly arrived message.
        /// </summary>
        public string MessageClass { get; private set; }
        /// <summary>
        /// Initializes a new instance of the MsgStoreNewMailEventArgs class. 
        /// </summary>
        /// <param name="storeID">The entry identification of message store.</param>
        /// <param name="notification">The new mail notification structure.</param>
        public MsgStoreNewMailEventArgs(EntryID storeID, NEWMAIL_NOTIFICATION notification)
        {
            StoreID = storeID;
            SBinary sbEntry = new SBinary() { cb = notification.cbEntryID, lpb = notification.pEntryID };
            SBinary sbParent = new SBinary() { cb = notification.cbParentID, lpb = notification.pParentID };
            EntryID = sbEntry.cb > 0 ? new EntryID(sbEntry.AsBytes) : null;
            ParentID = sbParent.cb > 0 ? new EntryID(sbParent.AsBytes) : null;
            MessageFlags = (int)notification.MessageFlags;
            if ((notification.Flags & (uint)CharacterSet.UNICODE) == (uint)CharacterSet.UNICODE)
                MessageClass = Marshal.PtrToStringUni(notification.MessageClass);
            else
                MessageClass = Marshal.PtrToStringAnsi(notification.MessageClass);
        }
    }
}
