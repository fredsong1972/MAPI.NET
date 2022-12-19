using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MAPI.NET
{
    /// <summary>
    /// MAPI Property Tag
    /// </summary>
    public enum PropTags : uint
    {
        /// <summary>
        /// Error
        /// </summary>
        PR_ERROR = 10,

        #region Common non-transmittable
        /// <summary>
        ///  Contains a MAPI entry identifier used to open and edit properties of a particular MAPI object.
        ///  <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ENTRYID = PT.PT_BINARY | 0x0FFF << 16,
        /// <summary>
        /// Provides file type information for a non-Windows attachment.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_OBJECT_TYPE = PT.PT_LONG | 0x0FFE << 16,
        /// <summary>
        /// Contains a bitmap of a full size icon for a form.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ICON = PT.PT_BINARY | 0x0FFD << 16,
        /// <summary>
        /// Contains a bitmap of a half-size icon for a form.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MINI_ICON = PT.PT_BINARY | 0x0FFC << 16,
        /// <summary>
        /// Contains the unique entry identifier of the message store in which an object resides.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_STORE_ENTRYID = PT.PT_BINARY | 0x0FFB << 16,
        /// <summary>
        /// Contains the unique binary-comparable identifier (record key) of the message store in which an object resides.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_STORE_RECORD_KEY = PT.PT_BINARY | 0x0FFA << 16,
        /// <summary>
        /// Contains a unique binary-comparable identifier for a specific object.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RECORD_KEY = PT.PT_BINARY | 0x0FF9 << 16,
        /// <summary>
        /// Contains the mapping signature for named properties of a particular MAPI object.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MAPPING_SIGNATURE = PT.PT_BINARY | 0x0FF8 << 16,
        /// <summary>
        /// Contains a bitmask of flags indicating the level at which a client application can access the open object.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ACCESS_LEVEL = PT.PT_LONG | 0x0FF7 << 16,
        /// <summary>
        /// Contains a value that uniquely identifies a row in a table.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_INSTANCE_KEY = PT.PT_BINARY | 0x0FF6 << 16,
        /// <summary>
        /// Contains a value that indicates the type of a row in a table.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ROW_TYPE = PT.PT_LONG | 0x0FF5 << 16,
        /// <summary>
        /// Contains a bitmask of flags indicating the operations a client application can perform on the open object.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ACCESS = PT.PT_LONG | 0x0FF4 << 16,

        #endregion

        #region Common transmittable
        /// <summary>
        /// Contains a unique identifier for a recipient in a recipient table or status table.
        ///  <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ROWID = PT.PT_LONG | 0x3000 << 16,
        /// <summary>
        /// Contains the display name of the message store.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DISPLAY_NAME = PT.PT_TSTRING | 0x3001 << 16,
        /// <summary>
        /// Contains the display name of the message store. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_DISPLAY_NAME_W = PT.PT_UNICODE | 0x3001 << 16,
        /// <summary>
        /// Contains the display name of the message store. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_DISPLAY_NAME_A = PT.PT_STRING8 | 0x3001 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP).
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ADDRTYPE = PT.PT_TSTRING | 0x3002 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP). UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ADDRTYPE_W = PT.PT_UNICODE | 0x3002 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address type, such as Simple Mail Transfer Protocol (SMTP). Non-UNICODE compilation.
        /// <br></br>Data Type:PT.PT_STRING8
        /// </summary>
        PR_ADDRTYPE_A = PT.PT_STRING8 | 0x3002 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_EMAIL_ADDRESS = PT.PT_TSTRING | 0x3003 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x3003 << 16,
        /// <summary>
        /// Contains the messaging user's e-mail address. Non-UNICODE compilation.
        /// <br></br>Data Type: PT.PT_STRING8
        /// </summary>
        PR_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x3003 << 16,
        /// <summary>
        ///  Contains a comment about the purpose or content of an object.
        ///  <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_COMMENT = PT.PT_TSTRING | 0x3004 << 16,
        /// <summary>
        /// Contains a comment about the purpose or content of an object. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_COMMENT_W = PT.PT_UNICODE | 0x3004 << 16,
        /// <summary>
        /// Contains a comment about the purpose or content of an object. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_COMMENT_A = PT.PT_STRING8 | 0x3004 << 16,
        /// <summary>
        /// Represents the relative level of indentation, or depth, of an object in a hierarchy table.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_DEPTH = PT.PT_LONG | 0x3005 << 16,
        /// <summary>
        /// Contains the vendor-defined display name for a service provider.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PROVIDER_DISPLAY = PT.PT_TSTRING | 0x3006 << 16,
        /// <summary>
        /// Contains the vendor-defined display name for a service provider. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_PROVIDER_DISPLAY_W = PT.PT_UNICODE | 0x3006 << 16,
        /// <summary>
        /// Contains the vendor-defined display name for a service provider. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_PROVIDER_DISPLAY_A = PT.PT_STRING8 | 0x3006 << 16,
        /// <summary>
        /// Contains the creation date and time for a message.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_CREATION_TIME = PT.PT_SYSTIME | 0x3007 << 16,
        /// <summary>
        /// Contains the date and time the object or subobject was last modified.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_LAST_MODIFICATION_TIME = PT.PT_SYSTIME | 0x3008 << 16,
        /// <summary>
        /// Contains a bitmask of flags for message services and providers.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RESOURCE_FLAGS = PT.PT_LONG | 0x3009 << 16,
        /// <summary>
        /// Contains the base filename of the MAPI service provider DLL.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PROVIDER_DLL_NAME = PT.PT_TSTRING | 0x300A << 16,
        /// <summary>
        /// Contains the base filename of the MAPI service provider DLL. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_PROVIDER_DLL_NAME_W = PT.PT_UNICODE | 0x300A << 16,
        /// <summary>
        /// Contains the base filename of the MAPI service provider DLL. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_PROVIDER_DLL_NAME_A = PT.PT_STRING8 | 0x300A << 16,
        /// <summary>
        /// Contains a binary-comparable key that identifies correlated objects for a search.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SEARCH_KEY = PT.PT_BINARY | 0x300B << 16,
        /// <summary>
        /// Contains a MAPIUID structure of the service provider that is handling a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_PROVIDER_UID = PT.PT_BINARY | 0x300C << 16,
        /// <summary>
        /// Contains the zero-based index of a service provider's position in the provider table.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_PROVIDER_ORDINAL = PT.PT_LONG | 0x300D << 16,

        #endregion

        #region Message Store Properties

        /// <summary>
        /// Contains TRUE if a message store is the default message store in the message store table.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_DEFAULT_STORE = PT.PT_BOOLEAN | 0x3400 << 16,
        /// <summary>
        /// Contains a bitmask of flags that client applications should query to determine the characteristics of a message store.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_STORE_SUPPORT_MASK = PT.PT_LONG | 0x340D << 16,
        /// <summary>
        /// Contains a flag that describes the state of the message store.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_STORE_STATE = PT.PT_LONG | 0x340E << 16,
        /// <summary>
        /// Was originally meant to contain the search key of the interpersonal message (IPM) root folder. No longer supported
        ///<br></br> Data Type: PT_BINARY
        /// </summary>
        PR_IPM_SUBTREE_SEARCH_KEY = PT.PT_BINARY | 0x3410 << 16,
        /// <summary>
        /// Was originally meant to contain the search key of the standard Outbox folder. No longer supported.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_OUTBOX_SEARCH_KEY = PT.PT_BINARY | 0x3411 << 16,
        /// <summary>
        /// Was originally meant to contain the search key of the standard Deleted Items folder. No longer supported.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_WASTEBASKET_SEARCH_KEY = PT.PT_BINARY | 0x3412 << 16,
        /// <summary>
        /// Was originally meant to contain the search key of the standard Sent Items folder. No longer supported.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_SENTMAIL_SEARCH_KEY = PT.PT_BINARY | 0x3413 << 16,
        /// <summary>
        /// Contains a provider-defined MAPIUID structure that indicates the type of the message store.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MDB_PROVIDER = PT.PT_BINARY | 0x3414 << 16,
        /// <summary>
        /// Contains a table of a message store's receive folder settings.
        /// <br></br>Data Type: PT_OBJECT
        /// </summary>
        PR_RECEIVE_FOLDER_SETTINGS = PT.PT_OBJECT | 0x3415 << 16,
        /// <summary>
        /// Contains a bitmask of flags that indicate the validity of the entry identifiers of the folders in a message store.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_VALID_FOLDER_MASK = PT.PT_LONG | 0x35DF << 16,
        /// <summary>
        /// Contains the entry identifier of the root of the IPM folder subtree in the message store's folder tree.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_SUBTREE_ENTRYID = PT.PT_BINARY | 0x35E0 << 16,
        /// <summary>
        /// Contains the entry identifier of the standard interpersonal message (IPM) Outbox folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_OUTBOX_ENTRYID = PT.PT_BINARY | 0x35E2 << 16,
        /// <summary>
        /// Contains the entry identifier of the standard IPM Deleted Items folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_WASTEBASKET_ENTRYID = PT.PT_BINARY | 0x35E3 << 16,
        /// <summary>
        /// Contains the entry identifier of the standard IPM Sent Items folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_SENTMAIL_ENTRYID = PT.PT_BINARY | 0x35E4 << 16,
        /// <summary>
        /// Contains the entry identifier of the user-defined Views folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_VIEWS_ENTRYID = PT.PT_BINARY | 0x35E5 << 16,
        /// <summary>
        /// Contains the entry identifier of the predefined common view folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_COMMON_VIEWS_ENTRYID = PT.PT_BINARY | 0x35E6 << 16,
        /// <summary>
        /// Contains the entry identifier for the folder in which search results are typically created.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_FINDER_ENTRYID = PT.PT_BINARY | 0x35E7 << 16,

        #endregion

        #region  Message non-transmittable properties
        /// <summary>
        /// Was originally meant to contain the current version of a message store. No longer supported.
        /// <br></br>Data Type: PT_I8
        /// </summary>
        PR_CURRENT_VERSION = PT.PT_I8 | 0x0E00 << 16,
        /// <summary>
        /// Contains TRUE if a client application wants MAPI to delete the associated message after submission.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_DELETE_AFTER_SUBMIT = PT.PT_BOOLEAN | 0x0E01 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;).
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DISPLAY_BCC = PT.PT_TSTRING | 0x0E02 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_DISPLAY_BCC_W = PT.PT_UNICODE | 0x0E02 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any blind carbon copy (BCC) message recipients, separated by semicolons (;). Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_DISPLAY_BCC_A = PT.PT_STRING8 | 0x0E02 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;).
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DISPLAY_CC = PT.PT_TSTRING | 0x0E03 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_DISPLAY_CC_W = PT.PT_UNICODE | 0x0E03 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of any carbon copy (CC) message recipients, separated by semicolons (;). Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_DISPLAY_CC_A = PT.PT_STRING8 | 0x0E03 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;).
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DISPLAY_TO = PT.PT_TSTRING | 0x0E04 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;). UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_DISPLAY_TO_W = PT.PT_UNICODE | 0x0E04 << 16,
        /// <summary>
        /// Contains an ASCII list of the display names of the primary (To) message recipients, separated by semicolons (;). Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_DISPLAY_TO_A = PT.PT_STRING8 | 0x0E04 << 16,
        /// <summary>
        /// Contains the display name of the folder in which a message was found during a search.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PARENT_DISPLAY = PT.PT_TSTRING | 0x0E05 << 16,
        /// <summary>
        /// Contains the display name of the folder in which a message was found during a search. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_PARENT_DISPLAY_W = PT.PT_UNICODE | 0x0E05 << 16,
        /// <summary>
        /// Contains the display name of the folder in which a message was found during a search. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_PARENT_DISPLAY_A = PT.PT_STRING8 | 0x0E05 << 16,
        /// <summary>
        /// Contains the date and time a message was delivered.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_MESSAGE_DELIVERY_TIME = PT.PT_SYSTIME | 0x0E06 << 16,
        /// <summary>
        /// Contains a bitmask of flags indicating the origin and current state of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_MESSAGE_FLAGS = PT.PT_LONG | 0x0E07 << 16,
        /// <summary>
        /// Contains the sum, in bytes, of the sizes of all properties on a message object.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_MESSAGE_SIZE = PT.PT_LONG | 0x0E08 << 16,
        /// <summary>
        /// Contains the entry identifier of the folder containing a folder or message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_PARENT_ENTRYID = PT.PT_BINARY | 0x0E09 << 16,
        /// <summary>
        /// Contains the entry identifier of the folder where the message should be moved after submission.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SENTMAIL_ENTRYID = PT.PT_BINARY | 0x0E0A << 16,
        /// <summary>
        /// Contains TRUE if the sender of a message requests the correlation feature of the messaging system.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_CORRELATE = PT.PT_BOOLEAN | 0x0E0C << 16,
        /// <summary>
        /// Contains the message transfer system (MTS) identifier used in correlating reports with sent messages.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CORRELATE_MTSID = PT.PT_BINARY | 0x0E0D << 16,
        /// <summary>
        /// Contains TRUE if a nondelivery report applies only to discrete members of a distribution list rather than the entire list.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_DISCRETE_VALUES = PT.PT_BOOLEAN | 0x0E0E << 16,
        /// <summary>
        /// Contains TRUE if some transport provider has already accepted responsibility for delivering the message to this recipient, and FALSE if the MAPI spooler considers that this transport provider should accept responsibility.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_RESPONSIBILITY = PT.PT_BOOLEAN | 0x0E0F << 16,
        /// <summary>
        /// Contains the status of the message based on information available to the MAPI spooler.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_SPOOLER_STATUS = PT.PT_LONG | 0x0E10 << 16,
        /// <summary>
        /// Contains a table of restrictions that can be applied to a contents table to find all messages that contain recipient subobjects meeting the restrictions.
        ///  <br></br>Data Type: PT_OBJECT
        /// </summary>
        PR_MESSAGE_RECIPIENTS = PT.PT_OBJECT | 0x0E12 << 16,
        /// <summary>
        /// Contains a table of restrictions that can be applied to a contents table to find all messages that contain attachment subobjects meeting the restrictions.
        /// <br></br>Data Type: PT_OBJECT
        /// </summary>
        PR_MESSAGE_ATTACHMENTS = PT.PT_OBJECT | 0x0E13 << 16,
        /// <summary>
        /// Contains a bitmask of flags indicating details about a message submission.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_SUBMIT_FLAGS = PT.PT_LONG | 0x0E14 << 16,
        /// <summary>
        /// Contains a value used by the MAPI spooler in assigning delivery responsibility among transport providers.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RECIPIENT_STATUS = PT.PT_LONG | 0x0E15 << 16,
        /// <summary>
        /// Contains a value used by the MAPI spooler to track the progress of an outbound message through the outgoing transport providers.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_TRANSPORT_KEY = PT.PT_LONG | 0x0E16 << 16,
        /// <summary>
        /// Contains a bitmask of property tags that define the status of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_MSG_STATUS = PT.PT_LONG | 0x0E17 << 16,
        /// <summary>
        /// Contains the estimated time to download a message from a remote server to a local message store.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_MESSAGE_DOWNLOAD_TIME = PT.PT_LONG | 0x0E18 << 16,
        /// <summary>
        /// Contains at least one attachment.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_HASATTACH = PT.PT_BOOLEAN | 0x0E1B << 16,
        /// <summary>
        /// Contains a circular redundancy check (CRC) value on the message text.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_BODY_CRC = PT.PT_LONG | 0x0E1C << 16,
        /// <summary>
        /// ontains the message subject with any prefix removed.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_NORMALIZED_SUBJECT = PT.PT_TSTRING | 0x0E1D << 16,
        /// <summary>
        /// Contains the message subject with any prefix removed. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_NORMALIZED_SUBJECT_W = PT.PT_UNICODE | 0x0E1D << 16,
        /// <summary>
        /// Contains the message subject with any prefix removed. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_NORMALIZED_SUBJECT_A = PT.PT_STRING8 | 0x0E1D << 16,
        /// <summary>
        /// Contains TRUE if PR_RTF_COMPRESSED has the same text content as PR_BODY for this message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_RTF_IN_SYNC = PT.PT_BOOLEAN | 0x0E1F << 16,
        /// <summary>
        /// Contains the sum, in bytes, of the sizes of all properties on an attachment.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ATTACH_SIZE = PT.PT_LONG | 0x0E20 << 16,
        /// <summary>
        /// Contains a number that uniquely identifies the attachment within its parent message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ATTACH_NUM = PT.PT_LONG | 0x0E21 << 16,
        /// <summary>
        /// Contains an ASN.1 proof of submission value.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PREPROCESS = PT.PT_BOOLEAN | 0x0E22 << 16,

        #endregion
        #region Message recipient properties
        
        /// <summary>
        /// Contains an ASN.1 content integrity check value that allows a message sender to help protect message content from disclosure to unauthorized recipients.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONTENT_INTEGRITY_CHECK = PT.PT_BINARY | 0x0C00 << 16,
        /// <summary>
        /// Contains a value that indicates that a message sender has requested a message content conversion for a particular recipient.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_EXPLICIT_CONVERSION = PT.PT_LONG | 0x0C01 << 16,
        /// <summary>
        /// Contains TRUE if this message should be returned with a report.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_IPM_RETURN_REQUESTED = PT.PT_BOOLEAN | 0x0C02 << 16,
        /// <summary>
        /// Contains an ASN.1 security token for a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MESSAGE_TOKEN = PT.PT_BINARY | 0x0C03 << 16,
        /// <summary>
        /// Contains an encoded reason for nondelivery that forms part of a nondelivery report.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_NDR_REASON_CODE = PT.PT_LONG | 0x0C04 << 16,
        /// <summary>
        /// Contains a diagnostic code that forms part of a nondelivery report.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_NDR_DIAG_CODE = PT.PT_LONG | 0x0C05 << 16,
        /// <summary>
        /// Contains TRUE if a message sender wants notification of non-receipt for a specified recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_NON_RECEIPT_NOTIFICATION_REQUESTED = PT.PT_BOOLEAN | 0x0C06 << 16,
        /// <summary>
        /// Contains a value that specifies the nature of the functional entity by means of which a message was or would have been delivered to the recipient.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_DELIVERY_POINT = PT.PT_LONG | 0x0C07 << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests a nondelivery report for a particular recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_ORIGINATOR_NON_DELIVERY_REPORT_REQUESTED = PT.PT_BOOLEAN | 0x0C08 << 16,
        /// <summary>
        /// Contains an entry identifier for an alternate recipient designated by the sender.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINATOR_REQUESTED_ALTERNATE_RECIPIENT = PT.PT_BINARY | 0x0C09 << 16,
        /// <summary>
        /// Contains TRUE if the messaging system should use a fax bureau for physical delivery of this message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PHYSICAL_DELIVERY_BUREAU_FAX_DELIVERY = PT.PT_BOOLEAN | 0x0C0A << 16,
        /// <summary>
        /// Contains a bitmask of flags defining the physical delivery mode (for example, special delivery) for a message designated for a specific recipient.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_PHYSICAL_DELIVERY_MODE = PT.PT_LONG | 0x0C0B << 16,
        /// <summary>
        /// Contains the mode of a report to be delivered to a particular message recipient upon completion of physical message delivery or delivery by the message handling system.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_PHYSICAL_DELIVERY_REPORT_REQUEST = PT.PT_LONG | 0x0C0C << 16,
        /// <summary>
        /// Contains the physical forwarding address of a message recipient and is used only with message reports.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_PHYSICAL_FORWARDING_ADDRESS = PT.PT_BINARY | 0x0C0D << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests the message transfer agent to attach a physical forwarding address for a message recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PHYSICAL_FORWARDING_ADDRESS_REQUESTED = PT.PT_BOOLEAN | 0x0C0E << 16,
        /// <summary>
        /// Contains TRUE if a message sender prohibits physical message forwarding for a specific recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PHYSICAL_FORWARDING_PROHIBITED = PT.PT_BOOLEAN | 0x0C0F << 16,
        /// <summary>
        /// Contains an ASN.1 object identifier used for rendering message attachments.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_PHYSICAL_RENDITION_ATTRIBUTES = PT.PT_BINARY | 0x0C10 << 16,
        /// <summary>
        /// Contains an ASN.1 proof of delivery value.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_PROOF_OF_DELIVERY = PT.PT_BINARY | 0x0C11 << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests proof of delivery for a particular recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PROOF_OF_DELIVERY_REQUESTED = PT.PT_BOOLEAN | 0x0C12 << 16,
        /// <summary>
        /// Contains a message recipient's ASN.1 certificate for use in a report.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RECIPIENT_CERTIFICATE = PT.PT_BINARY | 0x0C13 << 16,
        /// <summary>
        /// Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RECIPIENT_NUMBER_FOR_ADVICE = PT.PT_TSTRING | 0x0C14 << 16,
        /// <summary>
        /// Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RECIPIENT_NUMBER_FOR_ADVICE_W = PT.PT_UNICODE | 0x0C14 << 16,
        /// <summary>
        /// Contains a message recipient's voice telephone number to call to advise of the physical delivery of a message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RECIPIENT_NUMBER_FOR_ADVICE_A = PT.PT_STRING8 | 0x0C14 << 16,
        /// <summary>
        /// Contains the recipient type for a message recipient.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RECIPIENT_TYPE = PT.PT_LONG | 0x0C15 << 16,
        /// <summary>
        /// Contains the type of registration used for physical delivery of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_REGISTERED_MAIL_TYPE = PT.PT_LONG | 0x0C16 << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests a reply from a recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_REPLY_REQUESTED = PT.PT_BOOLEAN | 0x0C17 << 16,
        /// <summary>
        /// Contains a binary array of delivery methods (service providers), in order of a message sender's preference.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_REQUESTED_DELIVERY_METHOD                = PT.PT_LONG |      0x0C18 << 16,
        /// <summary>
        /// Contains the message sender's entry identifier.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SENDER_ENTRYID = PT.PT_BINARY | 0x0C19 << 16,
        /// <summary>
        /// Contains the message sender's display name.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SENDER_NAME = PT.PT_TSTRING | 0x0C1A << 16,
        /// <summary>
        /// Contains the message sender's display name. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENDER_NAME_W = PT.PT_UNICODE | 0x0C1A << 16,
        /// <summary>
        /// Contains the message sender's display name. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SENDER_NAME_A = PT.PT_STRING8 | 0x0C1A << 16,
        /// <summary>
        /// Contains additional information for use in a report.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SUPPLEMENTARY_INFO = PT.PT_TSTRING | 0x0C1B << 16,
        /// <summary>
        /// Contains additional information for use in a report. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SUPPLEMENTARY_INFO_W = PT.PT_UNICODE | 0x0C1B << 16,
        /// <summary>
        /// Contains additional information for use in a report. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SUPPLEMENTARY_INFO_A = PT.PT_STRING8 | 0x0C1B << 16,
        /// <summary>
        /// Contains the type of a message recipient, for use in a report.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_TYPE_OF_MTS_USER = PT.PT_LONG | 0x0C1C << 16,
        /// <summary>
        /// Contains the message sender's search key.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SENDER_SEARCH_KEY = PT.PT_BINARY | 0x0C1D << 16,
        /// <summary>
        /// Contains the message sender's e-mail address type.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SENDER_ADDRTYPE = PT.PT_TSTRING | 0x0C1E << 16,
        /// <summary>
        /// Contains the message sender's e-mail address type. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENDER_ADDRTYPE_W = PT.PT_UNICODE | 0x0C1E << 16,
        /// <summary>
        /// Contains the message sender's e-mail address type. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SENDER_ADDRTYPE_A = PT.PT_STRING8 | 0x0C1E << 16,
        /// <summary>
        /// Contains the message sender's e-mail address.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SENDER_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0C1F << 16,
        /// <summary>
        /// Contains the message sender's e-mail address. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENDER_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0C1F << 16,
        /// <summary>
        /// Contains the message sender's e-mail address. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SENDER_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0C1F << 16,

        #endregion

        #region Message envelope properties
        /// <summary>
        /// Represents a null value or setting of a property or reserves array space.
        /// <br></br>Data Type: PT_NULL
        /// </summary>
        PR_NULL = PT.PT_NULL | 0 << 16,
        /// <summary>
        /// Contains the identifier of the mode for message acknowledgment.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ACKNOWLEDGEMENT_MODE = PT.PT_LONG | 0x0001 << 16,
        /// <summary>
        /// Contains TRUE if the sender permits autoforwarding of this message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_ALTERNATE_RECIPIENT_ALLOWED = PT.PT_BOOLEAN | 0x0002 << 16,
        /// <summary>
        /// Contains a list of entry identifiers for users who have authorized the sending of a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_AUTHORIZING_USERS = PT.PT_BINARY | 0x0003 << 16,
        /// <summary>
        /// Contains a comment added by the autoforwarding agent.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_AUTO_FORWARD_COMMENT = PT.PT_TSTRING | 0x0004 << 16,
        /// <summary>
        /// Contains a comment added by the autoforwarding agent. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_AUTO_FORWARD_COMMENT_W = PT.PT_UNICODE | 0x0004 << 16,
        /// <summary>
        /// Contains a comment added by the autoforwarding agent. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_AUTO_FORWARD_COMMENT_A = PT.PT_STRING8 | 0x0004 << 16,
        /// <summary>
        /// Contains TRUE if an automatic agent has forwarded a message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_AUTO_FORWARDED = PT.PT_BOOLEAN | 0x0005 << 16,
        /// <summary>
        /// Contains an identifier for the algorithm used to confirm message content confidentiality.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONTENT_CONFIDENTIALITY_ALGORITHM_ID = PT.PT_BINARY | 0x0006 << 16,
        /// <summary>
        /// Contains a value the message sender can use to match a report with the original message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONTENT_CORRELATOR = PT.PT_BINARY | 0x0007 << 16,
        /// <summary>
        /// Contains a key value that enables the message recipient to identify its content.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_CONTENT_IDENTIFIER = PT.PT_TSTRING | 0x0008 << 16,
        /// <summary>
        /// Contains a key value that enables the message recipient to identify its content. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_CONTENT_IDENTIFIER_W = PT.PT_UNICODE | 0x0008 << 16,
        /// <summary>
        /// Contains a key value that enables the message recipient to identify its content. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_CONTENT_IDENTIFIER_A = PT.PT_STRING8 | 0x0008 << 16,
        /// <summary>
        /// Contains a message length, in bytes, passed to a client application or service provider to determine if a message of that length can be delivered.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_CONTENT_LENGTH = PT.PT_LONG | 0x0009 << 16,
        /// <summary>
        /// Contains TRUE if a message should be returned with a nondelivery report.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_CONTENT_RETURN_REQUESTED = PT.PT_BOOLEAN | 0x000A << 16,
        /// <summary>
        /// Contains the encoded information types (EITs) that are applied to a message in transit to describe conversions.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONVERSION_EITS = PT.PT_BINARY | 0x000C << 16,
        /// <summary>
        /// Contains TRUE if a message transfer agent (MTA) is prohibited from making message text conversions that lose information.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_CONVERSION_WITH_LOSS_PROHIBITED = PT.PT_BOOLEAN | 0x000D << 16,
        /// <summary>
        /// Contains an identifier for the types of text in a message after conversion.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONVERTED_EITS = PT.PT_BINARY | 0x000E << 16,
        /// <summary>
        /// Contains the date and time at which a message sender wants a message delivered.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_DEFERRED_DELIVERY_TIME = PT.PT_SYSTIME | 0x000F << 16,
        /// <summary>
        /// Contains the date and time at which the original message was delivered.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_DELIVER_TIME = PT.PT_SYSTIME | 0x0010 << 16,
        /// <summary>
        /// Contains a reason why a message transfer agent (MTA) has discarded a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_DISCARD_REASON = PT.PT_LONG | 0x0011 << 16,
        /// <summary>
        /// Contains TRUE if disclosure of recipients is allowed.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_DISCLOSURE_OF_RECIPIENTS = PT.PT_BOOLEAN | 0x0012 << 16,
        /// <summary>
        /// Contains a history showing how a distribution list has been expanded during message transmission.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_DL_EXPANSION_HISTORY = PT.PT_BINARY | 0x0013 << 16,
        /// <summary>
        /// Contains TRUE if a message transfer agent (MTA) is prohibited from expanding distribution lists.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_DL_EXPANSION_PROHIBITED = PT.PT_BOOLEAN | 0x0014 << 16,
        /// <summary>
        /// Contains the date and time at which the messaging system can invalidate the content of a message.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_EXPIRY_TIME = PT.PT_SYSTIME | 0x0015 << 16,
        /// <summary>
        /// Contains TRUE if a message transfer agent (MTA) is prohibited from making implicit message text conversions.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_IMPLICIT_CONVERSION_PROHIBITED = PT.PT_BOOLEAN | 0x0016 << 16,
        /// <summary>
        /// Contains a value indicating the message sender's opinion of the importance of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_IMPORTANCE = PT.PT_LONG | 0x0017 << 16,
        /// <summary>
        /// Contains the latest date and time when a message transfer agent (MTA) should deliver a message.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_LATEST_DELIVERY_TIME = PT.PT_SYSTIME | 0x0019 << 16,
        /// <summary>
        /// Contains a text string that identifies the sender-defined message class, such as IPM.Note.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_MESSAGE_CLASS = PT.PT_TSTRING | 0x001A << 16,
        /// <summary>
        /// Contains a text string that identifies the sender-defined message class, such as IPM.Note. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_MESSAGE_CLASS_W = PT.PT_UNICODE | 0x001A << 16,
        /// <summary>
        /// Contains a text string that identifies the sender-defined message class, such as IPM.Note. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_MESSAGE_CLASS_A = PT.PT_STRING8 | 0x001A << 16,
        /// <summary>
        /// Contains a message transfer system (MTS) identifier for a message delivered to a client application.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MESSAGE_DELIVERY_ID = PT.PT_BINARY | 0x001B << 16,
        /// <summary>
        /// Contains a security label for a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MESSAGE_SECURITY_LABEL = PT.PT_BINARY | 0x001E << 16,
        /// <summary>
        /// Contains the identifiers of messages that this message supersedes
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_OBSOLETED_IPMS = PT.PT_BINARY | 0x001F << 16,
        /// <summary>
        /// Contains the encoded name of the originally intended recipient of an autoforwarded message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIPIENT_NAME = PT.PT_BINARY | 0x0020 << 16,
        /// <summary>
        /// Contains a copy of the original encoded information types (EITs) for message text.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_EITS = PT.PT_BINARY | 0x0021 << 16,
        /// <summary>
        /// Contains an ASN.1 certificate for the message originator.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINATOR_CERTIFICATE = PT.PT_BINARY | 0x0022 << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests a delivery report for a particular recipient from the messaging system before the message is placed in the message store.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED = PT.PT_BOOLEAN | 0x0023 << 16,
        /// <summary>
        /// Contains the binary-encoded return address of the message originator.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINATOR_RETURN_ADDRESS = PT.PT_BINARY | 0x0024 << 16,
        /// <summary>
        /// Contains the relative priority of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_PRIORITY = PT.PT_LONG | 0x0026 << 16,
        /// <summary>
        /// Contains a binary verification value enabling a delivery report recipient to verify the origin of the original message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGIN_CHECK = PT.PT_BINARY | 0x0027 << 16,
        /// <summary>
        /// Contains TRUE if a message sender requests proof that the message transfer system has submitted a message for delivery to the originally intended recipient.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_PROOF_OF_SUBMISSION_REQUESTED = PT.PT_BOOLEAN | 0x0028 << 16,
        /// <summary>
        /// Contains TRUE if a message sender wants the messaging system to generate a read report when the recipient has read a message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_READ_RECEIPT_REQUESTED = PT.PT_BOOLEAN | 0x0029 << 16,
        /// <summary>
        /// Contains the date and time a delivery report is generated.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_RECEIPT_TIME = PT.PT_SYSTIME | 0x002A << 16,
        /// <summary>
        /// Contains TRUE if recipient reassignment is prohibited.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_RECIPIENT_REASSIGNMENT_PROHIBITED = PT.PT_BOOLEAN | 0x002B << 16,
        /// <summary>
        /// Contains information about the route covered by a delivered message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REDIRECTION_HISTORY = PT.PT_BINARY | 0x002C << 16,
        /// <summary>
        /// Contains a list of identifiers for messages to which a message is related.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RELATED_IPMS = PT.PT_BINARY | 0x002D << 16,
        /// <summary>
        /// Contains the sensitivity value assigned by the sender of the first version of a message — that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ORIGINAL_SENSITIVITY = PT.PT_LONG | 0x002E << 16,
        /// <summary>
        /// Contains an ASCII list of the languages incorporated in a message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_LANGUAGES = PT.PT_TSTRING | 0x002F << 16,
        /// <summary>
        /// Contains an ASCII list of the languages incorporated in a message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_LANGUAGES_W = PT.PT_UNICODE | 0x002F << 16,
        /// <summary>
        /// Contains an ASCII list of the languages incorporated in a message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_LANGUAGES_A = PT.PT_STRING8 | 0x002F << 16,
        /// <summary>
        /// Contains the date and time by which a reply is expected for a message.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_REPLY_TIME = PT.PT_SYSTIME | 0x0030 << 16,
        /// <summary>
        /// Contains a binary tag value that the messaging system should copy to any report generated for the message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPORT_TAG = PT.PT_BINARY | 0x0031 << 16,
        /// <summary>
        /// Contains the date and time when the messaging system generated a report.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_REPORT_TIME = PT.PT_SYSTIME | 0x0032 << 16,
        /// <summary>
        /// Contains TRUE if the original message is being returned with a nonread report.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_RETURNED_IPM = PT.PT_BOOLEAN | 0x0033 << 16,
        /// <summary>
        /// Contains a flag that indicates the security level of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_SECURITY = PT.PT_LONG | 0x0034 << 16,
        /// <summary>
        /// Contains TRUE if this message is an incomplete copy of another message.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_INCOMPLETE_COPY = PT.PT_BOOLEAN | 0x0035 << 16,
        /// <summary>
        /// Contains a value indicating the message sender's opinion of the sensitivity of a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_SENSITIVITY = PT.PT_LONG | 0x0036 << 16,
        /// <summary>
        /// Contains the full subject of a message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SUBJECT = PT.PT_TSTRING | 0x0037 << 16,
        /// <summary>
        /// Contains the full subject of a message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SUBJECT_W = PT.PT_UNICODE | 0x0037 << 16,
        /// <summary>
        /// Contains the full subject of a message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SUBJECT_A = PT.PT_STRING8 | 0x0037 << 16,
        /// <summary>
        /// Contains a binary value that is copied from the message for which a report is being generated.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SUBJECT_IPM = PT.PT_BINARY | 0x0038 << 16,
        /// <summary>
        /// Contains the date and time the message sender submitted a message.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_CLIENT_SUBMIT_TIME = PT.PT_SYSTIME | 0x0039 << 16,
        /// <summary>
        /// Contains the display name for the recipient that should get reports for this message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_REPORT_NAME = PT.PT_TSTRING | 0x003A << 16,
        /// <summary>
        /// Contains the display name for the recipient that should get reports for this message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_REPORT_NAME_W = PT.PT_UNICODE | 0x003A << 16,
        /// <summary>
        /// Contains the display name for the recipient that should get reports for this message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_REPORT_NAME_A = PT.PT_STRING8 | 0x003A << 16,
        /// <summary>
        /// Contains the search key for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SENT_REPRESENTING_SEARCH_KEY = PT.PT_BINARY | 0x003B << 16,
        /// <summary>
        /// Contains the content type for a submitted message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_X400_CONTENT_TYPE = PT.PT_BINARY | 0x003C << 16,
        /// <summary>
        /// Contains a subject prefix that typically indicates some action on a message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SUBJECT_PREFIX = PT.PT_TSTRING | 0x003D << 16,
        /// <summary>
        /// Contains a subject prefix that typically indicates some action on a message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SUBJECT_PREFIX_W = PT.PT_UNICODE | 0x003D << 16,
        /// <summary>
        /// Contains a subject prefix that typically indicates some action on a message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SUBJECT_PREFIX_A = PT.PT_STRING8 | 0x003D << 16,
        /// <summary>
        /// dContains reasons why a message was not received that forms part of a nondelivery report.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_NON_RECEIPT_REASON = PT.PT_LONG | 0x003E << 16,
        /// <summary>
        /// Contains the entry identifier of the messaging user that actually receives the message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RECEIVED_BY_ENTRYID = PT.PT_BINARY | 0x003F << 16,
        /// <summary>
        /// Contains the display name of the messaging user that actually receives the message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RECEIVED_BY_NAME = PT.PT_TSTRING | 0x0040 << 16,
        /// <summary>
        /// Contains the display name of the messaging user that actually receives the message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RECEIVED_BY_NAME_W = PT.PT_UNICODE | 0x0040 << 16,
        /// <summary>
        /// Contains the display name of the messaging user that actually receives the message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RECEIVED_BY_NAME_A = PT.PT_STRING8 | 0x0040 << 16,
        /// <summary>
        /// Contains the entry identifier for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_SENT_REPRESENTING_ENTRYID = PT.PT_BINARY | 0x0041 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SENT_REPRESENTING_NAME = PT.PT_TSTRING | 0x0042 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the sender. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENT_REPRESENTING_NAME_W = PT.PT_UNICODE | 0x0042 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the sender. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_SENT_REPRESENTING_NAME_A = PT.PT_STRING8 | 0x0042 << 16,
        /// <summary>
        /// Contains the entry identifier for the messaging user represented by the receiving user.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RCVD_REPRESENTING_ENTRYID = PT.PT_BINARY | 0x0043 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the receiving user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RCVD_REPRESENTING_NAME = PT.PT_TSTRING | 0x0044 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the receiving user. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RCVD_REPRESENTING_NAME_W = PT.PT_UNICODE | 0x0044 << 16,
        /// <summary>
        /// Contains the display name for the messaging user represented by the receiving user. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RCVD_REPRESENTING_NAME_A = PT.PT_STRING8 | 0x0044 << 16,
        /// <summary>
        /// Contains the entry identifier for the recipient that should get reports for this message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPORT_ENTRYID = PT.PT_BINARY | 0x0045 << 16,
        /// <summary>
        /// Contains an entry identifier for the messaging user to which the messaging system should direct a read report for this message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_READ_RECEIPT_ENTRYID = PT.PT_BINARY | 0x0046 << 16,
        /// <summary>
        /// Contains a message transfer system (MTS) identifier for the message transfer agent (MTA).
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_MESSAGE_SUBMISSION_ID = PT.PT_BINARY | 0x0047 << 16,
        /// <summary>
        /// Contains the date and time a transport provider passed a message to its underlying messaging system.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_PROVIDER_SUBMIT_TIME = PT.PT_SYSTIME | 0x0048 << 16,
        /// <summary>
        /// Contains the subject of an original message for use in a report about the message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SUBJECT = PT.PT_TSTRING | 0x0049 << 16,
        /// <summary>
        /// Contains the subject of an original message for use in a report about the message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SUBJECT_W = PT.PT_UNICODE | 0x0049 << 16,
        /// <summary>
        /// Contains the subject of an original message for use in a report about the message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SUBJECT_A = PT.PT_STRING8 | 0x0049 << 16,
        /// <summary>
        /// Contains the class of the original message for use in a report.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIG_MESSAGE_CLASS = PT.PT_TSTRING | 0x004B << 16,
        /// <summary>
        /// Contains the class of the original message for use in a report. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIG_MESSAGE_CLASS_W = PT.PT_UNICODE | 0x004B << 16,
        /// <summary>
        /// Contains the class of the original message for use in a report. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIG_MESSAGE_CLASS_A = PT.PT_STRING8 | 0x004B << 16,
        /// <summary>
        /// Contains the entry identifier of the author of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_AUTHOR_ENTRYID = PT.PT_BINARY | 0x004C << 16,
        /// <summary>
        /// Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_AUTHOR_NAME = PT.PT_TSTRING | 0x004D << 16,
        /// <summary>
        /// Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_AUTHOR_NAME_W = PT.PT_UNICODE | 0x004D << 16,
        /// <summary>
        /// Contains the display name of the author of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_AUTHOR_NAME_A = PT.PT_STRING8 | 0x004D << 16,
        /// <summary>
        /// Contains the original submission date and time of the message in the report.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_ORIGINAL_SUBMIT_TIME = PT.PT_SYSTIME | 0x004E << 16,
        /// <summary>
        /// Contains a sized array of entry identifiers for recipients that are to get a reply.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPLY_RECIPIENT_ENTRIES = PT.PT_BINARY | 0x004F << 16,
        /// <summary>
        /// Contains a list of display names for recipients that are to get a reply.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_REPLY_RECIPIENT_NAMES = PT.PT_TSTRING | 0x0050 << 16,
        /// <summary>
        /// Contains a list of display names for recipients that are to get a reply. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_REPLY_RECIPIENT_NAMES_W = PT.PT_UNICODE | 0x0050 << 16,
        /// <summary>
        /// Contains a list of display names for recipients that are to get a reply. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_REPLY_RECIPIENT_NAMES_A = PT.PT_STRING8 | 0x0050 << 16,
        /// <summary>
        /// Contains the search key of the messaging user that actually receives the message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RECEIVED_BY_SEARCH_KEY = PT.PT_BINARY | 0x0051 << 16,
        /// <summary>
        /// Contains the search key for the messaging user represented by the receiving user.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_RCVD_REPRESENTING_SEARCH_KEY = PT.PT_BINARY | 0x0052 << 16,
        /// <summary>
        /// Contains a search key for the messaging user to which the messaging system should direct a read report for a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_READ_RECEIPT_SEARCH_KEY = PT.PT_BINARY | 0x0053 << 16,
        /// <summary>
        /// Contains the search key for the recipient that should get reports for this message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPORT_SEARCH_KEY = PT.PT_BINARY | 0x0054 << 16,
        /// <summary>
        /// Contains a copy of the original message's delivery date and time in a thread.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_ORIGINAL_DELIVERY_TIME = PT.PT_SYSTIME | 0x0055 << 16,
        /// <summary>
        /// Contains the search key of the author of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_AUTHOR_SEARCH_KEY = PT.PT_BINARY | 0x0056 << 16,
        /// <summary>
        /// Contains TRUE if this messaging user is specifically named as a primary (To) recipient of this message and is not part of a distribution list.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_MESSAGE_TO_ME = PT.PT_BOOLEAN | 0x0057 << 16,
        /// <summary>
        /// Contains TRUE if this messaging user is specifically named as a carbon copy (CC) recipient of this message and is not part of a distribution list.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_MESSAGE_CC_ME = PT.PT_BOOLEAN | 0x0058 << 16,
        /// <summary>
        /// Contains TRUE if this messaging user is specifically named as a primary (To), carbon copy (CC), or blind carbon copy (BCC) recipient of this message and is not part of a distribution list.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_MESSAGE_RECIP_ME = PT.PT_BOOLEAN | 0x0059 << 16,
        /// <summary>
        /// Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SENDER_NAME = PT.PT_TSTRING | 0x005A << 16,
        /// <summary>
        /// Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENDER_NAME_W = PT.PT_UNICODE | 0x005A << 16,
        /// <summary>
        /// Contains the display name of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENDER_NAME_A = PT.PT_STRING8 | 0x005A << 16,
        /// <summary>
        /// Contains the entry identifier of the sender of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_SENDER_ENTRYID = PT.PT_BINARY | 0x005B << 16,
        /// <summary>
        /// Contains the search key for the sender of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_SENDER_SEARCH_KEY = PT.PT_BINARY | 0x005C << 16,
        /// <summary>
        /// Contains the display name of the messaging user on whose behalf the original message was sent.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_NAME = PT.PT_TSTRING | 0x005D << 16,
        /// <summary>
        /// Contains the display name of the messaging user on whose behalf the original message was sent. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_NAME_W = PT.PT_UNICODE | 0x005D << 16,
        /// <summary>
        /// Contains the display name of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_NAME_A = PT.PT_STRING8 | 0x005D << 16,
        /// <summary>
        /// Contains the entry identifier of the messaging user on whose behalf the original message was sent.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_ENTRYID = PT.PT_BINARY | 0x005E << 16,
        /// <summary>
        /// Contains the search key of the messaging user on whose behalf the original message was sent.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_SEARCH_KEY = PT.PT_BINARY | 0x005F << 16,
        /// <summary>
        /// Contains the starting date and time of an appointment as managed by a scheduling application.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_START_DATE = PT.PT_SYSTIME | 0x0060 << 16,
        /// <summary>
        /// Contains the ending date and time of an appointment as managed by a scheduling application.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_END_DATE = PT.PT_SYSTIME | 0x0061 << 16,
        /// <summary>
        /// Contains an identifier for an appointment in the owner's schedule.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_OWNER_APPT_ID = PT.PT_LONG | 0x0062 << 16,
        /// <summary>
        /// Contains TRUE if the message sender wants a response to a meeting request.
        /// <br></br>Data Type: PT_BOOLEAN
        /// </summary>
        PR_RESPONSE_REQUESTED = PT.PT_BOOLEAN |   0x0063 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_TSTRING 
        /// </summary>
        PR_SENT_REPRESENTING_ADDRTYPE = PT.PT_TSTRING | 0x0064 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the sender. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENT_REPRESENTING_ADDRTYPE_W = PT.PT_UNICODE | 0x0064 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the sender. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8 
        /// </summary>
        PR_SENT_REPRESENTING_ADDRTYPE_A = PT.PT_STRING8 | 0x0064 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_TSTRING 
        /// </summary>
        PR_SENT_REPRESENTING_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0065 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the sender. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_SENT_REPRESENTING_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0065 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the sender. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8 
        /// </summary>
        PR_SENT_REPRESENTING_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0065 << 16,
        /// <summary>
        /// Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING 
        /// </summary>
        PR_ORIGINAL_SENDER_ADDRTYPE = PT.PT_TSTRING | 0x0066 << 16,
        /// <summary>
        /// Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENDER_ADDRTYPE_W = PT.PT_UNICODE | 0x0066 << 16,
        /// <summary>
        /// Contains the address type of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENDER_ADDRTYPE_A = PT.PT_STRING8 | 0x0066 << 16,
        /// <summary>
        /// Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SENDER_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0067 << 16,
        /// <summary>
        /// Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENDER_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0067 << 16,
        /// <summary>
        /// Contains the e-mail address of the sender of the first version of a message, that is, the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENDER_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0067 << 16,
        /// <summary>
        /// Contains the address type of the messaging user on whose behalf the original message was sent.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE = PT.PT_TSTRING | 0x0068 << 16,
        /// <summary>
        /// Contains the address type of the messaging user on whose behalf the original message was sent. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_W = PT.PT_UNICODE | 0x0068 << 16,
        /// <summary>
        /// Contains the address type of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_ADDRTYPE_A = PT.PT_STRING8 | 0x0068 << 16,
        /// <summary>
        /// Contains the e-mail address of the messaging user on whose behalf the original message was sent.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0069 << 16,
        /// <summary>
        /// Contains the e-mail address of the messaging user on whose behalf the original message was sent. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0069 << 16,
        /// <summary>
        /// Contains the e-mail address of the messaging user on whose behalf the original message was sent. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_SENT_REPRESENTING_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0069 << 16,
        /// <summary>
        /// Contains the topic of the first message in a conversation thread.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_CONVERSATION_TOPIC = PT.PT_TSTRING | 0x0070 << 16,
        /// <summary>
        /// Contains the topic of the first message in a conversation thread. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_CONVERSATION_TOPIC_W = PT.PT_UNICODE | 0x0070 << 16,
        /// <summary>
        /// Contains the topic of the first message in a conversation thread. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_CONVERSATION_TOPIC_A = PT.PT_STRING8 | 0x0070 << 16,
        /// <summary>
        /// Contains a binary value that indicates the relative position of this message within a conversation thread.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_CONVERSATION_INDEX = PT.PT_BINARY | 0x0071 << 16,
        /// <summary>
        /// Contains the display names of any blind carbon copy (BCC) recipients of the original message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_DISPLAY_BCC = PT.PT_TSTRING | 0x0072 << 16,
        /// <summary>
        /// Contains the display names of any blind carbon copy (BCC) recipients of the original message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_DISPLAY_BCC_W = PT.PT_UNICODE | 0x0072 << 16,
        /// <summary>
        /// Contains the display names of any blind carbon copy (BCC) recipients of the original message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_DISPLAY_BCC_A = PT.PT_STRING8 | 0x0072 << 16,
        /// <summary>
        /// Contains the display names of any carbon copy (CC) recipients of the original message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_DISPLAY_CC = PT.PT_TSTRING | 0x0073 << 16,
        /// <summary>
        /// Contains the display names of any carbon copy (CC) recipients of the original message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_DISPLAY_CC_W = PT.PT_UNICODE | 0x0073 << 16,
        /// <summary>
        /// Contains the display names of any carbon copy (CC) recipients of the original message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_DISPLAY_CC_A = PT.PT_STRING8 | 0x0073 << 16,
        /// <summary>
        /// Contains the display names of the primary (To) recipients of the original message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_DISPLAY_TO = PT.PT_TSTRING | 0x0074 << 16,
        /// <summary>
        /// Contains the display names of the primary (To) recipients of the original message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_DISPLAY_TO_W = PT.PT_UNICODE | 0x0074 << 16,
        /// <summary>
        /// Contains the display names of the primary (To) recipients of the original message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_DISPLAY_TO_A = PT.PT_STRING8 | 0x0074 << 16,
        /// <summary>
        /// Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RECEIVED_BY_ADDRTYPE = PT.PT_TSTRING | 0x0075 << 16,
        /// <summary>
        /// Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RECEIVED_BY_ADDRTYPE_W = PT.PT_UNICODE | 0x0075 << 16,
        /// <summary>
        /// Contains the e-mail address type, such as SMTP, for the messaging user that actually receives the message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RECEIVED_BY_ADDRTYPE_A = PT.PT_STRING8 | 0x0075 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user that actually receives the message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RECEIVED_BY_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0076 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user that actually receives the message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RECEIVED_BY_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0076 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user that actually receives the message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RECEIVED_BY_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0076 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the user actually receiving the message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RCVD_REPRESENTING_ADDRTYPE = PT.PT_TSTRING | 0x0077 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the user actually receiving the message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RCVD_REPRESENTING_ADDRTYPE_W = PT.PT_UNICODE | 0x0077 << 16,
        /// <summary>
        /// Contains the address type for the messaging user represented by the user actually receiving the message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RCVD_REPRESENTING_ADDRTYPE_A = PT.PT_STRING8 | 0x0077 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the receiving user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RCVD_REPRESENTING_EMAIL_ADDRESS = PT.PT_TSTRING | 0x0078 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the receiving user. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_RCVD_REPRESENTING_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x0078 << 16,
        /// <summary>
        /// Contains the e-mail address for the messaging user represented by the receiving user. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_RCVD_REPRESENTING_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x0078 << 16,
        /// <summary>
        /// Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_AUTHOR_ADDRTYPE = PT.PT_TSTRING | 0x0079 << 16,
        /// <summary>
        /// Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_AUTHOR_ADDRTYPE_W = PT.PT_UNICODE | 0x0079 << 16,
        /// <summary>
        /// Contains the address type of the author of the first version of a message. That is — the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_AUTHOR_ADDRTYPE_A = PT.PT_STRING8 | 0x0079 << 16,
        /// <summary>
        /// Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS = PT.PT_TSTRING | 0x007A << 16,
        /// <summary>
        /// Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x007A << 16,
        /// <summary>
        /// Contains the e-mail address of the author of the first version of a message. That is — the message before being forwarded or replied to. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINAL_AUTHOR_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x007A << 16,
        /// <summary>
        /// Contains the address type of the originally intended recipient of an autoforwarded message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE = PT.PT_TSTRING | 0x007B << 16,
        /// <summary>
        /// Contains the address type of the originally intended recipient of an autoforwarded message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_W = PT.PT_UNICODE | 0x007B << 16,
        /// <summary>
        /// Contains the address type of the originally intended recipient of an autoforwarded message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_ADDRTYPE_A = PT.PT_STRING8 | 0x007B << 16,
        /// <summary>
        /// Contains the e-mail address of the originally intended recipient of an autoforwarded message.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS = PT.PT_TSTRING | 0x007C << 16,
        /// <summary>
        /// Contains the e-mail address of the originally intended recipient of an autoforwarded message. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_W = PT.PT_UNICODE | 0x007C << 16,
        /// <summary>
        /// Contains the e-mail address of the originally intended recipient of an autoforwarded message. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_ORIGINALLY_INTENDED_RECIP_EMAIL_ADDRESS_A = PT.PT_STRING8 | 0x007C << 16,
        /// <summary>
        /// Contains transport-specific message envelope information
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_TRANSPORT_MESSAGE_HEADERS = PT.PT_TSTRING | 0x007D << 16,
        /// <summary>
        /// Contains transport-specific message envelope information. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_TRANSPORT_MESSAGE_HEADERS_W = PT.PT_UNICODE | 0x007D << 16,
        /// <summary>
        /// Contains transport-specific message envelope information. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_TRANSPORT_MESSAGE_HEADERS_A = PT.PT_STRING8 | 0x007D << 16,
        /// <summary>
        /// Contains the converted value of the attDelegate workgroup property.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_DELEGATION = PT.PT_BINARY | 0x007E << 16,
        /// <summary>
        /// Contains a value used to correlate a Transport Neutral Encapsulation Format (TNEF) attachment with a message.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_TNEF_CORRELATION_KEY = PT.PT_BINARY | 0x007F << 16,

        #endregion

        #region Message content properties
        /// <summary>
        /// Contains the message text.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_BODY = PT.PT_TSTRING | 0x1000 << 16,
        /// <summary>
        /// Contains the message text. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_BODY_W = PT.PT_UNICODE | 0x1000 << 16,
        /// <summary>
        /// Contains the message text. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_BODY_A = PT.PT_STRING8 | 0x1000 << 16,
        /// <summary>
        /// Contains optional text for a report generated by the messaging system.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_REPORT_TEXT = PT.PT_TSTRING | 0x1001 << 16,
        /// <summary>
        /// Contains optional text for a report generated by the messaging system. UNICODE compilation.
        /// <br></br>Data Type: PT_UNICODE
        /// </summary>
        PR_REPORT_TEXT_W = PT.PT_UNICODE | 0x1001 << 16,
        /// <summary>
        /// Contains optional text for a report generated by the messaging system. Non-UNICODE compilation.
        /// <br></br>Data Type: PT_STRING8
        /// </summary>
        PR_REPORT_TEXT_A = PT.PT_STRING8 | 0x1001 << 16,
        /// <summary>
        /// Contains information about a message originator and a distribution list expansion history.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ORIGINATOR_AND_DL_EXPANSION_HISTORY = PT.PT_BINARY | 0x1002 << 16,
        /// <summary>
        /// Contains the display name of a distribution list for which the messaging system is delivering a report.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPORTING_DL_NAME = PT.PT_BINARY | 0x1003 << 16,
        /// <summary>
        /// Contains an identifier for the message transfer agent that generated a report.
        ///  <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_REPORTING_MTA_CERTIFICATE = PT.PT_BINARY | 0x1004 << 16,
        /// <summary>
        /// Contains the cyclical redundancy check (CRC) computed for the message text.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RTF_SYNC_BODY_CRC = PT.PT_LONG | 0x1006 << 16,
        /// <summary>
        /// Contains a count of the significant characters of the message text.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RTF_SYNC_BODY_COUNT = PT.PT_LONG | 0x1007 << 16,
        /// <summary>
        /// Contains significant characters that appear at the beginning of the message text.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_RTF_SYNC_BODY_TAG = PT.PT_TSTRING | 0x1008 << 16,
        /// <summary>
        /// Contains a count of the ignorable characters that appear before the significant characters of the message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RTF_SYNC_PREFIX_COUNT = PT.PT_LONG | 0x1010 << 16,
        /// <summary>
        /// Contains a count of the ignorable characters that appear after the significant characters of the message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RTF_SYNC_TRAILING_COUNT = PT.PT_LONG | 0x1011 << 16,
        #endregion

        #region Mail User Properties
        /// <summary>
        /// Contains the display name for the messaging user represented by the sender.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_GIVEN_NAME = PT.PT_TSTRING | 0x3A06 << 16,
        /// <summary>
        /// Contains the last name of the messaging user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SURNAME = PT.PT_TSTRING | 0x3A11 << 16,
        /// <summary>
        /// Contains the middle name of the messaging user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_MIDDLE_NAME = PT.PT_TSTRING | 0x3A44 << 16,
        /// <summary>
        /// Contains the name of the messaging user's line of business
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PROFESSION = PT.PT_TSTRING | 0x3A46 << 16,
        /// <summary>
        /// Contains the Web address, also know as the Uniform Resource Locator (URL), of the business home page of the messaging user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_BUSINESS_HOME_PAGE = PT.PT_TSTRING | 0x3A51 << 16,
        /// <summary>
        /// Contains the Web address, also known as the Uniform Resource Locator (URL), of the messaging user's personal home page.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PERSONAL_HOME_PAGE = PT.PT_TSTRING | 0x3A50 << 16,
        /// <summary>
        /// Contains the postal address of the messaging user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_POSTAL_ADDRESS = PT.PT_TSTRING | 0x3A15 << 16,
        /// <summary>
        /// Contains the job title of the messaging user.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_TITLE = PT.PT_TSTRING | 0x3A17 << 16,
        /// <summary>
        /// Contains the recipient's company name.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_COMPANY_NAME = PT.PT_TSTRING | 0x3A16 << 16,
        /// <summary>
        /// Contains the display name prefix for a given MAPI object.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DISPLAY_NAME_PREFIX = PT.PT_TSTRING | 0x3A45 << 16,
        /// <summary>
        /// Contains a generational abbreviation that follows the full name of the recipient.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_GENERATION = PT.PT_TSTRING | 0x3A05 << 16,
        /// <summary>
        /// Contains a name for the department in which the recipient works.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_DEPARTMENT_NAME = PT.PT_TSTRING | 0x3A18 << 16,
        /// <summary>
        /// Contains the recipient's office location.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_OFFICE_LOCATION = PT.PT_TSTRING | 0x3A19 << 16,
        /// <summary>
        /// Contains the name of the recipient's manager.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_MANAGER_NAME = PT.PT_TSTRING | 0x3A4E << 16,
        /// <summary>
        /// Contains the name of the recipient's administrative assistant.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ASSISTANT = PT.PT_TSTRING | 0x3A30 << 16,
        /// <summary>
        /// Contains the recipient's nickname.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_NICKNAME = PT.PT_TSTRING | 0x3A4F << 16,
        /// <summary>
        /// Contains the name of the recipient's spouse/partner.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SPOUSE_NAME = PT.PT_TSTRING | 0x3A48 << 16,
        /// <summary>
        /// Contains the date of the recipient's wedding anniversary.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_WEDDING_ANNIVERSARY = PT.PT_SYSTIME | 0x3A41 << 16,
        /// <summary>
        /// Contains the date of the recipient's birthday.
        /// <br></br>Data Type: PT_SYSTIME
        /// </summary>
        PR_BIRTHDAY = PT.PT_SYSTIME | 0x3A42 << 16,
        /// <summary>
        /// Contains the recipient's primary telephone number.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PRIMARY_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A1A << 16,
        /// <summary>
        /// Contains the telephone number of the recipient's primary fax machine.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_PRIMARY_FAX_NUMBER = PT.PT_TSTRING | 0x3A23 << 16,
        /// <summary>
        /// Contains the primary telephone number of the recipient's place of business.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_BUSINESS_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A08 << 16,
        /// <summary>
        /// Contains a secondary telephone number at the recipient's place of business.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_BUSINESS2_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A1B << 16,
        /// <summary>
        /// Contains the primary telephone number of the recipient's home.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_HOME_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A09 << 16,
        /// <summary>
        /// Contains a secondary telephone number at the recipient's home.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_HOME2_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A2F << 16,
        /// <summary>
        /// Contains the telephone number of the recipient's business fax machine.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_BUSINESS_FAX_NUMBER = PT.PT_TSTRING | 0x3A24 << 16,
        /// <summary>
        /// Contains the telephone number of the recipient's home fax machine.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_HOME_FAX_NUMBER = PT.PT_TSTRING | 0x3A25 << 16,
        /// <summary>
        /// Contains the recipient's mobile telephone number.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_MOBILE_TELEPHONE_NUMBER = PT.PT_TSTRING | 0x3A1C << 16,
        #endregion

        #region Attachment Properties

        /// <summary>
        /// Contains an attachment object typically accessed through the OLE IStorage:IUnknown interface.
        /// <br></br>Data Type: PT_OBJECT
        /// </summary>
        PR_ATTACH_DATA_OBJ = PT.PT_OBJECT | 0x3701 << 16,
        /// <summary>
        /// Contains binary attachment data typically accessed through the COM IStream:IUnknown interface.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ATTACH_DATA_BIN = PT.PT_BINARY | 0x3701 << 16,
        /// <summary>
        /// Contains an ASN.1 object identifier specifying the encoding for an attachment.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ATTACH_ENCODING = PT.PT_BINARY | 0x3702 << 16,
        /// <summary>
        /// Contains a filename extension that indicates the document type of an attachment.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_EXTENSION = PT.PT_TSTRING | 0x3703 << 16,
        /// <summary>
        /// Contains an attachment's base filename and extension, excluding path.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_FILENAME = PT.PT_TSTRING | 0x3704 << 16,
        /// <summary>
        /// Contains a MAPI-defined constant representing the way the contents of an attachment can be accessed.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_ATTACH_METHOD = PT.PT_LONG | 0x3705 << 16,
        /// <summary>
        /// Contains an attachment's long filename and extension, excluding path.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_LONG_FILENAME = PT.PT_TSTRING | 0x3707 << 16,
        /// <summary>
        /// Contains an attachment's fully qualified path and filename.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_PATHNAME = PT.PT_TSTRING | 0x3708 << 16,
        /// <summary>
        /// Contains a Microsoft Windows metafile with rendering information for an attachment.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ATTACH_RENDERING = PT.PT_BINARY | 0x3709 << 16,
        /// <summary>
        /// Contains an ASN.1 object identifier specifying the application that supplied an attachment.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ATTACH_TAG = PT.PT_BINARY | 0x370A << 16,
        /// <summary>
        /// Contains an offset, in characters, to use in rendering an attachment within the main message text.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RENDERING_POSITION = PT.PT_LONG | 0x370B << 16,
        /// <summary>
        /// Contains the name of an attachment file modified so that it can be correlated with TNEF messages.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_TRANSPORT_NAME = PT.PT_TSTRING | 0x370C << 16,
        /// <summary>
        /// Contains an attachment's fully qualified long path and filename.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_LONG_PATHNAME = PT.PT_TSTRING | 0x370D << 16,
        /// <summary>
        /// Contains formatting information about a Multipurpose Internet Mail Extensions (MIME) attachment.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_ATTACH_MIME_TAG = PT.PT_TSTRING | 0x370E << 16,
        /// <summary>
        /// Provides file type information for a non-Windows attachment.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_ATTACH_ADDITIONAL_INFO = PT.PT_BINARY | 0x370F << 16,

        #endregion

        #region Misc
        /// <summary>
        /// Specifies the format for an editor to use to display a message.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_MSG_EDITOR_FORMAT = PT.PT_LONG | 0x5909 << 16,
        /// <summary>
        /// Contains a value indicating the service provider type.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_RESOURCE_TYPE = PT.PT_LONG | 0x3E03 << 16,
        /// <summary>
        /// Contains the SMTP address for the address book object.
        /// <br></br>Data Type: PT_TSTRING
        /// </summary>
        PR_SMTP_ADDRESS = PT.PT_TSTRING | 0x39FE << 16,
        /// <summary>
        /// Contains a bitmask of flags indicating the current status of a session resource. All service providers set status codes as does MAPI to report on the status of the subsystem, the MAPI spooler, and the integrated address book.
        /// <br></br>Data Type: PT_LONG
        /// </summary>
        PR_STATUS_CODE = PT.PT_LONG | 0x3E04 << 16,

        PR_IPM_APPOINTMENT_ENTRYID  = PT.PT_BINARY | 0x36D0 << 16,
        PR_IPM_CONTACT_ENTRYID = PT.PT_BINARY | 0x36D1 << 16,
        PR_IPM_JOURNAL_ENTRYID = PT.PT_BINARY | 0x36D2 << 16,
        PR_IPM_NOTE_ENTRYID = PT.PT_BINARY | 0x36D3 << 16,
        PR_IPM_TASK_ENTRYID = PT.PT_BINARY | 0x36D4 << 16,
        /// <summary>
        /// Contains the EntryID of the Outlook Drafts folder.
        /// <br></br>Data Type: PT_BINARY
        /// </summary>
        PR_IPM_DRAFTS_ENTRYID  = PT.PT_BINARY | 0x36D7 << 16,
        PR_CONTAINER_CLASS = PT.PT_TSTRING | 0x3613 << 16,
        //PR_DEF_POST_MSGCLASS = 0x36E5001E,
        PR_BODY_HTML =  PT.PT_TSTRING | 0x1013 << 16,

        PR_RTF_COMPRESSED = PT.PT_BINARY |	0x1009 << 16,

        #endregion
    }

    /// <summary>
    /// Property Type
    /// </summary>
    public enum PT : uint
    {
        /// <summary>
        /// (Reserved for interface use) type doesn't matter to caller
        /// </summary>
        PT_UNSPECIFIED = 0,
        /// <summary>
        /// NULL property value
        /// </summary>
        PT_NULL = 1,
        /// <summary>
        /// Signed 16-bit value
        /// </summary>
        PT_I2 = 2,
        /// <summary>
        /// Signed 32-bit value
        /// </summary>
        PT_LONG = 3,
        /// <summary>
        /// 4-byte floating point
        /// </summary>
        PT_R4 = 4,
        /// <summary>
        /// Floating point double
        /// </summary>
        PT_DOUBLE = 5,
        /// <summary>
        /// Signed 64-bit int (decimal w/ 4 digits right of decimal pt)
        /// </summary>
        PT_CURRENCY = 6, 
        /// <summary>
        /// Application time
        /// </summary>
        PT_APPTIME = 7,
        /// <summary>
        /// 32-bit error value
        /// </summary>
        PT_ERROR = 10,
        /// <summary>
        /// 16-bit boolean (non-zero true,zero false)
        /// </summary>
        PT_BOOLEAN = 11, 
        /// <summary>
        /// 16-bit boolean (non-zero true)
        /// </summary>
        PT_BOOLEAN_DESKTOP = 11,
        /// <summary>
        /// Embedded object in a property
        /// </summary>
        PT_OBJECT = 13,
        /// <summary>
        /// 8-byte signed integer
        /// </summary>
        PT_I8 = 20, 
        /// <summary>
        /// Null terminated 8-bit character string
        /// </summary>
        PT_STRING8 = 30,
        /// <summary>
        /// Null terminated Unicode string
        /// </summary>
        PT_UNICODE = 31,
        /// <summary>
        /// FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601
        /// </summary>
        PT_SYSTIME = 64,
        /// <summary>
        /// OLE GUID
        /// </summary>
        PT_CLSID = 72,
        /// <summary>
        /// Uninterpreted (counted byte array)
        /// </summary>
        PT_BINARY = 258,
        /// <summary>
        /// IF MAPI Unicode, PT_TString is Unicode string; otheriwse 8-bit character string
        /// </summary>
        PT_TSTRING = PT_UNICODE
    }
}
