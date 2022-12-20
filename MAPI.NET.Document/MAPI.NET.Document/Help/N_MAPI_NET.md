# MAPI.NET Namespace


The MAPI.NET namespace contains interfaces and classes that define MAPI wrapper objects , which allow users to use MAPI functions from .Net level.



## Classes
<table>
<tr>
<td><a href="T_MAPI_NET_Appointment.md">Appointment</a></td>
<td>The Appointment class is used to maintain and provide access to the properties of appointments.</td></tr>
<tr>
<td><a href="T_MAPI_NET_Attachment.md">Attachment</a></td>
<td>Represents an attachment to an e-mail.</td></tr>
<tr>
<td><a href="T_MAPI_NET_Contact.md">Contact</a></td>
<td>The Contact class is used to maintain and provide access to the properties of contacts.</td></tr>
<tr>
<td><a href="T_MAPI_NET_EntryID.md">EntryID</a></td>
<td>Represents an entry identifier for a MAPI object.</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIAddressBook.md">MAPIAddressBook</a></td>
<td>IAddrBook .Net Wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIFolder.md">MAPIFolder</a></td>
<td>IMAPIFolder .Net Wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a></td>
<td>IMessage .Net Wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a></td>
<td>IMAPIProp .Net Wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIProp.md">MAPIProp</a></td>
<td>Represents a property of various MAPI objects.</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPISession.md">MAPISession</a></td>
<td>IMAPISession .Net wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIStatus.md">MAPIStatus</a></td>
<td>IMAPIStatus .Net wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPITable.md">MAPITable</a></td>
<td>.Net wrapper over IMAPITable interface.</td></tr>
<tr>
<td><a href="T_MAPI_NET_MsgStore.md">MsgStore</a></td>
<td>IMsgStore .Net Wrapper object</td></tr>
<tr>
<td><a href="T_MAPI_NET_MsgStoreNewMailEventArgs.md">MsgStoreNewMailEventArgs</a></td>
<td>Provides data for the arrival of a new message.</td></tr>
<tr>
<td><a href="T_MAPI_NET_MsgStoreObjectEventArgs.md">MsgStoreObjectEventArgs</a></td>
<td>Provides data for an object that has undergone a change, such as being copied or modified.</td></tr>
<tr>
<td><a href="T_MAPI_NET_Recipient.md">Recipient</a></td>
<td>Defines a Recipient for MAPI messages.</td></tr>
<tr>
<td><a href="T_MAPI_NET_RtfHtmlConverter.md">RtfHtmlConverter</a></td>
<td>Rtf Html Converter</td></tr>
</table>

## Structures
<table>
<tr>
<td><a href="T_MAPI_NET_ADRENTRY.md">ADRENTRY</a></td>
<td>Describes zero or more properties that belong to a recipient.</td></tr>
<tr>
<td><a href="T_MAPI_NET_ADRLIST.md">ADRLIST</a></td>
<td>An ADRLIST structure contains one or more ADRENTRY structures, each describing the properties of a recipient.</td></tr>
<tr>
<td><a href="T_MAPI_NET_NEWMAIL_NOTIFICATION.md">NEWMAIL_NOTIFICATION</a></td>
<td>Describes information that relate to the arrival of a new message.</td></tr>
<tr>
<td><a href="T_MAPI_NET_OBJECT_NOTIFICATION.md">OBJECT_NOTIFICATION</a></td>
<td>Contains information about an object that has undergone a change, such as being copied or modified.</td></tr>
<tr>
<td><a href="T_MAPI_NET_SRow.md">SRow</a></td>
<td>Describes a row from a table containing selected properties for a specific object.</td></tr>
<tr>
<td><a href="T_MAPI_NET_SSortOrder.md">SSortOrder</a></td>
<td>Defines how to sort rows of a table, describing the column to use as the sort key and the direction of the sort.</td></tr>
</table>

## Interfaces
<table>
<tr>
<td><a href="T_MAPI_NET_IAddrBook.md">IAddrBook</a></td>
<td>Supports access to the MAPI address book and includes operations such as displaying common dialog boxes; opening containers, messaging users, and distribution lists; and performing name resolution.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IAttach.md">IAttach</a></td>
<td>The IAttach interface is used to maintain and provide access to the properties of attachments in messages.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPIAdviseSink.md">IMAPIAdviseSink</a></td>
<td>The IMAPIAdviseSink interface is used to implement an Advise Sink object for handling notifications.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer</a></td>
<td>Manages high-level operations on container objects such as address books, distribution lists, and folders.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder</a></td>
<td>PerPerforms operations on the messages and subfolders in a folder.forms operations on the messages and subfolders in a folder.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a></td>
<td>Enables clients, service providers, and MAPI to work with properties. All objects that support properties implement this interface.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPISession.md">IMAPISession</a></td>
<td>Manages objects associated with a MAPI logon session.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPIStatus.md">IMAPIStatus</a></td>
<td>Provides status information about the MAPI subsystem, the integrated address book, and the MAPI spooler.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMAPITable.md">IMAPITable</a></td>
<td>Provides a read-only view of a table. IMAPITable is used by clients and service providers to manipulate the way a table appears.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMessage.md">IMessage</a></td>
<td>Manages messages, attachments, and recipients.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IMsgStore.md">IMsgStore</a></td>
<td>Provides access to message store information and to messages and folders.</td></tr>
<tr>
<td><a href="T_MAPI_NET_IPropValue.md">IPropValue</a></td>
<td>Provides an interface to manage MAPI property value</td></tr>
</table>

## Enumerations
<table>
<tr>
<td><a href="T_MAPI_NET_AddressType.md">AddressType</a></td>
<td>Recipient address type enumeration</td></tr>
<tr>
<td><a href="T_MAPI_NET_AttachMethod.md">AttachMethod</a></td>
<td>AttachMehod contains a MAPI-defined constant representing the way the contents of an attachment can be accessed.</td></tr>
<tr>
<td><a href="T_MAPI_NET_BookMark.md">BookMark</a></td>
<td>The bookmark identifying the starting position for the seek operation.</td></tr>
<tr>
<td><a href="T_MAPI_NET_CharacterSet.md">CharacterSet</a></td>
<td>Bitmask of flags that controls the mapi character set.</td></tr>
<tr>
<td><a href="T_MAPI_NET_EEventMask.md">EEventMask</a></td>
<td>A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration. There is a corresponding NOTIFICATION structure associated with each type of event that holds information about the event.</td></tr>
<tr>
<td><a href="T_MAPI_NET_FlushFlag.md">FlushFlag</a></td>
<td>Bitmask of flags that controls the flush operation.</td></tr>
<tr>
<td><a href="T_MAPI_NET_HRESULT.md">HRESULT</a></td>
<td>The HRESULT data type is a 32-bit value is used to describe an error or warning.</td></tr>
<tr>
<td><a href="T_MAPI_NET_Importance.md">Importance</a></td>
<td>Importance: low, normal, high</td></tr>
<tr>
<td><a href="T_MAPI_NET_MAPIFlag.md">MAPIFlag</a></td>
<td>Generic MAPI Bitmask</td></tr>
<tr>
<td><a href="T_MAPI_NET_MessageFormat.md">MessageFormat</a></td>
<td>Message Format</td></tr>
<tr>
<td><a href="T_MAPI_NET_ModifyRecipientFlag.md">ModifyRecipientFlag</a></td>
<td>A bitmask of flags that controls recipient</td></tr>
<tr>
<td><a href="T_MAPI_NET_ObjectType.md">ObjectType</a></td>
<td>Type of MAPI object</td></tr>
<tr>
<td><a href="T_MAPI_NET_OpenPropertyMode.md">OpenPropertyMode</a></td>
<td>A bitmask of flags that controls access to the property.</td></tr>
<tr>
<td><a href="T_MAPI_NET_PropTags.md">PropTags</a></td>
<td>MAPI Property Tag</td></tr>
<tr>
<td><a href="T_MAPI_NET_PT.md">PT</a></td>
<td>Property Type</td></tr>
<tr>
<td><a href="T_MAPI_NET_RecipientType.md">RecipientType</a></td>
<td>Contains the recipient type for a message recipient.</td></tr>
<tr>
<td><a href="T_MAPI_NET_SaveFlags.md">SaveFlags</a></td>
<td>A bitmask of flags that controls what happens to the object when the IMAPIProp::SaveChanges method is called.</td></tr>
<tr>
<td><a href="T_MAPI_NET_Sensitivity.md">Sensitivity</a></td>
<td>Sensitivity: none, personal, private, confidential</td></tr>
<tr>
<td><a href="T_MAPI_NET_TableSortOrder.md">TableSortOrder</a></td>
<td>Defines how to sort the rows of a table, what column to use as the sort key, and the direction of the sort.</td></tr>
</table>