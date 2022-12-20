# MsgStore Class


IMsgStore .Net Wrapper object



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public class MsgStore : MAPIObject
```
**VB**
``` VB
Public Class MsgStore
	Inherits MAPIObject
```
**C++**
``` C++
public ref class MsgStore : public MAPIObject
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>  →  MsgStore</td></tr>
</table>



## Constructors
<table>
<tr>
<td><a href="M_MAPI_NET_MsgStore__ctor.md">MsgStore</a></td>
<td>Initializes a new instance of the MsgStore class.</td></tr>
</table>

## Properties
<table>
<tr>
<td><a href="P_MAPI_NET_MAPIObject_EntryID.md">EntryID</a></td>
<td>EntryID of MAPIObject<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MsgStore_InboxEntryID.md">InboxEntryID</a></td>
<td>Get Inbox entry id</td></tr>
<tr>
<td><a href="P_MAPI_NET_MsgStore_InterfaceID.md">InterfaceID</a></td>
<td>Gets IMsgStore interface GUID<br />(Overrides <a href="P_MAPI_NET_MAPIObject_InterfaceID.md">MAPIObject.InterfaceID</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MsgStore_Session.md">Session</a></td>
<td>Get MAPI session</td></tr>
<tr>
<td><a href="P_MAPI_NET_MsgStore_StoreID.md">StoreID</a></td>
<td>Gets entry identificatio of Message Store</td></tr>
</table>

## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Close.md">Close</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and close the object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_CopyTo.md">CopyTo(UInt32[], IMAPIProp)</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_CopyTo_1.md">CopyTo(UInt32[], MAPIObject)</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_CreateAppointment.md">CreateAppointment</a></td>
<td>Create an appointnment</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_CreateContact.md">CreateContact</a></td>
<td>Create a contact</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_CreateMessage.md">CreateMessage</a></td>
<td>Create a message</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_Dispose.md">Dispose</a></td>
<td>Dispose Msgstore object<br />(Overrides <a href="M_MAPI_NET_MAPIObject_Dispose.md">MAPIObject.Dispose()</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.equals#system-object-equals(system-object)" target="_blank" rel="noopener noreferrer">Equals</a></td>
<td>Determines whether the specified object is equal to the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.finalize#system-object-finalize" target="_blank" rel="noopener noreferrer">Finalize</a></td>
<td>Allows an object to try to free resources and perform other cleanup operations before it is reclaimed by garbage collection.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_ForwardMessage.md">ForwardMessage</a></td>
<td>Forward message</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gethashcode#system-object-gethashcode" target="_blank" rel="noopener noreferrer">GetHashCode</a></td>
<td>Serves as the default hash function.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_GetMAPIMessage.md">GetMAPIMessage</a></td>
<td>Gets MAPI Message per the entry identification</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_GetMAPIObject.md">GetMAPIObject</a></td>
<td>Gets MAPI object per the entry identification</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookProperties.md">GetOutlookProperties</a></td>
<td>Get outlook properties<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookProperty.md">GetOutlookProperty</a></td>
<td>Get outlook properties<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookPropTag.md">GetOutlookPropTag</a></td>
<td>Get one outlook property tag<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetOutlookPropTags.md">GetOutlookPropTags</a></td>
<td>Get outlook property tags<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperties.md">GetProperties(PropTags[])</a></td>
<td>Retrieves the property value of one or more properties of an object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperties_1.md">GetProperties(UInt32[])</a></td>
<td>Retrieves the property value of one or more properties of an object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperty.md">GetProperty(PropTags)</a></td>
<td>Retrieves one property value of an object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_GetProperty_1.md">GetProperty(UInt32)</a></td>
<td>Retrieves one property value of an object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gettype#system-object-gettype" target="_blank" rel="noopener noreferrer">GetType</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.type" target="_blank" rel="noopener noreferrer">Type</a> of the current instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.memberwiseclone#system-object-memberwiseclone" target="_blank" rel="noopener noreferrer">MemberwiseClone</a></td>
<td>Creates a shallow copy of the current <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenCalendar.md">OpenCalendar</a></td>
<td>Open calendar folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenContacts.md">OpenContacts</a></td>
<td>Open contact folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenDeletedItems.md">OpenDeletedItems</a></td>
<td>Open deleted items folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenDraft.md">OpenDraft</a></td>
<td>Open draft folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenEntry.md">OpenEntry</a></td>
<td>Opens an object and returns an interface pointer for additional access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenFolder.md">OpenFolder</a></td>
<td>Open a MAPI folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenInbox.md">OpenInbox</a></td>
<td>Open inbox</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenJunkFolder.md">OpenJunkFolder</a></td>
<td>Open junk folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenOutbox.md">OpenOutbox</a></td>
<td>Open outbox</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty.md">OpenProperty(UInt32, Guid, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty_1.md">OpenProperty(UInt32, Guid, UInt32, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenRootFolder.md">OpenRootFolder</a></td>
<td>Open root folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_OpenSentItems.md">OpenSentItems</a></td>
<td>Open Sent folder</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_RegisterEvents.md">RegisterEvents</a></td>
<td>Registers to receive notification of specified events that affect the message store.</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save.md">Save()</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and keep the object open.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save_1.md">Save(SaveFlags)</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and control the object per the flag.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_SendMessage.md">SendMessage</a></td>
<td>Send message</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue.md">SetPropertyValue(IPropValue)</a></td>
<td>Updates one property of the object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue_1.md">SetPropertyValue(PropTags, Object)</a></td>
<td>Updates one property of the object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValue_2.md">SetPropertyValue(UInt32, Object)</a></td>
<td>Updates one property of the object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_SetPropertyValues.md">SetPropertyValues</a></td>
<td>Updates one or more properties of the object.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.tostring#system-object-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns a string that represents the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MsgStore_UnRegisteEvents.md">UnRegisteEvents</a></td>
<td>Cancels the sending of notifications.</td></tr>
</table>

## Events
<table>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnNewMail.md">OnNewMail</a></td>
<td>New mail event</td></tr>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnObjectCopied.md">OnObjectCopied</a></td>
<td>Object copied event</td></tr>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnObjectCreated.md">OnObjectCreated</a></td>
<td>Object created event</td></tr>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnObjectDeleted.md">OnObjectDeleted</a></td>
<td>Object deleted event</td></tr>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnObjectModified.md">OnObjectModified</a></td>
<td>Object modified event</td></tr>
<tr>
<td><a href="E_MAPI_NET_MsgStore_OnObjectMoved.md">OnObjectMoved</a></td>
<td>Object moved event</td></tr>
</table>

## Fields
<table>
<tr>
<td><a href="F_MAPI_NET_MAPIObject_Id_.md">Id_</a></td>
<td>EntryID of the object<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
