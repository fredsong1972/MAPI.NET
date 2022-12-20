# IMessage Interface


Manages messages, attachments, and recipients.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[ComImportAttribute]
[ComVisibleAttribute(false)]
[GuidAttribute("00020307-0000-0000-C000-000000000046")]
[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
public interface IMessage : IMAPIProp
```
**VB**
``` VB
<ComImportAttribute>
<ComVisibleAttribute(false)>
<GuidAttribute("00020307-0000-0000-C000-000000000046")>
<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IMessage
	Inherits IMAPIProp
```
**C++**
``` C++
[ComImportAttribute]
[ComVisibleAttribute(false)]
[GuidAttribute(L"00020307-0000-0000-C000-000000000046")]
[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIUnknown)]
public interface class IMessage : IMAPIProp
```

<table><tr><td><strong>Implements</strong></td><td><a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a></td></tr>
</table>



## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_CopyProps.md">CopyProps</a></td>
<td>Copies or moves selected properties.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_CopyTo.md">CopyTo</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_CreateAttach.md">CreateAttach</a></td>
<td>Creates a new attachment.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_DeleteAttach.md">DeleteAttach</a></td>
<td>Deletes an attachment.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_DeleteProps.md">DeleteProps</a></td>
<td>Deletes one or more properties from an object.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_GetAttachmentTable.md">GetAttachmentTable</a></td>
<td>Returns the message's attachment table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetIDsFromNames.md">GetIDsFromNames</a></td>
<td>Provides the property identifiers that correspond to one or more property names.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetLastError.md">GetLastError</a></td>
<td>Returns a MAPIERROR structure that contains information about the previous error.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetNamesFromIDs.md">GetNamesFromIDs</a></td>
<td>Provides the property names that correspond to one or more property identifiers.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetPropList.md">GetPropList</a></td>
<td>Returns property tags for all properties.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetProps.md">GetProps</a></td>
<td>Retrieves the property value of one or more properties of an object.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_GetRecipientTable.md">GetRecipientTable</a></td>
<td>Returns the message's recipient table.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_ModifyRecipients.md">ModifyRecipients</a></td>
<td>Adds, deletes, or modifies message recipients.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_OpenAttach.md">OpenAttach</a></td>
<td>Opens an attachment.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_OpenProperty.md">OpenProperty</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SaveChanges.md">SaveChanges</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SetProps.md">SetProps</a></td>
<td>Updates one or more properties.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_SetReadFlag.md">SetReadFlag</a></td>
<td>Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of the message and manages the sending of read reports.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMessage_SubmitMessage.md">SubmitMessage</a></td>
<td>Saves all of the message's properties and marks the message as ready to be sent.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
