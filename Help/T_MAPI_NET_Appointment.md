# Appointment Class


The Appointment class is used to maintain and provide access to the properties of appointments.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public class Appointment : MAPIMessage
```
**VB**
``` VB
Public Class Appointment
	Inherits MAPIMessage
```
**C++**
``` C++
public ref class Appointment : public MAPIMessage
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>  →  <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>  →  Appointment</td></tr>
</table>



## Properties
<table>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_AttachmentCount.md">AttachmentCount</a></td>
<td>Gets the count of attachmentsof the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Body.md">Body</a></td>
<td>Gets or sets Body of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_CreationTime.md">CreationTime</a></td>
<td>Gets/sets a Date indicating the date and time at which the item was created.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIObject_EntryID.md">EntryID</a></td>
<td>EntryID of MAPIObject<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Importance.md">Importance</a></td>
<td>Gets/sets Importance<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_InterfaceID.md">InterfaceID</a></td>
<td>IMessage inteface GUID<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_LastModifiedTime.md">LastModifiedTime</a></td>
<td>Gets the last modified time<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_MessageClass.md">MessageClass</a></td>
<td>Gets MessageClass property value<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_MessageFlags.md">MessageFlags</a></td>
<td>Gets MessageFlags property value<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_MessageFormat.md">MessageFormat</a></td>
<td>Gets/sets message format<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Priority.md">Priority</a></td>
<td>Gets/sets Priority<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_ReceivedTime.md">ReceivedTime</a></td>
<td>Returns a Date indicating the date and time at which the item was received. Read-only<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Recipients.md">Recipients</a></td>
<td>Gets a list of recipients of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Sender.md">Sender</a></td>
<td>Gets Sender of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Sensitivity.md">Sensitivity</a></td>
<td>Gets/Sets the sensitivity of the contact<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Subject.md">Subject</a></td>
<td>Gets or sets Subject property value<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_SubmitTime.md">SubmitTime</a></td>
<td>Gets/Sets SubmitTime<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="P_MAPI_NET_MAPIMessage_Unread.md">Unread</a></td>
<td>Returns or sets a Boolean value that is True if the item has not been opened (read).<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
</table>

## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_AddAttachment.md">AddAttachment(MAPIMessage, String)</a></td>
<td>Add a MAPI Message as an embedded attachment<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_AddAttachment_1.md">AddAttachment(String, String)</a></td>
<td>Add a file to an attachment of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_AddRecipient.md">AddRecipient(String)</a></td>
<td>Add recipient<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_AddRecipient_1.md">AddRecipient(String, MAPIAddressBook)</a></td>
<td>Add a recipient<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_AddRecipient_2.md">AddRecipient(String, RecipientType, AddressType, MAPIAddressBook)</a></td>
<td>Add a recipient<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
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
<td><a href="M_MAPI_NET_MAPIMessage_Dispose.md">Dispose</a></td>
<td>Dispose MAPIMessage object<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.equals#system-object-equals(system-object)" target="_blank" rel="noopener noreferrer">Equals</a></td>
<td>Determines whether the specified object is equal to the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.finalize#system-object-finalize" target="_blank" rel="noopener noreferrer">Finalize</a></td>
<td>Allows an object to try to free resources and perform other cleanup operations before it is reclaimed by garbage collection.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_GetAttachmentName.md">GetAttachmentName</a></td>
<td>Get an attachment display name<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gethashcode#system-object-gethashcode" target="_blank" rel="noopener noreferrer">GetHashCode</a></td>
<td>Serves as the default hash function.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
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
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty.md">OpenProperty(UInt32, Guid, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_OpenProperty_1.md">OpenProperty(UInt32, Guid, UInt32, OpenPropertyMode)</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_Refresh.md">Refresh</a></td>
<td>Refresh the cached property values.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save.md">Save()</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and keep the object open.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIObject_Save_1.md">Save(SaveFlags)</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation,and control the object per the flag.<br />(Inherited from <a href="T_MAPI_NET_MAPIObject.md">MAPIObject</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveAttachment.md">SaveAttachment(Int32)</a></td>
<td>Save attachment to a disk file.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveAttachment_1.md">SaveAttachment(String, Int32)</a></td>
<td>Save attachment to a disk file.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveAttachment_2.md">SaveAttachment(String, Int32, String)</a></td>
<td>Save attachment to a disk file.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveToMsg.md">SaveToMsg()</a></td>
<td>Save the message to MSG file.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveToMsg_1.md">SaveToMsg(String)</a></td>
<td>Save the message to a MSG file.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SaveToSentFolder.md">SaveToSentFolder</a></td>
<td>Set if save a message to the sent folder after sent.<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_Send.md">Send</a></td>
<td>Send the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
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
<td><a href="M_MAPI_NET_MAPIMessage_SetSender.md">SetSender(EntryID)</a></td>
<td>Set sender of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_MAPIMessage_SetSender_1.md">SetSender(String, String, MAPIAddressBook)</a></td>
<td>Set sender of the message<br />(Inherited from <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.tostring#system-object-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns a string that represents the current object.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
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
