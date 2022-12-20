# NEWMAIL_NOTIFICATION Structure


Describes information that relate to the arrival of a new message.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public struct NEWMAIL_NOTIFICATION
```
**VB**
``` VB
Public Structure NEWMAIL_NOTIFICATION
```
**C++**
``` C++
public value class NEWMAIL_NOTIFICATION
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  <a href="https://learn.microsoft.com/dotnet/api/system.valuetype" target="_blank" rel="noopener noreferrer">ValueType</a>  →  NEWMAIL_NOTIFICATION</td></tr>
</table>



## Methods
<table>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.valuetype.equals#system-valuetype-equals(system-object)" target="_blank" rel="noopener noreferrer">Equals</a></td>
<td>Indicates whether this instance and a specified object are equal.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.valuetype" target="_blank" rel="noopener noreferrer">ValueType</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.valuetype.gethashcode#system-valuetype-gethashcode" target="_blank" rel="noopener noreferrer">GetHashCode</a></td>
<td>Returns the hash code for this instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.valuetype" target="_blank" rel="noopener noreferrer">ValueType</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.object.gettype#system-object-gettype" target="_blank" rel="noopener noreferrer">GetType</a></td>
<td>Gets the <a href="https://learn.microsoft.com/dotnet/api/system.type" target="_blank" rel="noopener noreferrer">Type</a> of the current instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>)</td></tr>
<tr>
<td><a href="https://learn.microsoft.com/dotnet/api/system.valuetype.tostring#system-valuetype-tostring" target="_blank" rel="noopener noreferrer">ToString</a></td>
<td>Returns the fully qualified type name of this instance.<br />(Inherited from <a href="https://learn.microsoft.com/dotnet/api/system.valuetype" target="_blank" rel="noopener noreferrer">ValueType</a>)</td></tr>
</table>

## Fields
<table>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_cbEntryID.md">cbEntryID</a></td>
<td>Count of bytes in the entry identifier pointed to by the lpEntryID member.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_cbParentID.md">cbParentID</a></td>
<td>Count of bytes in the entry identifier pointed to by the lpParentID member.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_Flags.md">Flags</a></td>
<td>Bitmask of flags used to describe the format of the string properties included with the message.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_MessageClass.md">MessageClass</a></td>
<td>The message class of the newly arrived message.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_MessageFlags.md">MessageFlags</a></td>
<td>Bitmask of flags that describes the current state of the newly arrived message.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_pEntryID.md">pEntryID</a></td>
<td>Pointer to the entry identifier of the newly arrived message.</td></tr>
<tr>
<td><a href="F_MAPI_NET_NEWMAIL_NOTIFICATION_pParentID.md">pParentID</a></td>
<td>Pointer to the entry identifier of the receive folder for the newly arrived message.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
