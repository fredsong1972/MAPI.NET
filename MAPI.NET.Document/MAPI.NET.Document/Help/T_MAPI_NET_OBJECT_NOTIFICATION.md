# OBJECT_NOTIFICATION Structure


Contains information about an object that has undergone a change, such as being copied or modified.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public struct OBJECT_NOTIFICATION
```
**VB**
``` VB
Public Structure OBJECT_NOTIFICATION
```
**C++**
``` C++
public value class OBJECT_NOTIFICATION
```

<table><tr><td><strong>Inheritance</strong></td><td><a href="https://learn.microsoft.com/dotnet/api/system.object" target="_blank" rel="noopener noreferrer">Object</a>  →  <a href="https://learn.microsoft.com/dotnet/api/system.valuetype" target="_blank" rel="noopener noreferrer">ValueType</a>  →  OBJECT_NOTIFICATION</td></tr>
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
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_cbEntryID.md">cbEntryID</a></td>
<td>Count of bytes in the entry identifier.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_cbOldID.md">cbOldID</a></td>
<td>Count of bytes in the entry identifier pointed to by the pOldID member.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_cbOldParentID.md">cbOldParentID</a></td>
<td>Count of bytes in the entry identifier pointed to by the pOldParentID member.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_cbParentID.md">cbParentID</a></td>
<td>Count of bytes in the entry identifier pointed to by the pParentID member.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_ObjType.md">ObjType</a></td>
<td>Type of object affected.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_pEntryID.md">pEntryID</a></td>
<td>Pointer to the entry identifier of the affected object.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_pOldID.md">pOldID</a></td>
<td>Pointer to the entry identifier of the original object. This pointer can be NULL if the event does not require an original object.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_pOldParentID.md">pOldParentID</a></td>
<td>Pointer to the entry identifier of the parent of the original object. This pointer can be NULL if the event does not require an original object.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_pParentID.md">pParentID</a></td>
<td>Pointer to the entry identifier of the parent of the affected object.</td></tr>
<tr>
<td><a href="F_MAPI_NET_OBJECT_NOTIFICATION_PropTagArray.md">PropTagArray</a></td>
<td>Pointer to an SPropTagArray structure that contains the property tags identifying properties affected by the event.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
