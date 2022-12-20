# OBJECT_NOTIFICATION Fields




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
<a href="T_MAPI_NET_OBJECT_NOTIFICATION.md">OBJECT_NOTIFICATION Structure</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
