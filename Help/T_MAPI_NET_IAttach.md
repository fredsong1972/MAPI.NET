# IAttach Interface


The IAttach interface is used to maintain and provide access to the properties of attachments in messages.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[ComImportAttribute]
[ComVisibleAttribute(false)]
[GuidAttribute("00020308-0000-0000-C000-000000000046")]
[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
public interface IAttach : IMAPIProp
```
**VB**
``` VB
<ComImportAttribute>
<ComVisibleAttribute(false)>
<GuidAttribute("00020308-0000-0000-C000-000000000046")>
<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IAttach
	Inherits IMAPIProp
```
**C++**
``` C++
[ComImportAttribute]
[ComVisibleAttribute(false)]
[GuidAttribute(L"00020308-0000-0000-C000-000000000046")]
[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIUnknown)]
public interface class IAttach : IMAPIProp
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
<td><a href="M_MAPI_NET_IMAPIProp_DeleteProps.md">DeleteProps</a></td>
<td>Deletes one or more properties from an object.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
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
<td><a href="M_MAPI_NET_IMAPIProp_OpenProperty.md">OpenProperty</a></td>
<td>Returns a pointer to an interface that can be used to access a property.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SaveChanges.md">SaveChanges</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SetProps.md">SetProps</a></td>
<td>Updates one or more properties.<br />(Inherited from <a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>)</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
