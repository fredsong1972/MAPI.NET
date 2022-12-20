# IMAPIProp Interface


Enables clients, service providers, and MAPI to work with properties. All objects that support properties implement this interface.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
[GuidAttribute("00020303-0000-0000-C000-000000000046")]
public interface IMAPIProp
```
**VB**
``` VB
<ComImportAttribute>
<ComVisibleAttribute(false)>
<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
<GuidAttribute("00020303-0000-0000-C000-000000000046")>
Public Interface IMAPIProp
```
**C++**
``` C++
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIUnknown)]
[GuidAttribute(L"00020303-0000-0000-C000-000000000046")]
public interface class IMAPIProp
```



## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_CopyProps.md">CopyProps</a></td>
<td>Copies or moves selected properties.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_CopyTo.md">CopyTo</a></td>
<td>Copies or moves all properties, except for specifically excluded properties.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_DeleteProps.md">DeleteProps</a></td>
<td>Deletes one or more properties from an object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetIDsFromNames.md">GetIDsFromNames</a></td>
<td>Provides the property identifiers that correspond to one or more property names.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetLastError.md">GetLastError</a></td>
<td>Returns a MAPIERROR structure that contains information about the previous error.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetNamesFromIDs.md">GetNamesFromIDs</a></td>
<td>Provides the property names that correspond to one or more property identifiers.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetPropList.md">GetPropList</a></td>
<td>Returns property tags for all properties.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_GetProps.md">GetProps</a></td>
<td>Retrieves the property value of one or more properties of an object.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_OpenProperty.md">OpenProperty</a></td>
<td>Returns a pointer to an interface that can be used to access a property.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SaveChanges.md">SaveChanges</a></td>
<td>Makes permanent any changes that were made to an object since the last save operation.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPIProp_SetProps.md">SetProps</a></td>
<td>Updates one or more properties.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
