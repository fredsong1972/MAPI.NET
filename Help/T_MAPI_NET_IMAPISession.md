# IMAPISession Interface


Manages objects associated with a MAPI logon session.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
[GuidAttribute("00020300-0000-0000-C000-000000000046")]
public interface IMAPISession
```
**VB**
``` VB
<ComImportAttribute>
<ComVisibleAttribute(false)>
<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
<GuidAttribute("00020300-0000-0000-C000-000000000046")>
Public Interface IMAPISession
```
**C++**
``` C++
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIUnknown)]
[GuidAttribute(L"00020300-0000-0000-C000-000000000046")]
public interface class IMAPISession
```



## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_AdminServices.md">AdminServices</a></td>
<td>Returns an IMsgServiceAdmin pointer for making changes to message services.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_Advise.md">Advise</a></td>
<td>Registers to receive notification of specified events that affect the session.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_CompareEntryIDs.md">CompareEntryIDs</a></td>
<td>Compares two entry identifiers to determine whether they refer to the same entry in a message store.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_GetLastError.md">GetLastError</a></td>
<td>Returns a MAPIERROR structure that contains information about the previous session error.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_GetMsgStoresTable.md">GetMsgStoresTable</a></td>
<td>Provides access to the message store table that contains information about all the message stores in the session profile.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_GetStatusTable.md">GetStatusTable</a></td>
<td>Provides access to the status table, a table that contains information about all the MAPI resources in the session.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_Logoff.md">Logoff</a></td>
<td>Ends a MAPI session.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_OpenAddressBook.md">OpenAddressBook</a></td>
<td>Opens the MAPI integrated address book, returning an IAddrBook pointer for further access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_OpenEntry.md">OpenEntry</a></td>
<td>Opens an object and returns an interface pointer for additional access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_OpenMsgStore.md">OpenMsgStore</a></td>
<td>Opens a message store and returns an IMsgStore pointer for further access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_OpenProfileSection.md">OpenProfileSection</a></td>
<td>Opens a section of the current profile and returns an IProfSect pointer for further access.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_PrepareForm.md">PrepareForm</a></td>
<td>Creates a numeric token that the IMAPISession::ShowForm method uses to access a message.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_QueryIdentity.md">QueryIdentity</a></td>
<td>Returns the entry identifier of the object that provides the primary identity for the session.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_SetDefaultStore.md">SetDefaultStore</a></td>
<td>Establishes a message store as the default message store for the session.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_ShowForm.md">ShowForm</a></td>
<td>Displays a form.</td></tr>
<tr>
<td><a href="M_MAPI_NET_IMAPISession_Unadvise.md">Unadvise</a></td>
<td>Cancels the sending of notifications previously set up with a call to the IMAPISession::Advise method.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
