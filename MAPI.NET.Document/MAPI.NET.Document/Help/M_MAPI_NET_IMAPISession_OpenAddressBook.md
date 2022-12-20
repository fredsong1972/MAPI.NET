# OpenAddressBook Method


Opens the MAPI integrated address book, returning an IAddrBook pointer for further access.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT OpenAddressBook(
	uint ulUIParam,
	IntPtr lpInterface,
	uint ulFlags,
	out IntPtr lppAdrBook
)
```
**VB**
``` VB
Function OpenAddressBook ( 
	ulUIParam As UInteger,
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppAdrBook As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OpenAddressBook(
	unsigned int ulUIParam, 
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppAdrBook
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the common address dialog box and other related displays.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the address book.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the opening of the address book.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the address book.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the address book was successfully opened; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
