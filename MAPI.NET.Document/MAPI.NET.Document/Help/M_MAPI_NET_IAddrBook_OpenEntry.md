# OpenEntry Method


Opens an address book entry and returns a pointer to an interface that can be used to access the entry.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT OpenEntry(
	uint cbEntryID,
	IntPtr lpEntryID,
	IntPtr lpInterface,
	uint ulFlags,
	out uint lpulObjType,
	out IntPtr lppUnk
)
```
**VB**
``` VB
Function OpenEntry ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulObjType As UInteger,
	<OutAttribute> ByRef lppUnk As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OpenEntry(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpulObjType, 
	[OutAttribute] IntPtr% lppUnk
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier that represents the address book entry to open.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) of the interface to be used to access the open entry.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the entry is opened. The following flags can be set.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the type of the opened entry.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the opened entry.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the entry was successfully opened;otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
