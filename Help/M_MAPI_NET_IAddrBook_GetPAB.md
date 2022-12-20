# GetPAB Method


Returns the entry identifier of the container that is designated as the personal address book (PAB).



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetPAB(
	out uint lpcbEntryID,
	out IntPtr lppEntryID
)
```
**VB**
``` VB
Function GetPAB ( 
	<OutAttribute> ByRef lpcbEntryID As UInteger,
	<OutAttribute> ByRef lppEntryID As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetPAB(
	[OutAttribute] unsigned int% lpcbEntryID, 
	[OutAttribute] IntPtr% lppEntryID
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the entry identifier of the PAB.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the entry identifier of the PAB was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
