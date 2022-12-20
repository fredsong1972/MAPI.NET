# CreateOneOff Method


Creates an entry identifier for a one-off address.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CreateOneOff(
	string lpszName,
	string lpszAdrType,
	string lpszAddress,
	uint ulFlags,
	out uint lpcbEntryID,
	out IntPtr lppEntryID
)
```
**VB**
``` VB
Function CreateOneOff ( 
	lpszName As String,
	lpszAdrType As String,
	lpszAddress As String,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpcbEntryID As UInteger,
	<OutAttribute> ByRef lppEntryID As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT CreateOneOff(
	String^ lpszName, 
	String^ lpszAdrType, 
	String^ lpszAddress, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpcbEntryID, 
	[OutAttribute] IntPtr% lppEntryID
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The recipient’s display name.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The address type of the recipient</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The address of the recipient</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that affects the one-off recipient.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the entry identifier for the one-off recipient.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the one-off entry identifier was created successfully; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
