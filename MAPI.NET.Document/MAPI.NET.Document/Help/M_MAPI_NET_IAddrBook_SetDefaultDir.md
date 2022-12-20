# SetDefaultDir Method


Establishes the specified container as the default address book container.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetDefaultDir(
	uint cbEntryID,
	IntPtr lpEntryID
)
```
**VB**
``` VB
Function SetDefaultDir ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SetDefaultDir(
	unsigned int cbEntryID, 
	IntPtr lpEntryID
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the default address book container.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the default address book container was successfully set; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
