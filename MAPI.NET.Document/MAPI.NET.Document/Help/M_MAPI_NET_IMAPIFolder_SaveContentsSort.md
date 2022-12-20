# SaveContentsSort Method


Sets the default sort order for a folder's contents table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SaveContentsSort(
	IntPtr lpSortCriteria,
	uint ulFlags
)
```
**VB**
``` VB
Function SaveContentsSort ( 
	lpSortCriteria As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT SaveContentsSort(
	IntPtr lpSortCriteria, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>] A pointer to an SSortOrderSet structure that contains the default sort order.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the default sort order is set.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the sort order was successfully saved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
