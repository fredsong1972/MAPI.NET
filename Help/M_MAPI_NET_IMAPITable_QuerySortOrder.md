# QuerySortOrder Method


Retrieves the current sort order for a table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT QuerySortOrder(
	IntPtr lppSortCriteria
)
```
**VB**
``` VB
Function QuerySortOrder ( 
	lppSortCriteria As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT QuerySortOrder(
	IntPtr lppSortCriteria
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer to the SSortOrderSet structure holding the current sort order.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the current sort order was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
