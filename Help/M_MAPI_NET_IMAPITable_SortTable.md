# SortTable Method


Orders the rows of the table, depending on sort criteria.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SortTable(
	IntPtr lpSortCriteria,
	int ulFlags
)
```
**VB**
``` VB
Function SortTable ( 
	lpSortCriteria As IntPtr,
	ulFlags As Integer
) As HRESULT
```
**C++**
``` C++
HRESULT SortTable(
	IntPtr lpSortCriteria, 
	int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to an SSortOrderSet structure that contains the sort criteria to apply.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>Bitmask of flags that controls the timing of the IMAPITable::SortTable operation.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the sort operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
