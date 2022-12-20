# SortRows Method


Orders the rows of the table, depending on sort criteria.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SortRows(
	PropTags tag,
	TableSortOrder order
)
```
**VB**
``` VB
Public Function SortRows ( 
	tag As PropTags,
	order As TableSortOrder
) As Boolean
```
**C++**
``` C++
public:
bool SortRows(
	PropTags tag, 
	TableSortOrder order
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_PropTags.md">PropTags</a></dt><dd>Property tag identifying the sort key.</dd><dt>  <a href="T_MAPI_NET_TableSortOrder.md">TableSortOrder</a></dt><dd>The order in which the data is to be sorted.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the sort operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPITable.md">MAPITable Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
