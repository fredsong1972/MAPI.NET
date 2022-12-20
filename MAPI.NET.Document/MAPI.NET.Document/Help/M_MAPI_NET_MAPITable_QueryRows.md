# QueryRows Method


Returns one or more rows from a table, beginning at the current cursor position.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool QueryRows(
	int lRowCount,
	out SRow[] sRows
)
```
**VB**
``` VB
Public Function QueryRows ( 
	lRowCount As Integer,
	<OutAttribute> ByRef sRows As SRow()
) As Boolean
```
**C++**
``` C++
public:
bool QueryRows(
	int lRowCount, 
	[OutAttribute] array<SRow>^% sRows
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>Maximum number of rows to be returned.</dd><dt>  <a href="T_MAPI_NET_SRow.md">SRow</a>[]</dt><dd>an SRow array holding the table rows.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPITable.md">MAPITable Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
