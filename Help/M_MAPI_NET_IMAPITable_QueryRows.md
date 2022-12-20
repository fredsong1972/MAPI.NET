# QueryRows Method


Returns one or more rows from a table, beginning at the current cursor position.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT QueryRows(
	int lRowCount,
	uint ulFlags,
	out IntPtr lppRows
)
```
**VB**
``` VB
Function QueryRows ( 
	lRowCount As Integer,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppRows As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT QueryRows(
	int lRowCount, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppRows
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>Maximum number of rows to be returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that control how rows are returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer to an SRowSet structure holding the table rows.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the rows were successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
