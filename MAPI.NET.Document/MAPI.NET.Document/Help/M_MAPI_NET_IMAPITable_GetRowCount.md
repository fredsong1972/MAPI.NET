# GetRowCount Method


Returns the total number of rows in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetRowCount(
	uint ulFlags,
	out uint lpulCount
)
```
**VB**
``` VB
Function GetRowCount ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulCount As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT GetRowCount(
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpulCount
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Pointer to the number of rows in the table.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the row count was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
