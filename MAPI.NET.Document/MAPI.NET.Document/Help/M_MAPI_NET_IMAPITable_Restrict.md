# Restrict Method


Applies a filter to a table, reducing the row set to only those rows matching the specified criteria.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Restrict(
	out IntPtr lpRestriction,
	uint ulFlags
)
```
**VB**
``` VB
Function Restrict ( 
	<OutAttribute> ByRef lpRestriction As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT Restrict(
	[OutAttribute] IntPtr% lpRestriction, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to an SRestriction structure defining the conditions of the filter. Passing NULL in the lpRestriction parameter removes the current filter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls the timing of the restriction operation.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the filter was successfully applied; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
