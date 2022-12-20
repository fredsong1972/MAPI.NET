# FindRow Method


Finds the next row in a table that matches specific search criteria and moves the cursor to that row.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT FindRow(
	out IntPtr lpRestriction,
	uint BkOrigin,
	uint ulFlags
)
```
**VB**
``` VB
Function FindRow ( 
	<OutAttribute> ByRef lpRestriction As IntPtr,
	BkOrigin As UInteger,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT FindRow(
	[OutAttribute] IntPtr% lpRestriction, 
	unsigned int BkOrigin, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an SRestriction structure that describes the search criteria.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bookmark identifying the row where FindRow should begin its search.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the direction of the search.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the find operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
