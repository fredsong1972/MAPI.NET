# GetContentsTable Method


Returns a pointer to the container's contents table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetContentsTable(
	uint ulFlags,
	out IntPtr pTable
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetContentsTable ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef pTable As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetContentsTable(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% pTable
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the contents table is returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the contents table.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the contents table was successfully retrieved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
