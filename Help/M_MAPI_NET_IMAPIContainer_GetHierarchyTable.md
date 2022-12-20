# GetHierarchyTable Method


Returns a pointer to the container's hierarchy table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetHierarchyTable(
	uint ulFlags,
	out IntPtr pTable
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetHierarchyTable ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef pTable As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetHierarchyTable(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% pTable
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how information is returned in the table.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the hierarchy table.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the hierarchy table was successfully retrieved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
