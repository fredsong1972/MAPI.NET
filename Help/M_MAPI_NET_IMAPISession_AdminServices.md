# AdminServices Method


Returns an IMsgServiceAdmin pointer for making changes to message services.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT AdminServices(
	uint ulFlags,
	out IntPtr lppServiceAdmin
)
```
**VB**
``` VB
Function AdminServices ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lppServiceAdmin As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT AdminServices(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppServiceAdmin
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to a message service administration object.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if a pointer to a message service administration object was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
