# GetSearchPath Method


Returns an ordered list of entry identifiers of the containers to be included in the name resolution process initiated by the IAddrBook::ResolveName method.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetSearchPath(
	uint ulFlags,
	out IntPtr lppSearchPath
)
```
**VB**
``` VB
Function GetSearchPath ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lppSearchPath As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetSearchPath(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppSearchPath
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the strings returned in the search path.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to an ordered list of container entry identifiers.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the search path was successfully retrieved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
