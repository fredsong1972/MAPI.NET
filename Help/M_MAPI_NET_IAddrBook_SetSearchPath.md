# SetSearchPath Method


Sets a new search path in the profile that is used for the name resolution process.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetSearchPath(
	uint ulFlags,
	IntPtr lpSearchPath
)
```
**VB**
``` VB
Function SetSearchPath ( 
	ulFlags As UInteger,
	lpSearchPath As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SetSearchPath(
	unsigned int ulFlags, 
	IntPtr lpSearchPath
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the SRowSet structure used to hold the search path. The first property for each aRow member in SRowSet must be PR_ENTRYID.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if The search path was successfully set; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
