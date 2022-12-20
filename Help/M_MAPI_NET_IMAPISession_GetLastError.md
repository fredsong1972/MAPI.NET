# GetLastError Method


Returns a MAPIERROR structure that contains information about the previous session error.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetLastError(
	int hResult,
	uint ulFlags,
	out IntPtr lppMAPIError
)
```
**VB**
``` VB
Function GetLastError ( 
	hResult As Integer,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppMAPIError As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetLastError(
	int hResult, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppMAPIError
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>A handle to the error value generated in the previous method call.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of strings returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to a MAPIERROR structure that contains version, component, and context information for the error.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
