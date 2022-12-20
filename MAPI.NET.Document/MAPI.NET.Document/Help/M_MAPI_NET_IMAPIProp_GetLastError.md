# GetLastError Method


Returns a MAPIERROR structure that contains information about the previous error.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetLastError(
	int hResult,
	uint ulFlags,
	out IntPtr lppMAPIError
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetLastError ( 
	hResult As Integer,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppMAPIError As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetLastError(
	int hResult, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppMAPIError
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>A handle to the error code generated in the previous method call.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that indicates the format for the text returned in the MAPIERROR structure pointed to by lppMAPIError.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the MAPIERROR structure that contains version, component, and context information for the error. The lppMAPIError parameter can be set to NULL if there is no error information to return.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the error information was returned;otherwise failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
