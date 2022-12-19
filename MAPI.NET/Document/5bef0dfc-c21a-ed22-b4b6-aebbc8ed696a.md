# GetLastError Method


Returns a MAPIERROR structure that contains information about the previous error.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
abstract GetLastError : 
        hResult : int * 
        ulFlags : uint32 * 
        lppMAPIError : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>A handle to the error code generated in the previous method call.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that indicates the format for the text returned in the MAPIERROR structure pointed to by lppMAPIError.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the MAPIERROR structure that contains version, component, and context information for the error. The lppMAPIError parameter can be set to NULL if there is no error information to return.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the error information was returned;otherwise failed.

## See Also


#### Reference
<a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
