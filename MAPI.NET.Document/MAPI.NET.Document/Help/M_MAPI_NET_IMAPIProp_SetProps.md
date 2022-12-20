# SetProps Method


Updates one or more properties.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SetProps(
	uint cValues,
	IntPtr lpPropArray,
	out IntPtr lppProblems
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SetProps ( 
	cValues As UInteger,
	lpPropArray As IntPtr,
	<OutAttribute> ByRef lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SetProps(
	unsigned int cValues, 
	IntPtr lpPropArray, 
	[OutAttribute] IntPtr% lppProblems
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of property values pointed to by the lpPropArray parameter. The cValues parameter must not be 0.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of SPropValue structures that contain property values to be updated.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, indicating no need for error information. If lppProblems is a valid pointer on input, SetProps returns detailed information about errors in updating one or more properties.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the properties were successfully updated.; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
