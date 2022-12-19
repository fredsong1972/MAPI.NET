# SetProps Method


Updates one or more properties.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT SetProps(
	uint cValues,
	IntPtr lpPropArray,
	out IntPtr lppProblems
)
```
**VB**
``` VB
Function SetProps ( 
	cValues As UInteger,
	lpPropArray As IntPtr,
	<OutAttribute> ByRef lppProblems As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SetProps(
	unsigned int cValues, 
	IntPtr lpPropArray, 
	[OutAttribute] IntPtr% lppProblems
)
```
**F#**
``` F#
abstract SetProps : 
        cValues : uint32 * 
        lpPropArray : IntPtr * 
        lppProblems : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of property values pointed to by the lpPropArray parameter. The cValues parameter must not be 0.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of SPropValue structures that contain property values to be updated.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>On input, a pointer to a pointer to an SPropProblemArray structure; otherwise, NULL, indicating no need for error information. If lppProblems is a valid pointer on input, SetProps returns detailed information about errors in updating one or more properties.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the properties were successfully updated.; otherwise, failed.

## See Also


#### Reference
<a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
