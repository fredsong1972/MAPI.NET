# GetPropList Method


Returns property tags for all properties.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT GetPropList(
	uint ulFlags,
	out IntPtr lppPropTagArray
)
```
**VB**
``` VB
Function GetPropList ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lppPropTagArray As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetPropList(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppPropTagArray
)
```
**F#**
``` F#
abstract GetPropList : 
        ulFlags : uint32 * 
        lppPropTagArray : IntPtr byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the format for the strings in the returned property tags.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the property tag array that contains tags for all of the object's properties.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK All of the property tags were returned successfully. MAPI_E_BAD_CHARWIDTH Either the MAPI_UNICODE flag was set and the implementation does not support Unicode, or MAPI_UNICODE was not set and the implementation supports only Unicode.

## See Also


#### Reference
<a href="a20f5817-5533-814e-fd1d-0d3a9179b1b4.md">IMAPIProp Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
