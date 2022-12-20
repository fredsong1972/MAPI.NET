# GetSearchCriteria Method


Obtains the search criteria for the container.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetSearchCriteria(
	uint ulFlags,
	out IntPtr pRestriction,
	out IntPtr pContainerList,
	out uint searchState
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetSearchCriteria ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef pRestriction As IntPtr,
	<OutAttribute> ByRef pContainerList As IntPtr,
	<OutAttribute> ByRef searchState As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetSearchCriteria(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% pRestriction, 
	[OutAttribute] IntPtr% pContainerList, 
	[OutAttribute] unsigned int% searchState
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the passed-in strings.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to an SRestriction structure that defines the search criteria.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to an array of entry identifiers that represent containers to be included in the search.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a bitmask of flags used to indicate the current state of the search</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the search criteria was successfully obtained; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
