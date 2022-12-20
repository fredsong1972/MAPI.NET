# SetSearchCriteria Method


Establishes search criteria for the container.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SetSearchCriteria(
	IntPtr pRestriction,
	IntPtr pContainerList,
	uint ulSearchFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SetSearchCriteria ( 
	pRestriction As IntPtr,
	pContainerList As IntPtr,
	ulSearchFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SetSearchCriteria(
	IntPtr pRestriction, 
	IntPtr pContainerList, 
	unsigned int ulSearchFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an SRestriction structure that defines the search criteria.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of entry identifiers that represent containers to be included in the search.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that control how the search is performed.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the search criteria was successfully set; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIContainer.md">IMAPIContainer Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
