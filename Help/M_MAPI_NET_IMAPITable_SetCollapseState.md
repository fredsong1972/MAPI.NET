# SetCollapseState Method


Rebuilds the current expanded or collapsed state of a categorized table using data that was saved by a prior call to the IMAPITable::GetCollapseState method.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetCollapseState(
	uint ulFlags,
	uint cbCollapseState,
	IntPtr pbCollapseState,
	IntPtr lpbkLocation
)
```
**VB**
``` VB
Function SetCollapseState ( 
	ulFlags As UInteger,
	cbCollapseState As UInteger,
	pbCollapseState As IntPtr,
	lpbkLocation As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SetCollapseState(
	unsigned int ulFlags, 
	unsigned int cbCollapseState, 
	IntPtr pbCollapseState, 
	IntPtr lpbkLocation
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Count of bytes in the structure pointed to by the pbCollapseState parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the structures containing the data needed to rebuild the table view.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a bookmark identifying the row in the table at which the collapsed or expanded state should be rebuilt.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the state of the categorized table was successfully rebuilt; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
