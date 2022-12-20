# GetCollapseState Method


Returns the data that is needed to rebuild the current collapsed or expanded state of a categorized table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetCollapseState(
	uint ulFlags,
	uint cbInstanceKey,
	IntPtr lpbInstanceKey,
	IntPtr lpcbCollapseState,
	IntPtr lppbCollapseState
)
```
**VB**
``` VB
Function GetCollapseState ( 
	ulFlags As UInteger,
	cbInstanceKey As UInteger,
	lpbInstanceKey As IntPtr,
	lpcbCollapseState As IntPtr,
	lppbCollapseState As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetCollapseState(
	unsigned int ulFlags, 
	unsigned int cbInstanceKey, 
	IntPtr lpbInstanceKey, 
	IntPtr lpcbCollapseState, 
	IntPtr lppbCollapseState
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of bytes in the instance key pointed to by the lpbInstanceKey parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the PR_INSTANCE_KEY property of the row at which the current collapsed or expanded state should be rebuilt.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the count of structures pointed to by the lppbCollapseState parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to structures that contain data that describes the current table view.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the state for the categorized table was successfully saved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
