# CollapseRow Method


Collapses an expanded table category, removing any lower-level headings and leaf rows belonging to the category from the table view.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CollapseRow(
	uint cbInstanceKey,
	IntPtr pbInstanceKey,
	uint ulFlags,
	IntPtr lpulRowCount
)
```
**VB**
``` VB
Function CollapseRow ( 
	cbInstanceKey As UInteger,
	pbInstanceKey As IntPtr,
	ulFlags As UInteger,
	lpulRowCount As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT CollapseRow(
	unsigned int cbInstanceKey, 
	IntPtr pbInstanceKey, 
	unsigned int ulFlags, 
	IntPtr lpulRowCount
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of bytes in the PR_INSTANCE_KEY property pointed to by the pbInstanceKey parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the PR_INSTANCE_KEY property that identifies the heading row for the category.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the total number of rows that are being removed from the table view.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the collapse operation has succeeded; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
