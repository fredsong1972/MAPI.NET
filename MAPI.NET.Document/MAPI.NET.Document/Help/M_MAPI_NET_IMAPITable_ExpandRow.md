# ExpandRow Method


Expands a collapsed table category, adding the leaf or lower-level heading rows belonging to the category to the table view.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT ExpandRow(
	uint cbInstanceKey,
	IntPtr pbInstanceKey,
	uint ulRowCount,
	uint ulFlags,
	IntPtr lppRows,
	IntPtr lpulMoreRows
)
```
**VB**
``` VB
Function ExpandRow ( 
	cbInstanceKey As UInteger,
	pbInstanceKey As IntPtr,
	ulRowCount As UInteger,
	ulFlags As UInteger,
	lppRows As IntPtr,
	lpulMoreRows As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT ExpandRow(
	unsigned int cbInstanceKey, 
	IntPtr pbInstanceKey, 
	unsigned int ulRowCount, 
	unsigned int ulFlags, 
	IntPtr lppRows, 
	IntPtr lpulMoreRows
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The count of bytes in the PR_INSTANCE_KEY property pointed to by the pbInstanceKey parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the PR_INSTANCE_KEY property that identifies the heading row for the category.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The maximum number of rows to return in the lppRows parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an SRowSet structure receiving the first (up to ulRowCount) rows that have been inserted into the table view as a result of the expansion.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the total number of rows that were added to the table view.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the category was expanded successfully; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
