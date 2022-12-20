# SeekRow Method


Moves the cursor to a specific position in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SeekRow(
	int bkOrigin,
	int lRowCount,
	out IntPtr lplRowsSought
)
```
**VB**
``` VB
Function SeekRow ( 
	bkOrigin As Integer,
	lRowCount As Integer,
	<OutAttribute> ByRef lplRowsSought As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SeekRow(
	int bkOrigin, 
	int lRowCount, 
	[OutAttribute] IntPtr% lplRowsSought
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>The bookmark identifying the starting position for the seek operation.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>The signed count of the number of rows to move, starting from the bookmark identified by the bkOrigin parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>If lRowCount is a valid pointer on input, lplRowsSought points to the number of rows that were processed in the seek operation, the sign of which indicates the direction of search, forward or backward. If lRowCount is negative, then lplRowsSought is negative.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the seek operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
