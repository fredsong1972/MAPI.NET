# CreateBookmark Method


Creates a bookmark at the table's current position.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CreateBookmark(
	IntPtr lpbkPosition
)
```
**VB**
``` VB
Function CreateBookmark ( 
	lpbkPosition As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT CreateBookmark(
	IntPtr lpbkPosition
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the returned 32-bit bookmark value. This bookmark can later be passed in a call to the IMAPITable::SeekRow method</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
