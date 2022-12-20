# FreeBookmark Method


Releases the memory associated with a bookmark.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT FreeBookmark(
	IntPtr bkPosition
)
```
**VB**
``` VB
Function FreeBookmark ( 
	bkPosition As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT FreeBookmark(
	IntPtr bkPosition
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>The bookmark to be freed, created by calling the IMAPITable::CreateBookmark method.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the bookmark was successfully freed; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
