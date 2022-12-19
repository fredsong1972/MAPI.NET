# CreateBookmark Method


Creates a bookmark at the table's current position.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
abstract CreateBookmark : 
        lpbkPosition : IntPtr -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the returned 32-bit bookmark value. This bookmark can later be passed in a call to the IMAPITable::SeekRow method</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the call succeeded and has returned the expected value or values; otherwise, failed.

## See Also


#### Reference
<a href="06a9b727-f5d6-e992-c936-a2712197dcee.md">IMAPITable Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
