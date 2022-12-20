# EmptyFolder Method


Deletes all messages and subfolders from a folder without deleting the folder itself.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT EmptyFolder(
	uint ulUIParam,
	IntPtr lpProgress,
	uint ulFlags
)
```
**VB**
``` VB
Function EmptyFolder ( 
	ulUIParam As UInteger,
	lpProgress As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT EmptyFolder(
	unsigned int ulUIParam, 
	IntPtr lpProgress, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress object that displays a progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the folder is emptied.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the folder was successfully emptied; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
