# CopyMessages Method


Copies or moves one or more messages.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CopyMessages(
	IntPtr lpMsgList,
	IntPtr lpInterface,
	IntPtr lpDestFolder,
	uint ulUIParam,
	IntPtr lpProgress,
	uint ulFlags
)
```
**VB**
``` VB
Function CopyMessages ( 
	lpMsgList As IntPtr,
	lpInterface As IntPtr,
	lpDestFolder As IntPtr,
	ulUIParam As UInteger,
	lpProgress As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT CopyMessages(
	IntPtr lpMsgList, 
	IntPtr lpInterface, 
	IntPtr lpDestFolder, 
	unsigned int ulUIParam, 
	IntPtr lpProgress, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of ENTRYLIST structures that identify the message or messages to copy or move.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the destination folder pointed to by the lpDestFolder parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the open folder to receive the copied or moved messages.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of any dialog boxes or windows this method displays.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress object that displays a progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the copy or move operation is accomplished.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message or messages have been successfully copied or moved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
