# DeleteMessages Method


Deletes one or more messages.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT DeleteMessages(
	IntPtr lpMsgList,
	uint ulUIParam,
	IntPtr lpProgress,
	uint ulFlags
)
```
**VB**
``` VB
Function DeleteMessages ( 
	lpMsgList As IntPtr,
	ulUIParam As UInteger,
	lpProgress As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT DeleteMessages(
	IntPtr lpMsgList, 
	unsigned int ulUIParam, 
	IntPtr lpProgress, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an ENTRYLIST structure that contains the number of messages to delete and an array of ENTRYID structures that identify the messages.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the progress indicator. The ulUIParam parameter is ignored unless the MESSAGE_DIALOG flag is set in the ulFlags parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress object that displays a progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the messages are deleted.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the specified message or messages were successfully deleted; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
