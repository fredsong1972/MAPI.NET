# SetReadFlags Method


Sets or clears the MSGFLAG_READ flag in the PR_MESSAGE_FLAGS property of one or more of the folder's messages, and manages the sending of read reports.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetReadFlags(
	IntPtr lpMsgList,
	uint ulUIParam,
	IntPtr lpProgress,
	uint ulFlags
)
```
**VB**
``` VB
Function SetReadFlags ( 
	lpMsgList As IntPtr,
	ulUIParam As UInteger,
	lpProgress As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT SetReadFlags(
	IntPtr lpMsgList, 
	unsigned int ulUIParam, 
	IntPtr lpProgress, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an array of ENTRYLIST structures that identify the message or messages that have read flags to set or clear.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a progress object that displays a progress indicator.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the setting of a message's read flag and the processing of read reports.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the read flag for the specified message or messages was successfully set or cleared; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
