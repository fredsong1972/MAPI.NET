# CopyMessages(MAPIMessage, MAPIFolder) Method


Copy one message to different folder



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool CopyMessages(
	MAPIMessage message,
	MAPIFolder destFolder
)
```
**VB**
``` VB
Public Function CopyMessages ( 
	message As MAPIMessage,
	destFolder As MAPIFolder
) As Boolean
```
**C++**
``` C++
public:
bool CopyMessages(
	MAPIMessage^ message, 
	MAPIFolder^ destFolder
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a></dt><dd>message</dd><dt>  <a href="T_MAPI_NET_MAPIFolder.md">MAPIFolder</a></dt><dd>destination folder</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIFolder.md">MAPIFolder Class</a>  
<a href="Overload_MAPI_NET_MAPIFolder_CopyMessages.md">CopyMessages Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
