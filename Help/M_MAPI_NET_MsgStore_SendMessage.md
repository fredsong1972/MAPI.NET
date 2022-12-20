# SendMessage Method


Send message



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SendMessage(
	MAPIMessage message,
	bool saveToSentFolder
)
```
**VB**
``` VB
Public Function SendMessage ( 
	message As MAPIMessage,
	saveToSentFolder As Boolean
) As Boolean
```
**C++**
``` C++
public:
bool SendMessage(
	MAPIMessage^ message, 
	bool saveToSentFolder
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a></dt><dd>message</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a></dt><dd>if save message to sent folder</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MsgStore.md">MsgStore Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
