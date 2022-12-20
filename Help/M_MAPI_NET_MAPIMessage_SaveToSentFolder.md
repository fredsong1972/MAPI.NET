# SaveToSentFolder Method


Set if save a message to the sent folder after sent.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SaveToSentFolder(
	MsgStore messageStore,
	bool bSaveToSentFolder
)
```
**VB**
``` VB
Public Function SaveToSentFolder ( 
	messageStore As MsgStore,
	bSaveToSentFolder As Boolean
) As Boolean
```
**C++**
``` C++
public:
bool SaveToSentFolder(
	MsgStore^ messageStore, 
	bool bSaveToSentFolder
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_MsgStore.md">MsgStore</a></dt><dd>Message store object</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a></dt><dd>Boolean flag which controls if save the message to the sent folder after sent.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successfully; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
