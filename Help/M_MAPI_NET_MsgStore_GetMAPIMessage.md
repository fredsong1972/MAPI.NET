# GetMAPIMessage Method


Gets MAPI Message per the entry identification



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public MAPIMessage GetMAPIMessage(
	EntryID id,
	uint messageFlags,
	string messageClass
)
```
**VB**
``` VB
Public Function GetMAPIMessage ( 
	id As EntryID,
	messageFlags As UInteger,
	messageClass As String
) As MAPIMessage
```
**C++**
``` C++
public:
MAPIMessage^ GetMAPIMessage(
	EntryID^ id, 
	unsigned int messageFlags, 
	String^ messageClass
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_EntryID.md">EntryID</a></dt><dd>The entry identification of the message to get</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The message flags which is assocaited with the message</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The message class which is associated with the message</dd></dl>

#### Return Value
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>  
MAPI Message instance

## See Also


#### Reference
<a href="T_MAPI_NET_MsgStore.md">MsgStore Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
