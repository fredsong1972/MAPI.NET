# DeleteMessages Method


Delete messages



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool DeleteMessages(
	List<MAPIMessage> messages
)
```
**VB**
``` VB
Public Function DeleteMessages ( 
	messages As List(Of MAPIMessage)
) As Boolean
```
**C++**
``` C++
public:
bool DeleteMessages(
	List<MAPIMessage^>^ messages
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.collections.generic.list-1" target="_blank" rel="noopener noreferrer">List</a>(<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage</a>)</dt><dd>MAPI messages to delete</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the specified messages were successfully deleted; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIFolder.md">MAPIFolder Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
