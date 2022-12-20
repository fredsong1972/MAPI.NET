# AddAttachment(String, String) Method


Add a file to an attachment of the message



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool AddAttachment(
	string filePath,
	string name
)
```
**VB**
``` VB
Public Function AddAttachment ( 
	filePath As String,
	name As String
) As Boolean
```
**C++**
``` C++
public:
bool AddAttachment(
	String^ filePath, 
	String^ name
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>the file paty of the original file.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>Display name of the new attachment</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the attachment was added successfully; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="Overload_MAPI_NET_MAPIMessage_AddAttachment.md">AddAttachment Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
