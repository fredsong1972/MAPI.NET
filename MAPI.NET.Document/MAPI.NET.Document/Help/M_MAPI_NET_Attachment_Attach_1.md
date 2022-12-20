# Attach(String, String) Method


Creates a mail attachment using the content from the specified file, and the specified name.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool Attach(
	string filePath,
	string name
)
```
**VB**
``` VB
Public Function Attach ( 
	filePath As String,
	name As String
) As Boolean
```
**C++**
``` C++
public:
bool Attach(
	String^ filePath, 
	String^ name
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>file path</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>attachment name</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true if successful; otherwise, false

## See Also


#### Reference
<a href="T_MAPI_NET_Attachment.md">Attachment Class</a>  
<a href="Overload_MAPI_NET_Attachment_Attach.md">Attach Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
