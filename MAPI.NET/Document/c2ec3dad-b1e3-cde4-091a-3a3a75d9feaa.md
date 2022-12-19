# AddAttachment(String, String) Method


Add a file to an attachment of the message



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
member AddAttachment : 
        filePath : string * 
        name : string -> bool 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>the file paty of the original file.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>Display name of the new attachment</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the attachment was added successfully; otherwise, failed

## See Also


#### Reference
<a href="29b8d96c-1ec2-828d-35a5-fae12d8802c8.md">MAPIMessage Class</a>  
<a href="451c8a33-e103-5818-3316-7f4c519a9dbb.md">AddAttachment Overload</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
