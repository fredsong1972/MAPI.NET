# SaveAttachment(String, Int32, String) Method


Save attachment to a disk file.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
public string SaveAttachment(
	string folder,
	int index,
	string fileName
)
```
**VB**
``` VB
Public Function SaveAttachment ( 
	folder As String,
	index As Integer,
	fileName As String
) As String
```
**C++**
``` C++
public:
String^ SaveAttachment(
	String^ folder, 
	int index, 
	String^ fileName
)
```
**F#**
``` F#
member SaveAttachment : 
        folder : string * 
        index : int * 
        fileName : string -> string 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>output folder</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>index of the attachment</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>output file name</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a>  
the path of the output file

## See Also


#### Reference
<a href="29b8d96c-1ec2-828d-35a5-fae12d8802c8.md">MAPIMessage Class</a>  
<a href="20c0fe52-b7a4-b293-ac57-16aba77ed94e.md">SaveAttachment Overload</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  