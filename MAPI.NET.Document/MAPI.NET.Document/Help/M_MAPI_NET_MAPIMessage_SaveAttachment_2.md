# SaveAttachment(String, Int32, String) Method


Save attachment to a disk file.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

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



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>output folder</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.int32" target="_blank" rel="noopener noreferrer">Int32</a></dt><dd>index of the attachment</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>output file name</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a>  
the path of the output file

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="Overload_MAPI_NET_MAPIMessage_SaveAttachment.md">SaveAttachment Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
