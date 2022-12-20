# SetReceiveFolder Method


Establishes a folder as the destination for incoming messages of a particular message class.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT SetReceiveFolder(
	string lpszMessageClass,
	uint ulFlags,
	uint cbEntryID,
	IntPtr lpEntryID
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function SetReceiveFolder ( 
	lpszMessageClass As String,
	ulFlags As UInteger,
	cbEntryID As UInteger,
	lpEntryID As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT SetReceiveFolder(
	String^ lpszMessageClass, 
	unsigned int ulFlags, 
	unsigned int cbEntryID, 
	IntPtr lpEntryID
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>A message class is associated with the new receive folder</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the text in the passed-in strings.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>pointer to the entry identifier of the folder to establish as the receive folder.If the lpEntryID parameter is set to NULL, SetReceiveFolder replaces the current receive folder with the message store's default.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if a receive folder was successfully established; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_IMsgStore.md">IMsgStore Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
