# GetReceiveFolder Method


Obtains the folder that was established as the destination for incoming messages of a specified message class or as the default receive folder for the message store.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetReceiveFolder(
	string lpszMessageClass,
	uint ulFlags,
	out uint cbEntryID,
	out IntPtr lppEntryID,
	StringBuilder lppszExplicitClass
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetReceiveFolder ( 
	lpszMessageClass As String,
	ulFlags As UInteger,
	<OutAttribute> ByRef cbEntryID As UInteger,
	<OutAttribute> ByRef lppEntryID As IntPtr,
	lppszExplicitClass As StringBuilder
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetReceiveFolder(
	String^ lpszMessageClass, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% cbEntryID, 
	[OutAttribute] IntPtr% lppEntryID, 
	StringBuilder^ lppszExplicitClass
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>A message class is associated with a receive folder</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the passed-in and returned strings.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the byte count in the entry identifier pointed to by the lppEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the entry identifier for the requested receive folder.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.text.stringbuilder" target="_blank" rel="noopener noreferrer">StringBuilder</a></dt><dd>A pointer to a pointer to the message class that explicitly sets as its receive folder the folder pointed to by lppEntryID.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the receive folder was successfully returned; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_IMsgStore.md">IMsgStore Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
