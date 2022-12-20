# GetMessageStatus Method


Obtains the status associated with a message in a particular folder.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetMessageStatus(
	uint cbEntryID,
	IntPtr lpEntryID,
	uint ulFlags,
	out uint lpulMessageStatus
)
```
**VB**
``` VB
Function GetMessageStatus ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulMessageStatus As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT GetMessageStatus(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpulMessageStatus
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier for the message whose status is obtained.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a pointer to a bitmask of flags that indicate the message's status.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message status was successfully retrieved; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
