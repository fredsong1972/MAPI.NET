# AbortSubmit Method


Attempts to remove a message from the outgoing queue.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT AbortSubmit(
	uint cbEntryID,
	IntPtr lpEntryID,
	uint ulFlags
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function AbortSubmit ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT AbortSubmit(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the message to remove from the outgoing queue.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message was successfully removed from the outgoing queue; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_IMsgStore.md">IMsgStore Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
