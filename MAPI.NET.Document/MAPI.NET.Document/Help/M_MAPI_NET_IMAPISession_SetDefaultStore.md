# SetDefaultStore Method


Establishes a message store as the default message store for the session.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetDefaultStore(
	uint ulFlags,
	uint cbEntryID,
	IntPtr lpEntryID
)
```
**VB**
``` VB
Function SetDefaultStore ( 
	ulFlags As UInteger,
	cbEntryID As UInteger,
	lpEntryID As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT SetDefaultStore(
	unsigned int ulFlags, 
	unsigned int cbEntryID, 
	IntPtr lpEntryID
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the setting of the default message store.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the message store that is intended as the default.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded and returned the expected value or values; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
