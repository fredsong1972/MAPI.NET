# Unadvise Method


Cancels a notification registration previously established for an address book entry.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Unadvise(
	uint ulConnection
)
```
**VB**
``` VB
Function Unadvise ( 
	ulConnection As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT Unadvise(
	unsigned int ulConnection
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A connection number that represents the registration to be canceled. The ulConnection parameter should contain a value returned by a prior call to the IAddrBook::Advise method.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the registration was successfully canceled; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
