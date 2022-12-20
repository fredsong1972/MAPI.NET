# Unadvise Method


Cancels the sending of notifications previously set up with a call to the IMAPITable::Advise method.



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
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The number of the registration connection returned by a call to IMAPITable::Advise.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
