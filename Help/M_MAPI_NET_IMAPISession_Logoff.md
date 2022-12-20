# Logoff Method


Ends a MAPI session.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Logoff(
	uint ulUIParam,
	uint ulFlags,
	uint ulReserved
)
```
**VB**
``` VB
Function Logoff ( 
	ulUIParam As UInteger,
	ulFlags As UInteger,
	ulReserved As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT Logoff(
	unsigned int ulUIParam, 
	unsigned int ulFlags, 
	unsigned int ulReserved
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of any dialog boxes or windows to be displayed.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that control the logoff operation.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the logoff operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
