# Advise Method


Registers an advise sink object to receive notification of specified events affecting the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Advise(
	uint ulEventMask,
	IntPtr lpAdviseSink,
	IntPtr lpulConnection
)
```
**VB**
``` VB
Function Advise ( 
	ulEventMask As UInteger,
	lpAdviseSink As IntPtr,
	lpulConnection As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT Advise(
	unsigned int ulEventMask, 
	IntPtr lpAdviseSink, 
	IntPtr lpulConnection
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Value indicating the type of event that will generate the notification.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to an advise sink object to receive the subsequent notifications. This advise sink object must have been already allocated.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a nonzero value that represents the successful notification registration.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the notification registration successfully completed; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
