# Advise Method


Registers to receive notification of specified events that affect the message store.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT Advise(
	uint cbEntryID,
	IntPtr lpEntryID,
	uint ulEventMask,
	IMAPIAdviseSink pAdviseSink,
	out uint lpulConnection
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function Advise ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	ulEventMask As UInteger,
	pAdviseSink As IMAPIAdviseSink,
	<OutAttribute> ByRef lpulConnection As UInteger
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT Advise(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	unsigned int ulEventMask, 
	[InAttribute] IMAPIAdviseSink^ pAdviseSink, 
	[OutAttribute] unsigned int% lpulConnection
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the folder or message about which notifications should be generated, or null. If lpEntryID is set to NULL, Advise registers for notifications on the entire message store.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration.</dd><dt>  <a href="T_MAPI_NET_IMAPIAdviseSink.md">IMAPIAdviseSink</a></dt><dd>A pointer to an advise sink object to receive the subsequent notifications.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a nonzero number that represents the connection between the caller's advise sink object and the session.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the registration was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMsgStore.md">IMsgStore Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
