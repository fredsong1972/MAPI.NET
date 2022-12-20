# Advise Method


Registers to receive notification of specified events that affect the session.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Advise(
	uint cbEntryID,
	IntPtr lpEntryID,
	uint ulEventMask,
	IntPtr lpAdviseSink,
	out uint lpulConnection
)
```
**VB**
``` VB
Function Advise ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	ulEventMask As UInteger,
	lpAdviseSink As IntPtr,
	<OutAttribute> ByRef lpulConnection As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT Advise(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	unsigned int ulEventMask, 
	IntPtr lpAdviseSink, 
	[OutAttribute] unsigned int% lpulConnection
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the address book or message store object about which notifications should be generated, or NULL, which indicates that the client is registering to receive notifications about events that affect only the session.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A mask of values that indicate the types of notification events that the client is interested in and should be included in the registration.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an advise sink object to receive the subsequent notifications.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a nonzero number that represents the connection between the caller's advise sink object and the session.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the registration was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
