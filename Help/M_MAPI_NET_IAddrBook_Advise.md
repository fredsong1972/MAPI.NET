# Advise Method


Registers a client or service provider to receive notifications about changes to one or more entries in the address book.



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
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the address book container, messaging user, or distribution list that will generate a notification when a change occurs of the type or types described in the ulEventMask parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>One or more notification events that the caller is registering to receive. Each event is associated with a particular notification structure that contains information about the change that occurred.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the advise sink object to be called when the event for which notification has been requested occurs.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a nonzero connection number that represents the notification registration.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the notification registration was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
