# SetMessageStatus Method


Sets the status associated with a message.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetMessageStatus(
	uint cbEntryID,
	IntPtr lpEntryID,
	uint ulNewStatus,
	uint ulNewStatusMask,
	out uint lpulOldStatus
)
```
**VB**
``` VB
Function SetMessageStatus ( 
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	ulNewStatus As UInteger,
	ulNewStatusMask As UInteger,
	<OutAttribute> ByRef lpulOldStatus As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT SetMessageStatus(
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	unsigned int ulNewStatus, 
	unsigned int ulNewStatusMask, 
	[OutAttribute] unsigned int% lpulOldStatus
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier for the message whose status is set.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The new status to be assigned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that is applied to the new status and indicates the flags to be set.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the previous status of the message.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message status was successfully set; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
