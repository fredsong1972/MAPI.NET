# OpenMsgStore Method


Opens a message store and returns an IMsgStore pointer for further access.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT OpenMsgStore(
	uint ulUIParam,
	uint cbEntryID,
	IntPtr lpEntryID,
	IntPtr lpInterface,
	uint ulFlags,
	out IntPtr lppMDB
)
```
**VB**
``` VB
Function OpenMsgStore ( 
	ulUIParam As UInteger,
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppMDB As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OpenMsgStore(
	unsigned int ulUIParam, 
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppMDB
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the common address dialog box and other related displays.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the message store to be opened. The lpEntryID parameter must not be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the message store.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the object is opened.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer of the message store.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message store was successfully opened; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
