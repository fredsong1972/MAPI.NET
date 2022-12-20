# CreateMessage Method


Creates a new message.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CreateMessage(
	IntPtr lpInterface,
	uint ulFlags,
	out IntPtr lppMessage
)
```
**VB**
``` VB
Function CreateMessage ( 
	lpInterface As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lppMessage As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT CreateMessage(
	IntPtr lpInterface, 
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppMessage
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the new message.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the message is created.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the newly created message.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the message was successfully created; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIFolder.md">IMAPIFolder Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
