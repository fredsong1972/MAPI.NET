# GetRecipientTable Method


Returns the message's recipient table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT GetRecipientTable(
	uint ulFlags,
	out IntPtr lppTable
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function GetRecipientTable ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lppTable As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT GetRecipientTable(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppTable
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls the return of the table.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer to the recipient table.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the recipient table was returned successfully; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
