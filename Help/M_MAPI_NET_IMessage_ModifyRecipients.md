# ModifyRecipients Method


Adds, deletes, or modifies message recipients.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT ModifyRecipients(
	uint ulFlags,
	IntPtr lpMods
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function ModifyRecipients ( 
	ulFlags As UInteger,
	lpMods As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT ModifyRecipients(
	unsigned int ulFlags, 
	IntPtr lpMods
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls the recipient changes. If zero is passed for the ulFlags parameter, ModifyRecipients replaces all existing recipients with the recipient list pointed to by the lpMods parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to an ADRLIST structure that contains a list of recipients to be added, deleted, or modified in the message.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the recipient list was successfully modified; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
