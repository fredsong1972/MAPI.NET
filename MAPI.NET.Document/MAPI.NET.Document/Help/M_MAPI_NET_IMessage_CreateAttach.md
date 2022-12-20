# CreateAttach Method


Creates a new attachment.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT CreateAttach(
	IntPtr lpInterface,
	uint uFlags,
	out uint ulAttachmentNum,
	out IntPtr lpAttach
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function CreateAttach ( 
	lpInterface As IntPtr,
	uFlags As UInteger,
	<OutAttribute> ByRef ulAttachmentNum As UInteger,
	<OutAttribute> ByRef lpAttach As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT CreateAttach(
	IntPtr lpInterface, 
	unsigned int uFlags, 
	[OutAttribute] unsigned int% ulAttachmentNum, 
	[OutAttribute] IntPtr% lpAttach
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the interface identifier (IID) representing the interface to be used to access the message. Passing NULL results in the message's standard interface, or IMessage, being returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls how the attachment is created.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>An index number identifying the newly created attachment.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer to the open attachment object.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the attachment was successfully created; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
