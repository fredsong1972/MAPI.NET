# OpenAttach Method


Opens an attachment.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[PreserveSigAttribute]
HRESULT OpenAttach(
	uint ulAttachmentNum,
	IntPtr lpInterface,
	uint uFlags,
	out IntPtr lpAttach
)
```
**VB**
``` VB
<PreserveSigAttribute>
Function OpenAttach ( 
	ulAttachmentNum As UInteger,
	lpInterface As IntPtr,
	uFlags As UInteger,
	<OutAttribute> ByRef lpAttach As IntPtr
) As HRESULT
```
**C++**
``` C++
[PreserveSigAttribute]
HRESULT OpenAttach(
	unsigned int ulAttachmentNum, 
	IntPtr lpInterface, 
	unsigned int uFlags, 
	[OutAttribute] IntPtr% lpAttach
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Index number of the attachment to open, as stored in the attachment's PR_ATTACH_NUM property.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the interface identifier (IID) representing the interface to be used to access the attachment. Passing NULL results in the attachment's standard interface, or IAttach, being returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls how the attachment is opened.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a pointer to the open attachment.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the attachment was successfully opened; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMessage.md">IMessage Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
