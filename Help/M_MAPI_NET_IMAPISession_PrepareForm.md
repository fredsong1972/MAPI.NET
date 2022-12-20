# PrepareForm Method


Creates a numeric token that the IMAPISession::ShowForm method uses to access a message.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT PrepareForm(
	ref Guid lpInterface,
	IntPtr lpMessage,
	out uint lpulMessageToken
)
```
**VB**
``` VB
Function PrepareForm ( 
	ByRef lpInterface As Guid,
	lpMessage As IntPtr,
	<OutAttribute> ByRef lpulMessageToken As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT PrepareForm(
	Guid% lpInterface, 
	IntPtr lpMessage, 
	[OutAttribute] unsigned int% lpulMessageToken
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the message.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the message to be displayed in the form.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a message token, which is used by the IMAPISession::ShowForm method to access the message pointed to by lpMessage.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the form preparation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
