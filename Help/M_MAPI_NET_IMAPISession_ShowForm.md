# ShowForm Method


Displays a form.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT ShowForm(
	uint ulUIParam,
	IntPtr lpMsgStore,
	IntPtr lpParentFolder,
	ref Guid lpInterface,
	uint ulMessageToken,
	IntPtr lpMessageSent,
	uint ulFlags,
	uint ulMessageStatus,
	uint ulMessageFlags,
	uint ulAccess,
	string lpszMessageClass
)
```
**VB**
``` VB
Function ShowForm ( 
	ulUIParam As UInteger,
	lpMsgStore As IntPtr,
	lpParentFolder As IntPtr,
	ByRef lpInterface As Guid,
	ulMessageToken As UInteger,
	lpMessageSent As IntPtr,
	ulFlags As UInteger,
	ulMessageStatus As UInteger,
	ulMessageFlags As UInteger,
	ulAccess As UInteger,
	lpszMessageClass As String
) As HRESULT
```
**C++**
``` C++
HRESULT ShowForm(
	unsigned int ulUIParam, 
	IntPtr lpMsgStore, 
	IntPtr lpParentFolder, 
	Guid% lpInterface, 
	unsigned int ulMessageToken, 
	IntPtr lpMessageSent, 
	unsigned int ulFlags, 
	unsigned int ulMessageStatus, 
	unsigned int ulMessageFlags, 
	unsigned int ulAccess, 
	String^ lpszMessageClass
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of the form.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the message store that contains the folder pointed to by the lpParentFolder parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the folder in which the message associated with the ulMessageToken parameter was created.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.guid" target="_blank" rel="noopener noreferrer">Guid</a></dt><dd>A pointer to the interface identifier (IID) that represents the interface to be used to access the message that is displayed in the form.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The token that is associated with the message to be displayed in the form.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Reserved; must be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how and whether the message is saved.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags copied from the PR_MSG_STATUS property of the message associated with the token in the ulMessageToken parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags copied from the PR_MESSAGE_FLAGS property of the message associated with the token in the ulMessageToken parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A flag that indicates the permission level for the message that is displayed in the form.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The message class of the message being displayed in the form.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the form was successfully displayed; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
