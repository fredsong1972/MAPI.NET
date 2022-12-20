# ResolveName Method


Performs name resolution, assigning entry identifiers to recipients in a recipient list.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT ResolveName(
	uint ulUIParam,
	uint ulFlags,
	string lpszNewEntryTitle,
	IntPtr lpAdrList
)
```
**VB**
``` VB
Function ResolveName ( 
	ulUIParam As UInteger,
	ulFlags As UInteger,
	lpszNewEntryTitle As String,
	lpAdrList As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT ResolveName(
	unsigned int ulUIParam, 
	unsigned int ulFlags, 
	String^ lpszNewEntryTitle, 
	IntPtr lpAdrList
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window of a dialog box that is shown, if specified, to prompt the user to resolve ambiguity.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that control various aspects of the resolution process.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>A pointer to text for the title of the control in the dialog box that prompts the user to enter a recipient.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an ADRLIST structure that contains the list of recipient names to be resolved.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the name resolution process succeeded; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
