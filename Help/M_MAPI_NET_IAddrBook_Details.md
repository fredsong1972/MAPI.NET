# Details Method


Displays a dialog box that shows details about a particular address book entry.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Details(
	out uint lpulUIParam,
	IntPtr lpfnDismiss,
	IntPtr lpvDismissContext,
	uint cbEntryID,
	IntPtr lpEntryID,
	IntPtr lpfButtonCallback,
	IntPtr lpvButtonContext,
	string lpszButtonText,
	uint ulFlags
)
```
**VB**
``` VB
Function Details ( 
	<OutAttribute> ByRef lpulUIParam As UInteger,
	lpfnDismiss As IntPtr,
	lpvDismissContext As IntPtr,
	cbEntryID As UInteger,
	lpEntryID As IntPtr,
	lpfButtonCallback As IntPtr,
	lpvButtonContext As IntPtr,
	lpszButtonText As String,
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT Details(
	[OutAttribute] unsigned int% lpulUIParam, 
	IntPtr lpfnDismiss, 
	IntPtr lpvDismissContext, 
	unsigned int cbEntryID, 
	IntPtr lpEntryID, 
	IntPtr lpfButtonCallback, 
	IntPtr lpvButtonContext, 
	String^ lpszButtonText, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a handle of the parent window for the dialog box.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a function based on the DISMISSMODELESS prototype, or NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to context information to pass to the DISMISSMODELESS function pointed to by the lpfnDismiss parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier for the entry for which details are displayed.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a function based on the LPFNBUTTON function prototype</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to data that was used as a parameter for the function specified by the lpfButtonCallback parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>A pointer to a string that contains text to be applied to the added button, if that button is extensible.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the text for the lpszButtonText parameter.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the details dialog box was successfully displayed for the address book entry; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
