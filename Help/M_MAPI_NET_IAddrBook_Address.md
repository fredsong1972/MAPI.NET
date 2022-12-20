# Address Method


Displays the Outlook address book dialog box.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT Address(
	out uint lpulUIParam,
	IntPtr lpAdrParms,
	out IntPtr lppAdrList
)
```
**VB**
``` VB
Function Address ( 
	<OutAttribute> ByRef lpulUIParam As UInteger,
	lpAdrParms As IntPtr,
	<OutAttribute> ByRef lppAdrList As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT Address(
	[OutAttribute] unsigned int% lpulUIParam, 
	IntPtr lpAdrParms, 
	[OutAttribute] IntPtr% lppAdrList
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to a handle of the parent window of the dialog box.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an ADRPARM structure that controls the presentation and behavior of the address dialog box.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>pointer to a pointer to an ADRLIST structure that contains recipient information.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the common address dialog box was successfully displayed; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
