# NewEntry Method


Adds a new recipient to an address book container or to the recipient list of an outgoing message.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT NewEntry(
	uint ulUIParam,
	uint ulFlags,
	uint cbEIDContainer,
	IntPtr lpEIDContainer,
	uint cbEIDNewEntryTpl,
	IntPtr lpEIDNewEntryTpl,
	out uint lpcbEIDNewEntry,
	out IntPtr lppEIDNewEntry
)
```
**VB**
``` VB
Function NewEntry ( 
	ulUIParam As UInteger,
	ulFlags As UInteger,
	cbEIDContainer As UInteger,
	lpEIDContainer As IntPtr,
	cbEIDNewEntryTpl As UInteger,
	lpEIDNewEntryTpl As IntPtr,
	<OutAttribute> ByRef lpcbEIDNewEntry As UInteger,
	<OutAttribute> ByRef lppEIDNewEntry As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT NewEntry(
	unsigned int ulUIParam, 
	unsigned int ulFlags, 
	unsigned int cbEIDContainer, 
	IntPtr lpEIDContainer, 
	unsigned int cbEIDNewEntryTpl, 
	IntPtr lpEIDNewEntryTpl, 
	[OutAttribute] unsigned int% lpcbEIDNewEntry, 
	[OutAttribute] IntPtr% lppEIDNewEntry
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A handle to the parent window for the dialog box.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls the type of the text that is used.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEIDContainer parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the entry identifier of the container where the new recipient is to be added.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEIDNewEntryTpl parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a one-off template that will be used to create the new recipient.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the byte count in the entry identifier pointed to by the lppEIDNewEntry parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the new recipient's entry identifier.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the new address book entry was successfully created; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
