# PrepareRecips Method


Prepares a recipient list for later use by the messaging system.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT PrepareRecips(
	uint ulFlags,
	IntPtr lpSPropTagArray,
	IntPtr lpRecipList
)
```
**VB**
``` VB
Function PrepareRecips ( 
	ulFlags As UInteger,
	lpSPropTagArray As IntPtr,
	lpRecipList As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT PrepareRecips(
	unsigned int ulFlags, 
	IntPtr lpSPropTagArray, 
	IntPtr lpRecipList
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that controls how the entry is opened.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an SPropTagArray structure that contains an array of property tags that indicate the properties, if any, that require updating. The lpSPropTagArray parameter can be NULL.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an ADRLIST structure that contains the list of recipients.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the recipient list was successfully prepared; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
