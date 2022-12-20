# ResolveName Method


Performs name resolution, assigning entry identifiers to recipients in a recipient list.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool ResolveName(
	IntPtr pAdrList
)
```
**VB**
``` VB
Public Function ResolveName ( 
	pAdrList As IntPtr
) As Boolean
```
**C++**
``` C++
public:
bool ResolveName(
	IntPtr pAdrList
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to an ADRLIST structure that contains the list of recipient names to be resolved.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true if successful; otherwise, false

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIAddressBook.md">MAPIAddressBook Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
