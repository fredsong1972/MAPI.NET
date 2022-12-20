# CompareEntryIDs Method


Compares two entry identifiers to determine whether they refer to the same entry in a message store.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT CompareEntryIDs(
	uint cbEntryID1,
	IntPtr lpEntryID1,
	uint cbEntryID2,
	IntPtr lpEntryID2,
	uint ulFlags,
	out bool lpulResult
)
```
**VB**
``` VB
Function CompareEntryIDs ( 
	cbEntryID1 As UInteger,
	lpEntryID1 As IntPtr,
	cbEntryID2 As UInteger,
	lpEntryID2 As IntPtr,
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulResult As Boolean
) As HRESULT
```
**C++**
``` C++
HRESULT CompareEntryIDs(
	unsigned int cbEntryID1, 
	IntPtr lpEntryID1, 
	unsigned int cbEntryID2, 
	IntPtr lpEntryID2, 
	unsigned int ulFlags, 
	[OutAttribute] bool% lpulResult
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID1 parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the first entry identifier to be compared.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID2 parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the second entry identifier to be compared.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a></dt><dd>true if the two entry identifiers refer to the same object; otherwise, false</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the comparison was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
