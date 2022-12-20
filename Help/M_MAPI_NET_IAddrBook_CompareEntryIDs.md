# CompareEntryIDs Method


Compares two entry identifiers that belong to a particular address book provider to determine whether they refer to the same address book object.



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
	out uint lpulResult
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
	<OutAttribute> ByRef lpulResult As UInteger
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
	[OutAttribute] unsigned int% lpulResult
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID1 parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the first entry identifier to be compared.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The byte count in the entry identifier pointed to by the lpEntryID2 parameter.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to the second entry identifier to be compared.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A pointer to the result of the comparison. The contents of lpulResult are set to TRUE if the two entry identifiers refer to the same object; otherwise, the contents are set to FALSE.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the call succeeded and has returned the expected value or values;otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IAddrBook.md">IAddrBook Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
