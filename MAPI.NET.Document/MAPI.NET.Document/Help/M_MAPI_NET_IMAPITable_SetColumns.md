# SetColumns Method


Defines the particular properties and order of properties to appear as columns in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SetColumns(
	uint[] lpPropTagArray,
	uint ulFlags
)
```
**VB**
``` VB
Function SetColumns ( 
	lpPropTagArray As UInteger(),
	ulFlags As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT SetColumns(
	array<unsigned int>^ lpPropTagArray, 
	unsigned int ulFlags
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a>[]</dt><dd>Pointer to an array of property tags identifying properties to be included as columns in the table.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that controls the return of an asynchronous call to SetColumns.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the column setting operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
