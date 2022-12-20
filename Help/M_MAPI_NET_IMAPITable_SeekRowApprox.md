# SeekRowApprox Method


Moves the cursor to an approximate fractional position in the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT SeekRowApprox(
	uint ulNumerator,
	uint ulDenominator
)
```
**VB**
``` VB
Function SeekRowApprox ( 
	ulNumerator As UInteger,
	ulDenominator As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT SeekRowApprox(
	unsigned int ulNumerator, 
	unsigned int ulDenominator
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The numerator of the fraction representing the table position</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The denominator of the fraction representing the table position</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the seek operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
