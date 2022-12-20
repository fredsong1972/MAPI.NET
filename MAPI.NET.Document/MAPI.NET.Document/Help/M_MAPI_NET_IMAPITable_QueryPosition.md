# QueryPosition Method


Retrieves the current table row position of the cursor, based on a fractional value.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT QueryPosition(
	IntPtr lpulRow,
	IntPtr lpulNumerator,
	IntPtr lpulDenominator
)
```
**VB**
``` VB
Function QueryPosition ( 
	lpulRow As IntPtr,
	lpulNumerator As IntPtr,
	lpulDenominator As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT QueryPosition(
	IntPtr lpulRow, 
	IntPtr lpulNumerator, 
	IntPtr lpulDenominator
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the number of the current row.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the numerator for the fraction identifying the table position.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to the denominator for the fraction identifying the table position.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the method returned valid values in lpulRow, lpulNumerator, and lpulDenominator; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
