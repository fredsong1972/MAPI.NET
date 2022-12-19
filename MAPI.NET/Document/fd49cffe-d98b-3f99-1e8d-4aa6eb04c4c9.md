# SeekRowApprox Method


Moves the cursor to an approximate fractional position in the table.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
abstract SeekRowApprox : 
        ulNumerator : uint32 * 
        ulDenominator : uint32 -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The numerator of the fraction representing the table position</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>The denominator of the fraction representing the table position</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the seek operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="06a9b727-f5d6-e992-c936-a2712197dcee.md">IMAPITable Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
