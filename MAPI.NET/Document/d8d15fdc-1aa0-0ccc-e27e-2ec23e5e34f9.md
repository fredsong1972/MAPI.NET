# GetRowCount Method


Returns the total number of rows in the table.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

**C#**
``` C#
HRESULT GetRowCount(
	uint ulFlags,
	out uint lpulCount
)
```
**VB**
``` VB
Function GetRowCount ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lpulCount As UInteger
) As HRESULT
```
**C++**
``` C++
HRESULT GetRowCount(
	unsigned int ulFlags, 
	[OutAttribute] unsigned int% lpulCount
)
```
**F#**
``` F#
abstract GetRowCount : 
        ulFlags : uint32 * 
        lpulCount : uint32 byref -> HRESULT 
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Reserved; must be zero.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Pointer to the number of rows in the table.</dd></dl>

#### Return Value
<a href="50596607-a328-ef10-6ea9-0448fbb7d197.md">HRESULT</a>  
S_OK, if the row count was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="06a9b727-f5d6-e992-c936-a2712197dcee.md">IMAPITable Interface</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
