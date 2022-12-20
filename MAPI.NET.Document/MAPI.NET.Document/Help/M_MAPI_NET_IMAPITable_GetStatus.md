# GetStatus Method


Returns the table's status and type.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetStatus(
	IntPtr lpulTableStatus,
	IntPtr lpulTableType
)
```
**VB**
``` VB
Function GetStatus ( 
	lpulTableStatus As IntPtr,
	lpulTableType As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetStatus(
	IntPtr lpulTableStatus, 
	IntPtr lpulTableType
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a value indicating the status of the table.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to a value that indicates the table's type.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the table's status was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
