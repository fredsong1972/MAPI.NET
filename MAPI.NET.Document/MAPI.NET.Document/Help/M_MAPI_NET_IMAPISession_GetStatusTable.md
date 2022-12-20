# GetStatusTable Method


Provides access to the status table, a table that contains information about all the MAPI resources in the session.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT GetStatusTable(
	uint ulFlags,
	out IntPtr lppTable
)
```
**VB**
``` VB
Function GetStatusTable ( 
	ulFlags As UInteger,
	<OutAttribute> ByRef lppTable As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT GetStatusTable(
	unsigned int ulFlags, 
	[OutAttribute] IntPtr% lppTable
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>A bitmask of flags that determines the format for columns that are character strings.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>A pointer to a pointer to the status table.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the table was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPISession.md">IMAPISession Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
