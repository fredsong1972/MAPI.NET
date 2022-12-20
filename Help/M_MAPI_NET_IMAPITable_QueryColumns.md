# QueryColumns Method


Returns a list of columns for the table.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT QueryColumns(
	uint ulFlags,
	IntPtr lpPropTagArray
)
```
**VB**
``` VB
Function QueryColumns ( 
	ulFlags As UInteger,
	lpPropTagArray As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT QueryColumns(
	unsigned int ulFlags, 
	IntPtr lpPropTagArray
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Bitmask of flags that indicates which column set should be returned.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Pointer to an SPropTagArray structure containing the property tags for the column set.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the column set was successfully returned; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPITable.md">IMAPITable Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
