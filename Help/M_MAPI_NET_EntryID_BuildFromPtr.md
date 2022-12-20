# BuildFromPtr Method


Build EntryID from a unmanaged memory block.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public static EntryID BuildFromPtr(
	uint cb,
	IntPtr lpb
)
```
**VB**
``` VB
Public Shared Function BuildFromPtr ( 
	cb As UInteger,
	lpb As IntPtr
) As EntryID
```
**C++**
``` C++
public:
static EntryID^ BuildFromPtr(
	unsigned int cb, 
	IntPtr lpb
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>the size of block</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>the pointer of the unmanaged memory block</dd></dl>

#### Return Value
<a href="T_MAPI_NET_EntryID.md">EntryID</a>  
EntryID object

## See Also


#### Reference
<a href="T_MAPI_NET_EntryID.md">EntryID Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
