# FlushQueues Method


Forces all messages waiting to be sent or received to be immediately uploaded or downloaded.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool FlushQueues(
	FlushFlag flushFlags
)
```
**VB**
``` VB
Public Function FlushQueues ( 
	flushFlags As FlushFlag
) As Boolean
```
**C++**
``` C++
public:
bool FlushQueues(
	FlushFlag flushFlags
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_FlushFlag.md">FlushFlag</a></dt><dd>A bitmask of flags that controls the flush operation.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the flush operation was successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIStatus.md">MAPIStatus Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
