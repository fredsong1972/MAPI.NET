# RegisterEvents Method


Registers to receive notification of specified events that affect the message store.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool RegisterEvents(
	EEventMask eventMask
)
```
**VB**
``` VB
Public Function RegisterEvents ( 
	eventMask As EEventMask
) As Boolean
```
**C++**
``` C++
public:
bool RegisterEvents(
	EEventMask eventMask
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_EEventMask.md">EEventMask</a></dt><dd>A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if successful; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MsgStore.md">MsgStore Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
