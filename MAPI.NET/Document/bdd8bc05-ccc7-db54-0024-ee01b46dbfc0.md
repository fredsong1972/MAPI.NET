# RegisterEvents Method


Registers to receive notification of specified events that affect the message store.



## Definition
**Namespace:** <a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET</a>  
**Assembly:** AllNetMAPI (in AllNetMAPI.dll) Version: 1.0.0.0 (1.0.0.0)

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
**F#**
``` F#
member RegisterEvents : 
        eventMask : EEventMask -> bool 
```



#### Parameters
<dl><dt>  <a href="5a15b17a-4117-5c7c-d72d-e89c6cb03fe4.md">EEventMask</a></dt><dd>A mask of values that indicate the types of notification events that the caller is interested in and should be included in the registration.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
\[Missing &lt;returns&gt; documentation for "M:MAPI.NET.MsgStore.RegisterEvents(MAPI.NET.EEventMask)"\]

## See Also


#### Reference
<a href="6f2a2863-4894-51bc-e286-04b5a90167ef.md">MsgStore Class</a>  
<a href="5bef4637-66f8-16d4-e5f4-4d0da57a1538.md">MAPI.NET Namespace</a>  
