# OnObjectCreated Event


Object created event



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public event EventHandler<MsgStoreObjectEventArgs> OnObjectCreated
```
**VB**
``` VB
Public Event OnObjectCreated As EventHandler(Of MsgStoreObjectEventArgs)
```
**C++**
``` C++
public:
 event EventHandler<MsgStoreObjectEventArgs^>^ OnObjectCreated {
	void add (EventHandler<MsgStoreObjectEventArgs^>^ value);
	void remove (EventHandler<MsgStoreObjectEventArgs^>^ value);
}
```



#### Value
<a href="https://learn.microsoft.com/dotnet/api/system.eventhandler-1" target="_blank" rel="noopener noreferrer">EventHandler</a>(<a href="T_MAPI_NET_MsgStoreObjectEventArgs.md">MsgStoreObjectEventArgs</a>)

## See Also


#### Reference
<a href="T_MAPI_NET_MsgStore.md">MsgStore Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
