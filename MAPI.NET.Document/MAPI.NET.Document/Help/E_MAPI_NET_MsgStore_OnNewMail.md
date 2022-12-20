# OnNewMail Event


New mail event



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public event EventHandler<MsgStoreNewMailEventArgs> OnNewMail
```
**VB**
``` VB
Public Event OnNewMail As EventHandler(Of MsgStoreNewMailEventArgs)
```
**C++**
``` C++
public:
 event EventHandler<MsgStoreNewMailEventArgs^>^ OnNewMail {
	void add (EventHandler<MsgStoreNewMailEventArgs^>^ value);
	void remove (EventHandler<MsgStoreNewMailEventArgs^>^ value);
}
```



#### Value
<a href="https://learn.microsoft.com/dotnet/api/system.eventhandler-1" target="_blank" rel="noopener noreferrer">EventHandler</a>(<a href="T_MAPI_NET_MsgStoreNewMailEventArgs.md">MsgStoreNewMailEventArgs</a>)

## See Also


#### Reference
<a href="T_MAPI_NET_MsgStore.md">MsgStore Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
