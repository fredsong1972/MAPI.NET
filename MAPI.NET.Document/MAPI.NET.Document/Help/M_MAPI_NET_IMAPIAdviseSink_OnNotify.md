# OnNotify Method


The OnNotify method responds to a notification by performing one or more tasks, which depend on the object generating the notification, and type of event.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
HRESULT OnNotify(
	uint cNotify,
	IntPtr lpNotifications
)
```
**VB**
``` VB
Function OnNotify ( 
	cNotify As UInteger,
	lpNotifications As IntPtr
) As HRESULT
```
**C++**
``` C++
HRESULT OnNotify(
	unsigned int cNotify, 
	IntPtr lpNotifications
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>Ignored</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.intptr" target="_blank" rel="noopener noreferrer">IntPtr</a></dt><dd>Reference to one NOTIFICATION structure that provides information about the events that have occurred.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_HRESULT.md">HRESULT</a>  
S_OK, if the notification was processed successfully; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_IMAPIAdviseSink.md">IMAPIAdviseSink Interface</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
