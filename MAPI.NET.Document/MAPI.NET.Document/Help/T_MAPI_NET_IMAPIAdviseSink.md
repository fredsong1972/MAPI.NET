# IMAPIAdviseSink Interface


The IMAPIAdviseSink interface is used to implement an Advise Sink object for handling notifications.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
[GuidAttribute("00020302-0000-0000-C000-000000000046")]
public interface IMAPIAdviseSink
```
**VB**
``` VB
<ComImportAttribute>
<ComVisibleAttribute(false)>
<InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
<GuidAttribute("00020302-0000-0000-C000-000000000046")>
Public Interface IMAPIAdviseSink
```
**C++**
``` C++
[ComImportAttribute]
[ComVisibleAttribute(false)]
[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIUnknown)]
[GuidAttribute(L"00020302-0000-0000-C000-000000000046")]
public interface class IMAPIAdviseSink
```



## Methods
<table>
<tr>
<td><a href="M_MAPI_NET_IMAPIAdviseSink_OnNotify.md">OnNotify</a></td>
<td>The OnNotify method responds to a notification by performing one or more tasks, which depend on the object generating the notification, and type of event.</td></tr>
</table>

## See Also


#### Reference
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
