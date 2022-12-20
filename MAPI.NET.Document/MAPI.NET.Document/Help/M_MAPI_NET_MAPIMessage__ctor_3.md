# MAPIMessage(IMessage, UInt32, String) Constructor


Initializes a new instance of the MAPIMessage class.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public MAPIMessage(
	IMessage message,
	uint messageFlags,
	string messageClass
)
```
**VB**
``` VB
Public Sub New ( 
	message As IMessage,
	messageFlags As UInteger,
	messageClass As String
)
```
**C++**
``` C++
public:
MAPIMessage(
	IMessage^ message, 
	unsigned int messageFlags, 
	String^ messageClass
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_IMessage.md">IMessage</a></dt><dd>IMessage object</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.uint32" target="_blank" rel="noopener noreferrer">UInt32</a></dt><dd>MessageFlg property value of the IMessage object.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>MessageClass property value of the IMessage object.</dd></dl>

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="Overload_MAPI_NET_MAPIMessage__ctor.md">MAPIMessage Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
