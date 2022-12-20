# SetSender(String, String, MAPIAddressBook) Method


Set sender of the message



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool SetSender(
	string name,
	string emailAddress,
	MAPIAddressBook addressBook
)
```
**VB**
``` VB
Public Function SetSender ( 
	name As String,
	emailAddress As String,
	addressBook As MAPIAddressBook
) As Boolean
```
**C++**
``` C++
public:
bool SetSender(
	String^ name, 
	String^ emailAddress, 
	MAPIAddressBook^ addressBook
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>sender display name</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>sender email address</dd><dt>  <a href="T_MAPI_NET_MAPIAddressBook.md">MAPIAddressBook</a></dt><dd>address book object where the specified sender will be created one off.</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the sender was successfully set; otherwise, failed

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="Overload_MAPI_NET_MAPIMessage_SetSender.md">SetSender Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
