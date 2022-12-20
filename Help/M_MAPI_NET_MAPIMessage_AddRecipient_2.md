# AddRecipient(String, RecipientType, AddressType, MAPIAddressBook) Method


Add a recipient



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public bool AddRecipient(
	string strEmail,
	RecipientType recipientType,
	AddressType addressType,
	MAPIAddressBook addressBook
)
```
**VB**
``` VB
Public Function AddRecipient ( 
	strEmail As String,
	recipientType As RecipientType,
	addressType As AddressType,
	addressBook As MAPIAddressBook
) As Boolean
```
**C++**
``` C++
public:
bool AddRecipient(
	String^ strEmail, 
	RecipientType recipientType, 
	AddressType addressType, 
	MAPIAddressBook^ addressBook
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>email address of the recipient</dd><dt>  <a href="T_MAPI_NET_RecipientType.md">RecipientType</a></dt><dd>type of the recipient</dd><dt>  <a href="T_MAPI_NET_AddressType.md">AddressType</a></dt><dd>address type of the recipient</dd><dt>  <a href="T_MAPI_NET_MAPIAddressBook.md">MAPIAddressBook</a></dt><dd>address book object which resolves the recipient</dd></dl>

#### Return Value
<a href="https://learn.microsoft.com/dotnet/api/system.boolean" target="_blank" rel="noopener noreferrer">Boolean</a>  
true, if the recipient was successfully added; otherwise, failed.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIMessage.md">MAPIMessage Class</a>  
<a href="Overload_MAPI_NET_MAPIMessage_AddRecipient.md">AddRecipient Overload</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
