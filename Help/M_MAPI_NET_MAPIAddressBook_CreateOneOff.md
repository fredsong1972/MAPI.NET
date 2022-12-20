# CreateOneOff Method


Creates an entry identifier for a one-off address.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public EntryID CreateOneOff(
	string name,
	RecipientType recipientType,
	string emailAddress
)
```
**VB**
``` VB
Public Function CreateOneOff ( 
	name As String,
	recipientType As RecipientType,
	emailAddress As String
) As EntryID
```
**C++**
``` C++
public:
EntryID^ CreateOneOff(
	String^ name, 
	RecipientType recipientType, 
	String^ emailAddress
)
```



#### Parameters
<dl><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The recipient’s display name.</dd><dt>  <a href="T_MAPI_NET_RecipientType.md">RecipientType</a></dt><dd>The address type of the recipient.</dd><dt>  <a href="https://learn.microsoft.com/dotnet/api/system.string" target="_blank" rel="noopener noreferrer">String</a></dt><dd>The address of the recipient</dd></dl>

#### Return Value
<a href="T_MAPI_NET_EntryID.md">EntryID</a>  
true if successful; otherwise, false

## See Also


#### Reference
<a href="T_MAPI_NET_MAPIAddressBook.md">MAPIAddressBook Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
