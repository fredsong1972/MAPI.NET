# OpenEntry Method


Opens an object and returns an interface pointer for additional access.



## Definition
**Namespace:** <a href="N_MAPI_NET.md">MAPI.NET</a>  
**Assembly:** MAPI.NET (in MAPI.NET.dll) Version: 1.0.1

**C#**
``` C#
public IMAPIProp OpenEntry(
	EntryID id
)
```
**VB**
``` VB
Public Function OpenEntry ( 
	id As EntryID
) As IMAPIProp
```
**C++**
``` C++
public:
IMAPIProp^ OpenEntry(
	EntryID^ id
)
```



#### Parameters
<dl><dt>  <a href="T_MAPI_NET_EntryID.md">EntryID</a></dt><dd>The entry identifier of the object to open.</dd></dl>

#### Return Value
<a href="T_MAPI_NET_IMAPIProp.md">IMAPIProp</a>  
The IMAPIProp object which was opened.

## See Also


#### Reference
<a href="T_MAPI_NET_MAPISession.md">MAPISession Class</a>  
<a href="N_MAPI_NET.md">MAPI.NET Namespace</a>  
